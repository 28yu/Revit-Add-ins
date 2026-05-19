using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    internal static class ElementCollector
    {
        /// <summary>
        /// 型枠不要として除外された要素 1 件分の情報。
        /// </summary>
        internal class ExcludedEntry
        {
            public Element Element;
            public ElementSource Source;
            public ExclusionKind Kind;
            public string Layer = string.Empty;
            public string Reason = string.Empty;
        }

        /// <summary>
        /// 要素収集結果。型枠算出対象の Targets と、除外された Excluded の 2 リスト。
        /// 各要素は ElementSource にラップされており、ホスト/リンクの区別と
        /// 配置トランスフォームを保持する。
        /// </summary>
        internal class CollectionResult
        {
            public List<ElementSource> Targets = new List<ElementSource>();
            public List<ExcludedEntry> Excluded = new List<ExcludedEntry>();
            public ElementSourceRegistry Registry = new ElementSourceRegistry();
            public int LinkedInstanceCount;
            public int LinkedDocumentCount;
        }

        private static readonly Dictionary<string, BuiltInCategory> _nameToCat
            = new Dictionary<string, BuiltInCategory>
            {
                { "StructuralColumns", BuiltInCategory.OST_StructuralColumns },
                { "StructuralFraming", BuiltInCategory.OST_StructuralFraming },
                { "Walls", BuiltInCategory.OST_Walls },
                { "Floors", BuiltInCategory.OST_Floors },
                { "StructuralFoundation", BuiltInCategory.OST_StructuralFoundation },
                { "Stairs", BuiltInCategory.OST_Stairs },
                { "Roofs", BuiltInCategory.OST_Roofs },
            };

        internal static CollectionResult CollectAndClassify(
            Document doc, FormworkSettings settings, View activeView)
        {
            var cr = new CollectionResult();

            // 1) ホストドキュメントの要素を収集して登録
            var hostRaw = CollectFromDoc(doc, settings, activeView, isLinked: false);
            foreach (var elem in hostRaw)
            {
                cr.Registry.RegisterHost(elem, doc);
            }

            // 2) リンクモデルの要素を収集して登録
            if (settings != null && settings.IncludeLinkedModels)
            {
                CollectFromLinkedModels(doc, settings, activeView, cr);
            }

            ClassifyAndFilter(doc, cr, settings);
            return cr;
        }

        /// <summary>
        /// リンクモデルから対象要素を収集して登録する。
        /// 「現在のビュー」モードでは、リンクインスタンスがアクティブビューに表示されている
        /// もののみ対象とする (リンク内の要素単位の可視判定は行わない簡易版)。
        ///
        /// BIM360 ワークシェアリングモデル対応:
        ///   リンクドキュメントがワークシェアリング (IsWorkshared=true) の場合、
        ///   全要素を走査すると `get_Geometry()` 等がクラウド同期を待機してフリーズすることがある。
        ///   対策として CurrentView モードでは 3D ビューのセクションボックスを使い、
        ///   リンクモデル座標系に逆変換した BoundingBoxIntersectsFilter で要素数を絞り込む。
        /// </summary>
        private static void CollectFromLinkedModels(
            Document hostDoc, FormworkSettings settings, View activeView, CollectionResult cr)
        {
            var linkInstances = new FilteredElementCollector(hostDoc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            FormworkDebugLog.Log($"  [LinkedCollect] linkInstances found: {linkInstances.Count}");

            int linkedDocCount = 0;
            int instanceCount = 0;
            foreach (var rli in linkInstances)
            {
                Document linkDoc = null;
                try { linkDoc = rli.GetLinkDocument(); } catch { }
                if (linkDoc == null) continue;

                // 「現在のビュー」モードではリンクインスタンスがアクティブビューに表示されているか確認
                if (settings.Scope == CalculationScope.CurrentView && activeView != null)
                {
                    string rliName = MakeLinkSourceName(rli, linkDoc);

                    // (1) 個別非表示・カテゴリ非表示チェック
                    bool hidden = true;
                    try { hidden = rli.IsHidden(activeView); } catch { hidden = true; }

                    // (2) ワークセット経由の非表示チェック
                    // IsHidden() はワークセット非表示を検出しない。
                    // RevitLinkInstance はホストドキュメントの要素なので、ホストビューの
                    // GetWorksetVisibility() で判定できる (クロスドキュメント問題なし)。
                    if (!hidden)
                    {
                        try
                        {
                            if (hostDoc.IsWorkshared)
                            {
                                WorksetVisibility wsVis = activeView.GetWorksetVisibility(rli.WorksetId);
                                if (wsVis == WorksetVisibility.Hidden)
                                {
                                    hidden = true;
                                    FormworkDebugLog.Log(
                                        $"  [LinkInstVis] {rliName}: " +
                                        $"ワークセット非表示 wsId={rli.WorksetId.IntegerValue} → スキップ");
                                }
                                else if (wsVis == WorksetVisibility.UseGlobalSetting)
                                {
                                    var wdvs = WorksetDefaultVisibilitySettings
                                        .GetWorksetDefaultVisibilitySettings(hostDoc);
                                    if (wdvs != null && !wdvs.IsWorksetVisible(rli.WorksetId))
                                    {
                                        hidden = true;
                                        FormworkDebugLog.Log(
                                            $"  [LinkInstVis] {rliName}: " +
                                            $"ワークセット既定非表示 wsId={rli.WorksetId.IntegerValue} → スキップ");
                                    }
                                    else if (FormworkDebugLog.Enabled)
                                    {
                                        FormworkDebugLog.Log(
                                            $"  [LinkInstVis] {rliName}: " +
                                            $"wsId={rli.WorksetId.IntegerValue} UseGlobalSetting→visible");
                                    }
                                }
                                else if (FormworkDebugLog.Enabled)
                                {
                                    FormworkDebugLog.Log(
                                        $"  [LinkInstVis] {rliName}: " +
                                        $"wsId={rli.WorksetId.IntegerValue} wsVis={wsVis} → visible");
                                }
                            }
                        }
                        catch { }
                    }
                    else if (FormworkDebugLog.Enabled)
                    {
                        FormworkDebugLog.Log($"  [LinkInstVis] {rliName}: IsHidden=true → スキップ");
                    }

                    if (hidden) continue;
                }

                Transform transform = null;
                try { transform = rli.GetTotalTransform(); } catch { }
                if (transform == null) transform = Transform.Identity;

                string sourceName = MakeLinkSourceName(rli, linkDoc);

                // ワークシェアリング診断ログ
                bool isWorkshared = false;
                try { isWorkshared = linkDoc.IsWorkshared; } catch { }
                if (isWorkshared)
                {
                    FormworkDebugLog.Log(
                        $"  [LinkedCollect] ⚠️ {sourceName}: IsWorkshared=true (BIM360 or Revit Server). " +
                        $"セクションボックスフィルタで要素数を絞り込みます。");
                }

                // CurrentView + 3D セクションボックスがあればリンクローカル座標系に変換した
                // BoundingBoxIntersectsFilter を生成し、大規模モデルの要素数を絞り込む。
                Outline linkLocalOutline = null;
                if (settings.Scope == CalculationScope.CurrentView)
                {
                    linkLocalOutline = GetLinkLocalSectionBoxOutline(activeView, transform);
                    if (linkLocalOutline != null)
                    {
                        FormworkDebugLog.Log(
                            $"  [LinkedCollect] {sourceName}: セクションボックスフィルタ適用 " +
                            $"link-local=[({linkLocalOutline.MinimumPoint.X:F2},{linkLocalOutline.MinimumPoint.Y:F2},{linkLocalOutline.MinimumPoint.Z:F2}) - " +
                            $"({linkLocalOutline.MaximumPoint.X:F2},{linkLocalOutline.MaximumPoint.Y:F2},{linkLocalOutline.MaximumPoint.Z:F2})]");
                    }
                }

                // フリーズ診断: CollectFromDoc 呼び出し前にフラッシュ。
                // BIM360 ワークシェアリング doc では FilteredElementCollector.ToList() が
                // クラウド同期待ちでここで止まる場合があるため、呼び出し直前ログを残す。
                FormworkDebugLog.Log(
                    $"  [LinkedCollect] {sourceName}: CollectFromDoc 開始" +
                    $" cats={settings.IncludedCategories.Count}" +
                    $" outline={(linkLocalOutline != null ? "有" : "無(全要素対象)")}");
                FormworkDebugLog.Flush();

                var linkedElems = CollectFromDoc(linkDoc, settings, null, isLinked: true, linkLocalOutline: linkLocalOutline);

                if (activeView != null && settings.Scope == CalculationScope.CurrentView)
                {
                    // (Z) リンクの表示モードが「リンクされたビューで設定」(ByLinkedView) の場合、
                    // 割当られたリンクビューで実際に可視な要素のみに絞り込む。
                    // ホストビューの V/G → Revit リンク → 「表示設定」が「リンクされたビューで設定」
                    // になっている場合、その割当ビューの設定 (セクションボックス・ワークセット・
                    // フィルタ・要素個別非表示) に従って表示される。エンジンは従来これを考慮しておらず、
                    // 全リンク要素を対象としていたが、ソースビューが特定エリアを切り出しているケースに
                    // 対応するため、リンクビューでの可視要素のみを残す。
                    try
                    {
                        var linkOv = activeView.GetLinkOverrides(rli.Id);
                        if (linkOv != null &&
                            linkOv.LinkVisibilityType == LinkVisibility.ByLinkedView)
                        {
                            var lvId = linkOv.LinkedViewId;
                            if (lvId != null && lvId != ElementId.InvalidElementId)
                            {
                                var lvVisible = new HashSet<int>(
                                    new FilteredElementCollector(linkDoc, lvId)
                                        .WhereElementIsNotElementType()
                                        .ToElementIds()
                                        .Select(eid => eid.IntegerValue));
                                int beforeLV = linkedElems.Count;
                                linkedElems = linkedElems
                                    .Where(e => lvVisible.Contains(e.Id.IntegerValue))
                                    .ToList();
                                int hiddenLV = beforeLV - linkedElems.Count;
                                FormworkDebugLog.Log(
                                    $"  [VisFilter] {sourceName}: リンクビュー({lvId.IntValue()})フィルタ除外 " +
                                    $"{hiddenLV}/{beforeLV} (ByLinkedView mode)");
                            }
                            else
                            {
                                FormworkDebugLog.Log(
                                    $"  [VisFilter] {sourceName}: ByLinkedView だが LinkedViewId 未設定");
                            }
                        }
                        else if (linkOv != null && FormworkDebugLog.Enabled)
                        {
                            FormworkDebugLog.Log(
                                $"  [VisFilter] {sourceName}: LinkVisibilityType={linkOv.LinkVisibilityType} (ByLinkedView 以外、絞り込みなし)");
                        }
                    }
                    catch (Exception exLv)
                    {
                        FormworkDebugLog.Log(
                            $"  [VisFilter] {sourceName}: GetLinkOverrides EX: {exLv.Message}");
                    }

                    // (A) グローバルカテゴリ非表示: V/G → モデルカテゴリ タブ
                    int beforeCat = linkedElems.Count;
                    linkedElems = linkedElems.Where(e =>
                    {
                        var cat = e?.Category;
                        if (cat == null) return true;
                        try { return !activeView.GetCategoryHidden(cat.Id); }
                        catch { return true; }
                    }).ToList();
                    int hiddenCat = beforeCat - linkedElems.Count;
                    FormworkDebugLog.Log(
                        $"  [VisFilter] {sourceName}: カテゴリ非表示除外 {hiddenCat}/{beforeCat}");

                    // (B) ワークセット非表示: V/G → ワークセット / Revitリンク → ワークセット
                    if (isWorkshared)
                    {
                        var hiddenWorksets = BuildHiddenWorksetIds(linkDoc, activeView, sourceName);
                        if (hiddenWorksets.Count > 0)
                        {
                            int beforeWs = linkedElems.Count;
                            linkedElems = linkedElems.Where(e =>
                            {
                                try { return !hiddenWorksets.Contains(e.WorksetId.IntegerValue); }
                                catch { return true; }
                            }).ToList();
                            int hiddenWs = beforeWs - linkedElems.Count;
                            FormworkDebugLog.Log(
                                $"  [VisFilter] {sourceName}: ワークセット非表示除外 {hiddenWs}/{beforeWs}");
                        }
                    }

                    // (C) IsHidden チェック (Revitバージョンによってはリンク要素にも有効)
                    int beforeIH = linkedElems.Count;
                    linkedElems = linkedElems.Where(e =>
                    {
                        try { return !e.IsHidden(activeView); }
                        catch { return true; }
                    }).ToList();
                    int hiddenIH = beforeIH - linkedElems.Count;
                    if (hiddenIH > 0)
                        FormworkDebugLog.Log(
                            $"  [VisFilter] {sourceName}: IsHidden 除外 {hiddenIH}/{beforeIH}");

                    // (D) 診断ログ: 残った要素のサンプル (先頭5件) のワークセット情報
                    if (linkedElems.Count > 0 && isWorkshared)
                    {
                        var samples = linkedElems.Take(5).ToList();
                        foreach (var s in samples)
                        {
                            try
                            {
                                int wsInt = s.WorksetId.IntegerValue;
                                Workset ws = null;
                                try { ws = linkDoc.GetWorksetTable().GetWorkset(s.WorksetId); } catch { }
                                FormworkDebugLog.Log(
                                    $"  [VisFilter:sample] {sourceName} E{s.Id.IntValue()} cat='{s.Category?.Name}' " +
                                    $"workset=[{wsInt}'{ws?.Name ?? "?"}']");
                            }
                            catch { }
                        }
                    }
                }

                foreach (var elem in linkedElems)
                {
                    cr.Registry.RegisterLinked(elem, linkDoc, transform, sourceName, rli.Id);
                }
                instanceCount++;
                linkedDocCount++;
                FormworkDebugLog.Log(
                    $"  [LinkedCollect] {sourceName} elements={linkedElems.Count} " +
                    $"isWorkshared={isWorkshared} " +
                    $"transform=Translation({transform.Origin.X:F2},{transform.Origin.Y:F2},{transform.Origin.Z:F2})");
            }

            cr.LinkedInstanceCount = instanceCount;
            cr.LinkedDocumentCount = linkedDocCount;
        }

        /// <summary>
        /// リンク文書のユーザーワークセットのうち、リンク文書の既定値で非表示とみなされる
        /// ワークセット ID の集合を返す。
        ///
        /// 注意: hostView.GetWorksetVisibility(ws.Id) はホストモデルのワークセット ID で
        /// 動作する API であり、リンクモデルのワークセット ID をそのまま渡すと、
        /// たまたま同じ整数値を持つホスト側ワークセットの状態を返してしまう。
        /// 例: ホストとリンク両方に id=0 の「ワークセット1」があり、ホストビューで
        ///     ホスト側を非表示にするとリンク要素まで除外される誤動作が発生する。
        /// このためリンクモデルの可視性はリンク文書自身の既定値のみで判断する。
        /// </summary>
        private static HashSet<int> BuildHiddenWorksetIds(Document linkDoc, View hostView, string sourceName)
        {
            var hidden = new HashSet<int>();
            if (linkDoc == null || !linkDoc.IsWorkshared) return hidden;

            WorksetDefaultVisibilitySettings wdvs = null;
            try { wdvs = WorksetDefaultVisibilitySettings.GetWorksetDefaultVisibilitySettings(linkDoc); }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [WsCheck] {sourceName}: WorksetDefaultVisibilitySettings 取得失敗: {ex.Message}");
            }

            int total = 0;
            int defaultHidden = 0;
            try
            {
                var wsCol = new FilteredWorksetCollector(linkDoc)
                    .OfKind(WorksetKind.UserWorkset);

                foreach (Workset ws in wsCol)
                {
                    total++;
                    bool isHidden = false;
                    string reason = "visible";

                    // リンク文書の既定値で非表示かチェック。
                    // hostView.GetWorksetVisibility() はホストモデルの API であり
                    // リンクモデルのワークセット ID には使用しない (誤判定の原因となるため)。
                    if (wdvs != null)
                    {
                        try
                        {
                            if (!wdvs.IsWorksetVisible(ws.Id))
                            {
                                isHidden = true;
                                reason = "linkDoc.default.Hidden";
                                defaultHidden++;
                            }
                        }
                        catch { }
                    }

                    FormworkDebugLog.Log(
                        $"  [WsCheck] {sourceName} ws='{ws.Name}' id={ws.Id.IntegerValue} → {reason}");

                    if (isHidden) hidden.Add(ws.Id.IntegerValue);
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [WsCheck] {sourceName}: ワークセット列挙失敗: {ex.Message}");
            }

            FormworkDebugLog.Log(
                $"  [WsCheck] {sourceName}: ワークセット総計={total} " +
                $"既定非表示={defaultHidden} → 除外対象={hidden.Count}");

            return hidden;
        }

        /// <summary>
        /// アクティブビュー (3D) のセクションボックスをリンクモデルのローカル座標系に変換した
        /// Outline を返す。セクションボックスが無効 / 取得失敗の場合は null を返す。
        /// </summary>
        private static Outline GetLinkLocalSectionBoxOutline(View activeView, Transform linkTransform)
        {
            if (!(activeView is View3D v3d)) return null;

            BoundingBoxXYZ sb = null;
            try
            {
                if (v3d.IsSectionBoxActive)
                    sb = v3d.GetSectionBox();
            }
            catch { }
            if (sb == null) return null;

            Transform invLink = null;
            try { invLink = linkTransform.Inverse; } catch { return null; }

            // セクションボックスは独自のローカル Transform を持つため、
            // 8 コーナーを sbTransform でワールド座標 → invLink でリンクローカル座標 に変換する
            Transform sbTrans = sb.Transform ?? Transform.Identity;

            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

            bool valid = false;
            foreach (double cx in new[] { sb.Min.X, sb.Max.X })
            foreach (double cy in new[] { sb.Min.Y, sb.Max.Y })
            foreach (double cz in new[] { sb.Min.Z, sb.Max.Z })
            {
                XYZ world;
                try { world = sbTrans.OfPoint(new XYZ(cx, cy, cz)); } catch { return null; }
                XYZ local;
                try { local = invLink.OfPoint(world); } catch { return null; }
                minX = Math.Min(minX, local.X);
                minY = Math.Min(minY, local.Y);
                minZ = Math.Min(minZ, local.Z);
                maxX = Math.Max(maxX, local.X);
                maxY = Math.Max(maxY, local.Y);
                maxZ = Math.Max(maxZ, local.Z);
                valid = true;
            }
            if (!valid) return null;

            // 境界要素が欠けないよう 500mm のマージンを追加
            double tol = UnitUtils.ConvertToInternalUnits(500, UnitTypeId.Millimeters);
            try
            {
                return new Outline(
                    new XYZ(minX - tol, minY - tol, minZ - tol),
                    new XYZ(maxX + tol, maxY + tol, maxZ + tol));
            }
            catch { return null; }
        }

        /// <summary>
        /// リンクのソース名を生成する。リンクインスタンス名 (シンボル名 + コピー番号) を優先し、
        /// 取得できない場合はリンクファイルのタイトルを使う。
        /// </summary>
        private static string MakeLinkSourceName(RevitLinkInstance rli, Document linkDoc)
        {
            string baseName = string.Empty;
            try
            {
                baseName = Path.GetFileNameWithoutExtension(linkDoc?.Title ?? string.Empty);
            }
            catch { }
            if (string.IsNullOrEmpty(baseName))
            {
                try { baseName = rli.Name ?? string.Empty; } catch { }
            }
            return string.IsNullOrEmpty(baseName) ? "リンク" : baseName;
        }

        /// <summary>
        /// 鉄骨・デッキスラブ・鉄骨階段・LGS壁等の除外判定を行い、
        /// cr.Registry に登録された全要素を Targets と Excluded に振り分ける。
        /// </summary>
        private static void ClassifyAndFilter(Document hostDoc, CollectionResult cr, FormworkSettings settings)
        {
            bool exclude = settings?.ExcludeSteelMembers ?? true;
            bool excludeStair = settings?.ExcludeSteelStairs ?? true;
            bool excludeLgs = settings?.ExcludeLgsWalls ?? true;

            FormworkDebugLog.Section("Exclusion Detection (Steel + DeckSlab + SteelStair + LgsWall)");
            int steelCount = 0;
            int deckCount = 0;
            int steelStairCount = 0;
            int lgsCount = 0;
            int totalCount = 0;

            foreach (var src in cr.Registry.All())
            {
                totalCount++;
                var elem = src.Element;
                var srcDoc = src.SourceDoc;
                string srcTag = src.IsLinked ? $"[link:{src.SourceName}]" : "[host]";

                if (exclude)
                {
                    // 構造柱・構造フレーム → 鉄骨判定
                    // リンク要素は skipShapeAnalysis=true: L2 の get_Geometry を除外フェーズで呼ばない。
                    // BIM360 ワークシェアリングモデルでは get_Geometry がクラウド待ちでブロックする
                    // 可能性があるため。L2 で見逃した鉄骨は Pass 1 (ClassifyElementFaces) の
                    // get_Geometry 呼び出しで solid が 0 になるため、自然に除外される。
                    if (IsSteelDetectionTarget(elem))
                    {
                        var det = SteelMemberDetector.Detect(elem, srcDoc, skipShapeAnalysis: src.IsLinked);
                        if (det != null && det.IsSteel)
                        {
                            cr.Excluded.Add(new ExcludedEntry
                            {
                                Element = elem,
                                Source = src,
                                Kind = ExclusionKind.Steel,
                                Layer = det.Layer.ToString(),
                                Reason = det.Reason ?? string.Empty,
                            });
                            FormworkDebugLog.Log(
                                $"  [SteelExclude] {srcTag} E{elem.Id.IntValue()} " +
                                $"Cat={elem.Category?.Name} Name='{elem.Name}' " +
                                $"L={det.Layer} reason={det.Reason}");
                            steelCount++;
                            continue;
                        }
                    }

                    // 階段 → 鉄骨階段判定
                    if (excludeStair && IsStairDetectionTarget(elem))
                    {
                        var stairRes = SteelStairDetector.Detect(elem, srcDoc);
                        if (stairRes != null && stairRes.IsSteel)
                        {
                            cr.Excluded.Add(new ExcludedEntry
                            {
                                Element = elem,
                                Source = src,
                                Kind = ExclusionKind.SteelStair,
                                Layer = stairRes.Layer,
                                Reason = stairRes.Reason ?? string.Empty,
                            });
                            FormworkDebugLog.Log(
                                $"  [SteelStairExclude] {srcTag} E{elem.Id.IntValue()} " +
                                $"reason={stairRes.Reason}");
                            steelStairCount++;
                            continue;
                        }
                    }

                    // 壁 → LGS壁判定
                    if (excludeLgs && IsWallDetectionTarget(elem))
                    {
                        if (LgsWallDetector.IsLgsWall(elem, srcDoc, out string lgsReason))
                        {
                            cr.Excluded.Add(new ExcludedEntry
                            {
                                Element = elem,
                                Source = src,
                                Kind = ExclusionKind.LgsWall,
                                Layer = "CompoundStructure",
                                Reason = lgsReason,
                            });
                            FormworkDebugLog.Log(
                                $"  [LgsExclude] {srcTag} E{elem.Id.IntValue()} reason={lgsReason}");
                            lgsCount++;
                            continue;
                        }
                    }

                    // 床 → デッキスラブ判定
                    if (IsDeckSlabDetectionTarget(elem))
                    {
                        if (DeckSlabDetector.IsDeckSlab(elem, srcDoc, out string deckReason))
                        {
                            cr.Excluded.Add(new ExcludedEntry
                            {
                                Element = elem,
                                Source = src,
                                Kind = ExclusionKind.DeckSlab,
                                Layer = "NamePattern",
                                Reason = deckReason,
                            });
                            FormworkDebugLog.Log(
                                $"  [DeckSlabExclude] {srcTag} E{elem.Id.IntValue()} reason={deckReason}");
                            deckCount++;
                            continue;
                        }
                    }
                }

                cr.Targets.Add(src);
            }
            FormworkDebugLog.Log(
                $"  Exclusion detection: total={totalCount} steelExcluded={steelCount} " +
                $"deckSlabExcluded={deckCount} steelStairExcluded={steelStairCount} " +
                $"lgsExcluded={lgsCount} kept={cr.Targets.Count} " +
                $"linkedInstances={cr.LinkedInstanceCount}");
            FormworkDebugLog.Flush();
        }

        /// <summary>
        /// 鉄骨検出の対象カテゴリ判定。構造柱・構造フレームのみ対象とする。
        /// </summary>
        private static bool IsSteelDetectionTarget(Element elem)
        {
            if (elem?.Category == null) return false;
            int catId = elem.Category.Id.IntValue();
            return catId == (int)BuiltInCategory.OST_StructuralColumns
                || catId == (int)BuiltInCategory.OST_StructuralFraming;
        }

        /// <summary>
        /// デッキスラブ検出の対象カテゴリ判定。床のみ対象とする。
        /// </summary>
        private static bool IsDeckSlabDetectionTarget(Element elem)
        {
            if (elem?.Category == null) return false;
            return elem.Category.Id.IntValue() == (int)BuiltInCategory.OST_Floors;
        }

        /// <summary>
        /// 鉄骨階段検出の対象カテゴリ判定。階段のみ対象とする。
        /// </summary>
        private static bool IsStairDetectionTarget(Element elem)
        {
            if (elem?.Category == null) return false;
            return elem.Category.Id.IntValue() == (int)BuiltInCategory.OST_Stairs;
        }

        /// <summary>
        /// 壁検出の対象カテゴリ判定。壁本体 (Wall) のみ対象とし、WallSweep は除外する。
        /// </summary>
        private static bool IsWallDetectionTarget(Element elem)
        {
            if (elem == null) return false;
            if (elem is WallSweep) return false;
            return elem is Wall;
        }

        /// <summary>
        /// 単一ドキュメントから対象カテゴリの要素を収集する (ホスト・リンクで共通)。
        /// activeView は host の現在ビューモード用。リンク収集時には null。
        /// linkLocalOutline が非 null の場合、BoundingBoxIntersectsFilter で要素範囲を絞り込む
        /// (BIM360/大規模リンクモデルでのフリーズ防止)。
        /// </summary>
        internal static List<Element> CollectFromDoc(
            Document doc, FormworkSettings settings, View activeView, bool isLinked,
            Outline linkLocalOutline = null)
        {
            var result = new List<Element>();
            var seenIds = new HashSet<int>();
            bool useViewFilter = !isLinked
                && settings.Scope == CalculationScope.CurrentView
                && activeView != null;

            foreach (var key in settings.IncludedCategories)
            {
                if (!_nameToCat.TryGetValue(key, out var bic))
                    continue;

                FilteredElementCollector col = useViewFilter
                    ? new FilteredElementCollector(doc, activeView.Id)
                    : new FilteredElementCollector(doc);

                IList<Element> elems;
                if (linkLocalOutline != null)
                {
                    // セクションボックス範囲フィルタ: BIM360 等の大規模リンクモデルで
                    // ビュー内の要素のみに絞り込み、不要な geometry アクセスを回避する
                    elems = col.OfCategory(bic)
                        .WhereElementIsNotElementType()
                        .WherePasses(new BoundingBoxIntersectsFilter(linkLocalOutline))
                        .ToList();
                }
                else
                {
                    elems = col.OfCategory(bic).WhereElementIsNotElementType().ToList();
                }

                if (bic == BuiltInCategory.OST_Walls)
                {
                    elems = elems.Where(e =>
                    {
                        if (!(e is Wall w)) return false;
                        var wt = doc.GetElement(w.GetTypeId()) as WallType;
                        if (wt == null) return true;
                        return wt.Function == WallFunction.Retaining ||
                               wt.Kind == WallKind.Basic;
                    }).ToList();

                    foreach (var e in elems)
                    {
                        if (seenIds.Add(e.Id.IntValue())) result.Add(e);
                    }

                    // WallSweep (壁スイープ・リビール) も追加。リンクモデルの WallSweep も
                    // 含めることで、リンクの壁にスイープが取り付いている場合に、
                    // スイープ自体の外側面 (前/上/下) に対する型枠も算出できる。
                    FilteredElementCollector swCol = useViewFilter
                        ? new FilteredElementCollector(doc, activeView.Id)
                        : new FilteredElementCollector(doc);
                    ICollection<Element> sweeps;
                    if (linkLocalOutline != null)
                    {
                        sweeps = swCol.OfClass(typeof(WallSweep))
                            .WhereElementIsNotElementType()
                            .WherePasses(new BoundingBoxIntersectsFilter(linkLocalOutline))
                            .ToList();
                    }
                    else
                    {
                        sweeps = swCol.OfClass(typeof(WallSweep))
                            .WhereElementIsNotElementType()
                            .ToList();
                    }
                    int addedSweeps = 0;
                    foreach (var sw in sweeps)
                    {
                        if (seenIds.Add(sw.Id.IntValue())) { result.Add(sw); addedSweeps++; }
                    }
                    if (isLinked && FormworkDebugLog.Enabled)
                    {
                        FormworkDebugLog.Log(
                            $"  [LinkSweepDiag] CollectFromDoc(linkDoc='{doc?.Title}') " +
                            $"WallSweep found={sweeps.Count} added={addedSweeps}");
                    }
                }
                else
                {
                    foreach (var e in elems)
                    {
                        if (seenIds.Add(e.Id.IntValue())) result.Add(e);
                    }
                }
            }

            // 個別非表示 (右クリック「ビューで非表示 > 要素」) の除外。
            // FilteredElementCollector(doc, viewId) はカテゴリ・ワークセット非表示は
            // 除外するが、個別要素レベルの非表示は除外しないため、ここで明示的に除く。
            if (useViewFilter)
            {
                int before = result.Count;
                result = result.Where(e =>
                {
                    try { return !e.IsHidden(activeView); }
                    catch { return true; }
                }).ToList();
                int hidden = before - result.Count;
                if (hidden > 0 && FormworkDebugLog.Enabled)
                    FormworkDebugLog.Log(
                        $"  [HostCollect] 個別非表示要素を除外: {hidden} 要素 (view='{activeView.Name}')");
            }
            return result;
        }

        /// <summary>互換のため残しているが、新規コードからは呼ばない。CollectFromDoc を使う。</summary>
        internal static List<Element> Collect(Document doc, FormworkSettings settings, View activeView)
        {
            return CollectFromDoc(doc, settings, activeView, isLinked: false);
        }

        internal static CategoryGroup ToCategoryGroup(Element elem)
        {
            if (elem == null) return CategoryGroup.Other;
            if (elem is WallSweep) return CategoryGroup.Wall;
            if (elem.Category == null) return CategoryGroup.Other;
            switch ((BuiltInCategory)elem.Category.Id.IntValue())
            {
                case BuiltInCategory.OST_StructuralColumns: return CategoryGroup.Column;
                case BuiltInCategory.OST_StructuralFraming: return CategoryGroup.Beam;
                case BuiltInCategory.OST_Walls: return CategoryGroup.Wall;
                case BuiltInCategory.OST_Floors: return CategoryGroup.Slab;
                case BuiltInCategory.OST_StructuralFoundation: return CategoryGroup.Foundation;
                case BuiltInCategory.OST_Stairs: return CategoryGroup.Stairs;
                case BuiltInCategory.OST_Roofs: return CategoryGroup.Roof;
                default: return CategoryGroup.Other;
            }
        }

        /// <summary>
        /// 要素の関連レベル名を取得する（スラブ・壁・柱・梁など各カテゴリを個別に対応）。
        /// 見つからない場合は空文字を返す。
        /// </summary>
        internal static string GetElementLevelName(Element elem)
        {
            if (elem == null) return string.Empty;
            var doc = elem.Document;

            ElementId levelId = null;
            try
            {
                if (elem is Wall w) levelId = w.LevelId;
                else if (elem is Floor floor) levelId = floor.LevelId;
                else if (elem.LevelId != null && elem.LevelId != ElementId.InvalidElementId)
                    levelId = elem.LevelId;
            }
            catch { }

            if (levelId == null || levelId == ElementId.InvalidElementId)
            {
                var candidates = new[]
                {
                    BuiltInParameter.LEVEL_PARAM,
                    BuiltInParameter.FAMILY_LEVEL_PARAM,
                    BuiltInParameter.SCHEDULE_LEVEL_PARAM,
                    BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                    BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
                };
                foreach (var bip in candidates)
                {
                    try
                    {
                        var p = elem.get_Parameter(bip);
                        if (p != null && p.HasValue && p.StorageType == StorageType.ElementId)
                        {
                            var id = p.AsElementId();
                            if (id != null && id != ElementId.InvalidElementId)
                            {
                                levelId = id;
                                break;
                            }
                        }
                    }
                    catch { }
                }
            }

            if (levelId == null || levelId == ElementId.InvalidElementId) return string.Empty;
            var level = doc.GetElement(levelId) as Level;
            return level?.Name ?? string.Empty;
        }

        internal static string GetParameterString(Element elem, string paramName)
        {
            if (elem == null || string.IsNullOrEmpty(paramName)) return string.Empty;
            Parameter p = elem.LookupParameter(paramName);
            if (p == null) return string.Empty;
            switch (p.StorageType)
            {
                case StorageType.String:
                    return p.AsString() ?? string.Empty;
                case StorageType.Integer:
                    return p.AsInteger().ToString();
                case StorageType.Double:
                    return p.AsValueString() ?? p.AsDouble().ToString("F2");
                case StorageType.ElementId:
                    return p.AsValueString() ?? string.Empty;
                default:
                    return string.Empty;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 型枠数量算出のメインクラス。リンクモデル対応・壁スイープ面の型枠算出対応済み。
    /// Pass 1: 要素毎の Solid 取得・面分類
    /// Pass 2: 幾何学的な接触面検出 → DeductedContact
    /// Pass 3: 開口加算・集計
    /// </summary>
    internal class FormworkCalcEngine
    {
        private readonly Document _doc;
        private readonly FormworkSettings _settings;
        private readonly View _activeView;
        private readonly IProgress<string> _progress;

        // アクティブビューに切断ボックスがある場合の Solid (ホスト世界座標)。
        // CurrentView モードでリンク要素の solid をこの範囲にクリップする。
        private Solid _sectionBoxSolid;

        internal FormworkCalcEngine(
            Document doc,
            FormworkSettings settings,
            View activeView,
            IProgress<string> progress = null)
        {
            _doc = doc;
            _settings = settings;
            _activeView = activeView;
            _progress = progress;
            _sectionBoxSolid = BuildSectionBoxSolid(activeView, settings);
        }

        /// <summary>
        /// CurrentView モードでアクティブな 3D ビューに切断ボックスがある場合、
        /// その切断ボックスをホスト世界座標系の Solid として構築する。
        /// 該当しない場合は null を返す。
        /// </summary>
        private static Solid BuildSectionBoxSolid(View view, FormworkSettings settings)
        {
            if (settings == null) return null;
            if (settings.Scope != CalculationScope.CurrentView
                && settings.Scope != CalculationScope.SelectedViews) return null;
            if (!(view is View3D v3d)) return null;
            BoundingBoxXYZ sb = null;
            try { if (v3d.IsSectionBoxActive) sb = v3d.GetSectionBox(); }
            catch { }
            if (sb == null) return null;

            try
            {
                var min = sb.Min;
                var max = sb.Max;
                var p0 = new XYZ(min.X, min.Y, min.Z);
                var p1 = new XYZ(max.X, min.Y, min.Z);
                var p2 = new XYZ(max.X, max.Y, min.Z);
                var p3 = new XYZ(min.X, max.Y, min.Z);
                var loop = new CurveLoop();
                loop.Append(Line.CreateBound(p0, p1));
                loop.Append(Line.CreateBound(p1, p2));
                loop.Append(Line.CreateBound(p2, p3));
                loop.Append(Line.CreateBound(p3, p0));
                double height = max.Z - min.Z;
                Solid s = GeometryCreationUtilities.CreateExtrusionGeometry(
                    new[] { loop }, XYZ.BasisZ, height);
                var tr = sb.Transform;
                if (tr != null && !tr.IsIdentity)
                    s = SolidUtils.CreateTransformed(s, tr);
                FormworkDebugLog.Log(
                    $"[Engine] 切断ボックスSolid構築: vol={s?.Volume:F3}ft³");
                return s;
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"[Engine] BuildSectionBoxSolid EX: {ex.Message}");
                return null;
            }
        }

        internal FormworkResult Run()
        {
            var result = new FormworkResult();

            FormworkDebugLog.Initialize(_settings?.EnableDebugLog == true);
            // エンジン開始直後に即フラッシュ: フリーズ→強制終了してもログ開始が残る。
            // これ以前にログが無ければフリーズはダイアログ内か Initialize 前で発生している。
            FormworkDebugLog.Log(
                $"[Engine] Run() 開始 Scope={_settings?.Scope} " +
                $"IncludeLinks={_settings?.IncludeLinkedModels} " +
                $"EnableDebugLog={_settings?.EnableDebugLog}");
            // 切断ボックスSolid状態を Initialize 後にログ（コンストラクタで構築済み）
            if (_sectionBoxSolid != null)
                FormworkDebugLog.Log($"[Engine] 切断ボックスSolid: 有効 vol={_sectionBoxSolid.Volume:F3}ft³");
            else
                FormworkDebugLog.Log($"[Engine] 切断ボックスSolid: なし (Scope={_settings?.Scope} / SectionBox未設定またはCurrentView以外)");
            FormworkDebugLog.Flush();
            // 注: ここで Close() しない。後続の CreateVisualization / CreateSchedule もログ出力するため、
            // クローズは Command 側 (FormworkCalculatorCommand) で全処理完了後に行う。
            return RunCore(result);
        }

        private FormworkResult RunCore(FormworkResult result)
        {
            _progress?.Report("要素を収集中...");
            FormworkDebugLog.Log("[Engine] CollectAndClassify 開始");
            FormworkDebugLog.Flush();
            var collection = ElementCollector.CollectAndClassify(_doc, _settings, _activeView);
            var sources = collection.Targets;
            result.ProcessedElementCount = sources.Count;
            result.SourceRegistry = collection.Registry;
            result.LinkedInstanceCount = collection.LinkedInstanceCount;
            FormworkDebugLog.Log(
                $"Collected elements: targets={sources.Count} " +
                $"excluded={collection.Excluded.Count} " +
                $"linkedInstances={collection.LinkedInstanceCount}");
            FormworkDebugLog.Flush(); // 収集完了確認ポイント

            // [LinkMap] リンクソースごとの集計を 1 行で出力 (要素数節約のため per-element ログは廃止)。
            // 個別 real ID → surrogate ID マッピングが必要な場合は Pass 1 の [Pass1] ログを参照。
            if (FormworkDebugLog.Enabled)
            {
                var linkSummary = new System.Collections.Generic.Dictionary<string, (int count, int minSrg, int maxSrg)>();
                foreach (var src in collection.Registry.All())
                {
                    if (!src.IsLinked) continue;
                    if (!linkSummary.TryGetValue(src.SourceName, out var s))
                        s = (0, src.SurrogateId, src.SurrogateId);
                    linkSummary[src.SourceName] = (
                        s.count + 1,
                        Math.Min(s.minSrg, src.SurrogateId),
                        Math.Max(s.maxSrg, src.SurrogateId));
                }
                int linkMapTotal = 0;
                foreach (var kv in linkSummary)
                {
                    FormworkDebugLog.Log(
                        $"  [LinkMap] src='{kv.Key}' count={kv.Value.count} " +
                        $"srgRange=[{kv.Value.minSrg}..{kv.Value.maxSrg}]");
                    linkMapTotal += kv.Value.count;
                }
                FormworkDebugLog.Log($"  [LinkMap] total linked elements mapped: {linkMapTotal}");
                FormworkDebugLog.Flush(); // 収集完了確認ポイント
            }

            // 除外要素を ExcludedResult として記録 (集計には含めない)
            foreach (var ex in collection.Excluded)
            {
                var e = ex.Element;
                result.ExcludedResults.Add(new ExcludedResult
                {
                    ElementId = ex.Source != null ? ex.Source.SurrogateId : e.Id.IntValue(),
                    ElementName = e.Name ?? string.Empty,
                    Category = ElementCollector.ToCategoryGroup(e),
                    CategoryName = e.Category?.Name ?? string.Empty,
                    SourceName = ex.Source?.SourceName ?? ElementSourceRegistry.HostSourceName,
                    Kind = ex.Kind,
                    DetectionLayer = ex.Layer,
                    DetectionReason = ex.Reason,
                });
            }

            if (sources.Count == 0) return result;

            // 壁の化粧目地 (Reveal, OST_Reveals — subtractive WallSweep) を型枠計算から除外する。
            // 化粧目地は壁を切り欠く feature であり、独立した躯体要素ではない。
            // - 壁の solid は化粧目地によってカットされ、切欠き面 (cut face) は wall solid 側に含まれる
            // - 切欠き面の型枠は ContactFaceDetector のアシメトリ保護 + SlopedWallTop の絞り込みで
            //   wall 側に FormworkRequired として残る
            // - 化粧目地 element 自体は cut volume を表現する「ファントム」であり、
            //   露出面に型枠を生成すべきではない
            // ※ Sweep 型 (OST_Cornices — 装飾的な加算 sweep) は除外しない (露出面の型枠は必要)
            var revealsToExclude = sources.Where(s =>
                s.Element is WallSweep &&
                s.Element.Category?.Id.IntValue() == (int)BuiltInCategory.OST_Reveals
            ).ToList();
            if (revealsToExclude.Count > 0)
            {
                foreach (var src in revealsToExclude)
                {
                    var elem = src.Element;
                    result.ExcludedResults.Add(new ExcludedResult
                    {
                        ElementId = src.SurrogateId,
                        ElementName = elem.Name ?? string.Empty,
                        Category = CategoryGroup.Wall,
                        CategoryName = elem.Category?.Name ?? string.Empty,
                        SourceName = src.SourceName,
                        Kind = ExclusionKind.WallSweep,
                        DetectionLayer = "Reveal",
                        DetectionReason = "壁の化粧目地 (型枠は壁の切欠き面で算入)",
                    });
                }
                sources.RemoveAll(s =>
                    s.Element is WallSweep &&
                    s.Element.Category?.Id.IntValue() == (int)BuiltInCategory.OST_Reveals);
                FormworkDebugLog.Log(
                    $"  Excluded {revealsToExclude.Count} Reveal (化粧目地) elements " +
                    $"(formwork accounted on wall cut-face side)");
            }

            // 開口処理はホスト要素のみで完結 (リンク内の開口はリンク側に閉じる)
            var hostElements = sources
                .Where(s => !s.IsLinked)
                .Select(s => s.Element)
                .ToList();
            var openingDeltas = OpeningProcessor.Compute(_doc, hostElements);
            var openingMap = openingDeltas.ToDictionary(o => o.HostElementId, o => o);

            // Pass 1: 要素毎に面を分類
            _progress?.Report($"Pass 1: 面分類中... 0 / {sources.Count}");
            FormworkDebugLog.Log(
                $"[Pass1] 開始: total={sources.Count} " +
                $"linked={sources.Count(s => s.IsLinked)} host={sources.Count(s => !s.IsLinked)}");
            FormworkDebugLog.Flush();
            var contexts = new List<ContactFaceDetector.ElementFacesContext>();
            var srcByContext = new Dictionary<int, ElementSource>();

            int idx = 0;
            int linkedCount = sources.Count(s => s.IsLinked);
            int linkedIdx = 0;
            foreach (var src in sources)
            {
                idx++;
                if (src.IsLinked) linkedIdx++;

                // プログレス表示: リンク要素は個別に更新してフリーズと区別できるようにする
                if (src.IsLinked)
                    _progress?.Report(
                        $"Pass 1: リンク要素 {linkedIdx}/{linkedCount} 処理中... " +
                        $"({src.Element.Category?.Name ?? "?"} E{src.Element.Id.IntValue()})");
                else if (idx % 10 == 0)
                    _progress?.Report($"Pass 1: 面分類中... {idx} / {sources.Count}");

                // フリーズ診断: リンク要素の GetSolids (= get_Geometry) 呼び出し前にログ。
                // 100要素ごとにフラッシュ (全件フラッシュはI/O過多でかえって遅くなるため)。
                if (src.IsLinked && FormworkDebugLog.Enabled)
                {
                    FormworkDebugLog.Log(
                        $"[Pass1] リンク [{linkedIdx}/{linkedCount}] " +
                        $"E{src.Element.Id.IntValue()} " +
                        $"cat='{src.Element.Category?.Name}' " +
                        $"src='{src.SourceName}' → GetSolids開始");
                    if (linkedIdx % 100 == 0) FormworkDebugLog.Flush();
                }

                var elem = src.Element;
                try
                {
                    var ctx = ClassifyElementFaces(src);
                    if (ctx != null)
                    {
                        contexts.Add(ctx);
                        srcByContext[ctx.ElementId] = src;
                    }
                }
                catch (Exception ex)
                {
                    result.Errors.Add(new ErrorLogEntry
                    {
                        ElementId = src.SurrogateId,
                        CategoryName = elem.Category?.Name ?? string.Empty,
                        ElementName = elem.Name,
                        ErrorKind = "ClassifyError",
                        Message = ex.Message,
                    });
                }
            }

            // Pass 1 完了時: 各要素の面分類内訳をログ
            LogPass1Summary(contexts);

            // [LinkSweepDiag] Pass 2 前後の face type 変化を追跡するためのスナップショット
            var linkSweepFaceSnapshot = SnapshotLinkSweepFaceTypes(contexts, srcByContext);

            // Pass 2: 幾何学的な直接検査で接触面を検出して控除
            _progress?.Report("Pass 2: 接触面を検出中...");
            ContactFaceDetector.RefineContactFaces(contexts);
            LogLinkSweepFaceTypeDelta(
                contexts, srcByContext, linkSweepFaceSnapshot, "post-RefineContactFaces");

            // Pass 2b: WallSweep (スイープ・リビール) の host 壁面を直接 DeductedContact 化
            // していたが、ユーザー仕様「壁と化粧目地の接触面に型枠を残す」のため無効化。
            // ContactFaceDetector 側で WallSweep 接触の壁側を FormworkRequired のまま保持する
            // ようになったので、本処理は不要 (WallSweep が solid を持つ場合は ContactFaceDetector
            // が処理、持たない場合は wall 面が単純に FormworkRequired のまま残るのが望ましい)。
            // WallSweepFaceDeductor.DeductWallFacesNearSweeps(_doc, contexts, srcByContext);
            LogLinkSweepFaceTypeDelta(
                contexts, srcByContext, linkSweepFaceSnapshot, "post-WallSweepFaceDeductor");

            // Pass 2 完了時: 最終 FormworkRequired 面数
            LogPostPass2Summary(contexts);

            // Pass 3: 開口加算 + ElementResult 作成
            _progress?.Report("Pass 3: 集計中...");
            foreach (var ctx in contexts)
            {
                var src = srcByContext[ctx.ElementId];
                try
                {
                    var er = BuildElementResult(src, ctx, openingMap);
                    if (er != null) result.ElementResults.Add(er);
                }
                catch (Exception ex)
                {
                    result.Errors.Add(new ErrorLogEntry
                    {
                        ElementId = ctx.ElementId,
                        CategoryName = src.Element.Category?.Name ?? string.Empty,
                        ElementName = src.Element.Name,
                        ErrorKind = "AggregateError",
                        Message = ex.Message,
                    });
                }
            }

            // 壁スイープ・リビールの ElementResult はそのまま残す。
            // スイープ自体の外側面 (前/上/下/端部) は型枠が必要な部分なので、
            // 通常の躯体要素と同様に DirectShape として可視化・集計する。
            // (旧実装では ExcludedResults に移していたが、スイープ面の型枠が
            //  算出されない問題があったため除去)

            // 診断ログ: 各要素の集計結果を一覧出力 (⚠️ マーカーで疑わしい要素を強調)
            LogElementDiagnostics(collection.Registry, contexts, result);

            Aggregate(result);
            return result;
        }

        /// <summary>
        /// 各要素の集計結果と面の状態を診断ログに出力する。
        /// 疑わしいパターンには ⚠️ マーカーを付け、原因特定の手がかりとする。
        /// 検索しやすいように 1 要素 1 行で出力。
        /// </summary>
        private static void LogElementDiagnostics(
            ElementSourceRegistry registry,
            List<ContactFaceDetector.ElementFacesContext> contexts,
            FormworkResult result)
        {
            if (!FormworkDebugLog.Enabled) return;

            var ctxById = new Dictionary<int, ContactFaceDetector.ElementFacesContext>();
            foreach (var c in contexts) ctxById[c.ElementId] = c;

            FormworkDebugLog.Section($"Element Diagnostics ({result.ElementResults.Count} elements)");
            FormworkDebugLog.Log("  fmt: [ElemDiag] E<id> (cat) 'name' type='type' formwork=Xm² faces=R/T/B/C/B(GL)/I parts=N");
            FormworkDebugLog.Log("       ⚠️ markers: ZERO=formwork≈0, MOSTLY_DED=控除>>formwork, ALL_PARTIAL=全FormworkRequired面が部分接触");

            foreach (var er in result.ElementResults)
            {
                if (!ctxById.TryGetValue(er.ElementId, out var ctx)) continue;

                string typeName = "?";
                try
                {
                    var src = registry?.Get(er.ElementId);
                    var elem = src?.Element;
                    if (elem != null)
                    {
                        var t = src.SourceDoc.GetElement(elem.GetTypeId());
                        typeName = t?.Name ?? "?";
                    }
                }
                catch { }

                int reqCount = 0, topCount = 0, botCount = 0, conCount = 0, incCount = 0;
                int partialFaceCount = 0;
                int allReqFaces = 0;
                foreach (var fi in ctx.Faces)
                {
                    switch (fi.FaceType)
                    {
                        case FaceType.FormworkRequired:
                            reqCount++;
                            allReqFaces++;
                            if (fi.PartialContacts.Count > 0) partialFaceCount++;
                            break;
                        case FaceType.DeductedTop: topCount++; break;
                        case FaceType.DeductedBottom: botCount++; break;
                        case FaceType.DeductedContact: conCount++; break;
                        case FaceType.Inclined: incCount++; break;
                    }
                }

                // ⚠️ マーカー判定 (見やすくするため絞り込み)
                var marks = new List<string>();
                // CONTACT_HEAVY: dedTop/dedBot は基礎・スラブで自然に大きいため除外し、
                //                純粋な contact 控除 (dedCon) が formwork の 3 倍以上の場合のみ警告
                if (er.FormworkArea > 0 && er.DeductedContactArea > er.FormworkArea * 3)
                    marks.Add("⚠️CONTACT_HEAVY");
                // ZERO: FormworkRequired 面が無く、contact 面が多い (柱の埋め込み等)
                if (er.FormworkArea < 0.01 && conCount >= 4)
                    marks.Add("⚠️ZERO_EMBEDDED");
                else if (er.FormworkArea < 0.01 && reqCount > 0)
                    marks.Add("⚠️ZERO");
                if (allReqFaces > 0 && partialFaceCount == allReqFaces)
                    marks.Add("⚠️ALL_PARTIAL");
                string marker = marks.Count > 0 ? " " + string.Join("|", marks) : "";

                // 寸法情報 (BB から要素の概算寸法を算出)
                string dim = "?";
                try
                {
                    if (ctx.BB != null)
                    {
                        double lx = (ctx.BB.Max.X - ctx.BB.Min.X) * 304.8;  // ft → mm
                        double ly = (ctx.BB.Max.Y - ctx.BB.Min.Y) * 304.8;
                        double lz = (ctx.BB.Max.Z - ctx.BB.Min.Z) * 304.8;
                        dim = $"{lx:F0}x{ly:F0}x{lz:F0}";
                    }
                }
                catch { }

                string cat = CategoryShort(er.Category);
                FormworkDebugLog.Log(
                    $"[ElemDiag] E{er.ElementId} ({cat}) '{er.ElementName}' type='{typeName}' " +
                    $"dim={dim}mm formwork={er.FormworkArea:F2}m² faces={reqCount}/{topCount}/{botCount}/{conCount}/{incCount} " +
                    $"parts={partialFaceCount} dedTop={er.DeductedTopArea:F2} dedBot={er.DeductedBottomArea:F2} " +
                    $"dedCon={er.DeductedContactArea:F2}{marker}");
            }
            FormworkDebugLog.Flush();
        }

        /// <summary>
        /// 面の水平射影 (XY平面) のバウンディングボックス短辺を feet 単位で返す。
        /// 壁の斜めの天端で「斜面の幅」を判定するために使用する。
        /// 面のエッジを 10 分割サンプリングして XY 投影の min/max を求める。
        /// </summary>
        private static double ComputeFaceHorizontalShortDim(Face face)
        {
            if (face == null) return 0;
            double minX = double.MaxValue, maxX = double.MinValue;
            double minY = double.MaxValue, maxY = double.MinValue;
            int sampleCount = 0;
            try
            {
                foreach (EdgeArray ea in face.EdgeLoops)
                {
                    foreach (Edge e in ea)
                    {
                        Curve c = null;
                        try { c = e.AsCurve(); } catch { }
                        if (c == null) continue;
                        for (int i = 0; i <= 10; i++)
                        {
                            double t = i / 10.0;
                            XYZ p;
                            try { p = c.Evaluate(t, true); } catch { continue; }
                            if (p == null) continue;
                            if (p.X < minX) minX = p.X;
                            if (p.X > maxX) maxX = p.X;
                            if (p.Y < minY) minY = p.Y;
                            if (p.Y > maxY) maxY = p.Y;
                            sampleCount++;
                        }
                    }
                }
            }
            catch { return 0; }
            if (sampleCount < 2) return 0;
            double dx = maxX - minX;
            double dy = maxY - minY;
            return Math.Min(dx, dy);
        }

        private static string CategoryShort(CategoryGroup cg)
        {
            switch (cg)
            {
                case CategoryGroup.Column: return "柱";
                case CategoryGroup.Beam: return "梁";
                case CategoryGroup.Wall: return "壁";
                case CategoryGroup.Slab: return "スラブ";
                case CategoryGroup.Foundation: return "基礎";
                case CategoryGroup.Stairs: return "階段";
                case CategoryGroup.Roof: return "屋根";
                default: return "他";
            }
        }

        /// <summary>
        /// WallSweep (壁スイープ・リビール) を ElementResults から ExcludedResults に移動する。
        /// 早期除外せず Pass 1/2 に参加させる理由: Wall の top 面が reveal で切り取られた
        /// 部分を contact detection によって DeductedContact 化するため。
        /// </summary>
        private static void MoveWallSweepsToExcluded(ElementSourceRegistry registry, FormworkResult result)
        {
            var toMove = new List<ElementResult>();
            foreach (var er in result.ElementResults)
            {
                var src = registry?.Get(er.ElementId);
                if (src?.Element is WallSweep) toMove.Add(er);
            }
            if (toMove.Count == 0) return;

            foreach (var er in toMove)
            {
                var src = registry?.Get(er.ElementId);
                var elem = src?.Element;
                result.ElementResults.Remove(er);
                result.ExcludedResults.Add(new ExcludedResult
                {
                    ElementId = er.ElementId,
                    ElementName = er.ElementName,
                    Category = CategoryGroup.Wall,
                    CategoryName = elem?.Category?.Name ?? string.Empty,
                    SourceName = src?.SourceName ?? ElementSourceRegistry.HostSourceName,
                    Kind = ExclusionKind.WallSweep,
                    DetectionLayer = "WallSweep",
                    DetectionReason = "壁スイープ・リビール (post-processed after contact detection)",
                });
            }
            FormworkDebugLog.Log($"  Moved {toMove.Count} WallSweep elements to ExcludedResults");
        }

        private static void LogPass1Summary(List<ContactFaceDetector.ElementFacesContext> contexts)
        {
            if (!FormworkDebugLog.Enabled) return;
            FormworkDebugLog.Section($"Pass 1: Face Classification (elements={contexts.Count})");
            int totalReq = 0, totalTop = 0, totalBot = 0, totalCon = 0, totalInc = 0, totalErr = 0;
            foreach (var ctx in contexts)
            {
                int req = 0, top = 0, bot = 0, con = 0, inc = 0, err = 0;
                foreach (var f in ctx.Faces)
                {
                    switch (f.FaceType)
                    {
                        case FaceType.FormworkRequired: req++; break;
                        case FaceType.DeductedTop: top++; break;
                        case FaceType.DeductedBottom: bot++; break;
                        case FaceType.DeductedContact: con++; break;
                        case FaceType.Inclined: inc++; break;
                        case FaceType.Error: err++; break;
                    }
                }
                totalReq += req; totalTop += top; totalBot += bot; totalCon += con;
                totalInc += inc; totalErr += err;
                FormworkDebugLog.Log(
                    $"  E{ctx.ElementId} ({ctx.Category}/{ctx.CategoryName}) " +
                    $"total={ctx.Faces.Count} Req={req} Top={top} Bot={bot} Con={con} Inc={inc} Err={err}");
            }
            FormworkDebugLog.Log(
                $"  TOTAL: Req={totalReq} Top={totalTop} Bot={totalBot} Con={totalCon} " +
                $"Inc={totalInc} Err={totalErr}");
            FormworkDebugLog.Flush();
        }

        /// <summary>
        /// [LinkSweepDiag] Pass 2 前のリンク WallSweep の face type 一覧を撮影する。
        /// </summary>
        private static Dictionary<int, List<FaceType>> SnapshotLinkSweepFaceTypes(
            List<ContactFaceDetector.ElementFacesContext> contexts,
            Dictionary<int, ElementSource> srcByContext)
        {
            var snap = new Dictionary<int, List<FaceType>>();
            if (!FormworkDebugLog.Enabled) return snap;
            foreach (var ctx in contexts)
            {
                if (!srcByContext.TryGetValue(ctx.ElementId, out var src)) continue;
                if (!src.IsLinked) continue;
                if (!(src.Element is WallSweep)) continue;
                var copy = new List<FaceType>(ctx.Faces.Count);
                foreach (var f in ctx.Faces) copy.Add(f.FaceType);
                snap[ctx.ElementId] = copy;
            }
            return snap;
        }

        /// <summary>
        /// [LinkSweepDiag] スナップショットと比較して、各リンク WallSweep の face type が
        /// 何個 → 何個に変化したかをログ出力する。
        /// </summary>
        private static void LogLinkSweepFaceTypeDelta(
            List<ContactFaceDetector.ElementFacesContext> contexts,
            Dictionary<int, ElementSource> srcByContext,
            Dictionary<int, List<FaceType>> snap,
            string stageLabel)
        {
            if (!FormworkDebugLog.Enabled || snap == null || snap.Count == 0) return;
            foreach (var ctx in contexts)
            {
                if (!snap.TryGetValue(ctx.ElementId, out var before)) continue;
                if (!srcByContext.TryGetValue(ctx.ElementId, out var src)) continue;
                int reqB = before.Count(t => t == FaceType.FormworkRequired);
                int conB = before.Count(t => t == FaceType.DeductedContact);
                int reqA = 0, conA = 0;
                foreach (var f in ctx.Faces)
                {
                    if (f.FaceType == FaceType.FormworkRequired) reqA++;
                    else if (f.FaceType == FaceType.DeductedContact) conA++;
                }
                FormworkDebugLog.Log(
                    $"  [LinkSweepDiag] {stageLabel} E{src.Element.Id.IntValue()} src='{src.SourceName}' " +
                    $"FormworkRequired {reqB}→{reqA}  DeductedContact {conB}→{conA}");
                // 各 face の遷移詳細 (面数 ≤ 8 のみ詳細出力)
                if (ctx.Faces.Count <= 8)
                {
                    for (int i = 0; i < ctx.Faces.Count && i < before.Count; i++)
                    {
                        if (before[i] != ctx.Faces[i].FaceType)
                        {
                            FormworkDebugLog.Log(
                                $"    face[{i}] {before[i]} → {ctx.Faces[i].FaceType} " +
                                $"(area={ctx.Faces[i].Area:F4}ft², partials={ctx.Faces[i].PartialContacts.Count})");
                        }
                    }
                }
                // スナップショットを最新状態に更新 (次の stage の比較用)
                snap[ctx.ElementId] = ctx.Faces.Select(f => f.FaceType).ToList();
            }
        }

        private static void LogPostPass2Summary(List<ContactFaceDetector.ElementFacesContext> contexts)
        {
            if (!FormworkDebugLog.Enabled) return;
            int req = 0, con = 0;
            foreach (var ctx in contexts)
            {
                foreach (var f in ctx.Faces)
                {
                    if (f.FaceType == FaceType.FormworkRequired) req++;
                    else if (f.FaceType == FaceType.DeductedContact) con++;
                }
            }
            FormworkDebugLog.Section("Post Pass 2 Totals");
            FormworkDebugLog.Log($"  FormworkRequired: {req}");
            FormworkDebugLog.Log($"  DeductedContact:  {con}");

            // 各面の詳細ダンプ (個別要素の不具合追跡用)。面数 ≥ 6 の要素のみ。
            FormworkDebugLog.Section("Per-Face Detail (elements with ≥6 faces)");
            FormworkDebugLog.Log("  fmt: [FaceDetail] E<id> face[i] type=<T> area=<ft²> n=(x,y,z) center=(x,y,z)");
            foreach (var ctx in contexts)
            {
                if (ctx.Faces.Count < 6) continue;
                for (int i = 0; i < ctx.Faces.Count; i++)
                {
                    var fi = ctx.Faces[i];
                    var n = fi.Normal;
                    string nStr = n != null
                        ? $"({n.X:F3},{n.Y:F3},{n.Z:F3})" : "(null)";
                    string cStr = "(?,?,?)";
                    try
                    {
                        var bb = fi.Face.GetBoundingBox();
                        if (bb != null)
                        {
                            var midUV = (bb.Min + bb.Max) * 0.5;
                            var p = fi.Face.Evaluate(midUV);
                            if (p != null)
                                cStr = $"({p.X:F3},{p.Y:F3},{p.Z:F3})";
                        }
                    }
                    catch { }
                    FormworkDebugLog.Log(
                        $"[FaceDetail] E{ctx.ElementId} face[{i}] type={fi.FaceType} " +
                        $"area={fi.Area:F4} n={nStr} center={cStr} partials={fi.PartialContacts.Count}");
                }
            }
            FormworkDebugLog.Flush();
        }

        private ContactFaceDetector.ElementFacesContext ClassifyElementFaces(
            ElementSource src)
        {
            var elem = src.Element;

            // リンク要素は WallSweep フォールバック (IncludeNonVisibleObjects=true) をスキップ:
            // BIM360ワークシェアリング環境では2回目の get_Geometry が非常に遅いため。
            var solids = SolidUnionProcessor.GetSolids(elem, src.Transform,
                skipWallSweepRetry: src.IsLinked);
            bool isLinkSweep = src.IsLinked && elem is WallSweep;

            // リンク要素のみ: アクティブビューの切断ボックスでクリップする。
            // ホスト要素は FilteredElementCollector(doc, viewId) が切断ボックス外を除外するが、
            // straddle 要素 (切断ボックス境界をまたぐ要素) 残るため、host も同様にクリップする。
            if (_sectionBoxSolid != null && src.IsLinked && solids.Count > 0)
            {
                double volBefore = solids.Sum(s => s?.Volume ?? 0);
                var clipped = new List<Solid>();
                foreach (var s in solids)
                {
                    if (s == null) continue;
                    try
                    {
                        var c = BooleanOperationsUtils.ExecuteBooleanOperation(
                            s, _sectionBoxSolid, BooleanOperationsType.Intersect);
                        if (c != null && c.Volume > 1e-6) clipped.Add(c);
                    }
                    catch (Exception ex)
                    {
                        if (FormworkDebugLog.Enabled)
                            FormworkDebugLog.Log(
                                $"  [SectionBoxClip] E{elem.Id.IntValue()} src='{src.SourceName}' " +
                                $"intersect EX: {ex.Message} → 元のSolidを保持");
                        clipped.Add(s);
                    }
                }
                if (FormworkDebugLog.Enabled)
                {
                    double volAfter = clipped.Sum(s => s?.Volume ?? 0);
                    if (clipped.Count == 0)
                        FormworkDebugLog.Log(
                            $"  [SectionBoxClip] E{elem.Id.IntValue()} src='{src.SourceName}' " +
                            $"切断ボックス外 → スキップ (vol_before={volBefore:F3}ft³)");
                    else
                        FormworkDebugLog.Log(
                            $"  [SectionBoxClip] E{elem.Id.IntValue()} src='{src.SourceName}' " +
                            $"クリップ完了 vol {volBefore:F3}→{volAfter:F3}ft³ clipped={clipped.Count}");
                }
                solids = clipped;
            }

            if (solids.Count == 0)
            {
                if (isLinkSweep && FormworkDebugLog.Enabled)
                {
                    FormworkDebugLog.Log(
                        $"  [LinkSweepDiag] Pass1-ABORT E{elem.Id.IntValue()} src='{src.SourceName}' " +
                        $"transformedSolids=0 → skipped (no retry for linked WallSweep)");
                }
                return null;
            }
            if (isLinkSweep && FormworkDebugLog.Enabled)
            {
                double txVol = 0;
                foreach (var s in solids) txVol += s?.Volume ?? 0;
                FormworkDebugLog.Log(
                    $"  [LinkSweepDiag] Pass1-postXform E{elem.Id.IntValue()} src='{src.SourceName}' " +
                    $"transformedSolids={solids.Count} txVol={txVol:F4}ft³");
            }

            Solid unioned = SolidUnionProcessor.Union(solids);
            var finalSolids = unioned != null ? new List<Solid> { unioned } : solids;

            var (minZ, maxZ) = FaceClassifier.GetZRange(finalSolids);
            var faceInfos = FaceClassifier.ClassifyAll(finalSolids, minZ, maxZ);

            // 基礎以外の最下面は FormworkRequired として扱う（梁底・柱底等は形枠必要）
            // 他要素との接触は ContactDetector が個別に検出する
            bool isFoundation = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Foundation;
            bool isSlab = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Slab;
            bool isWall = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Wall;
            bool isRoof = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Roof;

            if (!isFoundation)
            {
                foreach (var fi in faceInfos)
                {
                    if (fi.FaceType == FaceType.DeductedBottom)
                        fi.FaceType = FaceType.FormworkRequired;
                }
            }

            // 床(スラブ)・構造基礎・壁の天端は常に型枠不要: 段差・凹み・リビール溝底を含め
            // 全ての上向き水平面を除外する (上側はコンクリート打設時に開放されているため)
            // 屋根は勾配を持つ場合があるため、勾配付き上向き面 (nz>0.1) も上面とみなして除外する
            if (isSlab || isFoundation || isWall || isRoof)
            {
                double upTol = isRoof ? 0.1 : 0.99;
                foreach (var fi in faceInfos)
                {
                    if (fi.Normal != null && fi.Normal.Z > upTol &&
                        fi.FaceType == FaceType.FormworkRequired)
                    {
                        fi.FaceType = FaceType.DeductedTop;
                    }
                }
            }

            // 床(スラブ)のボイド/開口の内部縦面は、以前は「打設されない空間」として
            // 型枠不要扱いにしていたが、これは誤り。ボイドや貫通開口を区切るためには
            // コンクリート打設時に型枠が必要 (壁がボイドを貫通している場合の控除は
            // ContactFaceDetector が DeductedContact として処理する)。
            // 旧 ExcludeSlabInteriorVoidFaces の呼び出しは廃止。

            // 壁の斜めの天端: 斜面の水平射影の短辺 (壁厚方向の幅) が
            // しきい値 (既定 30mm) 以上なら型枠不要 (天端扱い) として控除する。
            // 小さな面取り (< 30mm) は型枠必要のまま残す。
            //
            // ⚠️ 追加条件: 面中心 z が壁の最高点近く (上位 15% 領域) に限定する。
            // これがないと leaning wall (傾斜壁) の主側面 (nz が少し正) まで誤って
            // 天端扱いされる問題があった (壁 E4551240 等)。
            if (isWall && _settings.SlopedWallTopWidthThresholdMm > 0)
            {
                double thresholdFeet = UnitUtils.ConvertToInternalUnits(
                    _settings.SlopedWallTopWidthThresholdMm, UnitTypeId.Millimeters);
                double wallHeight = maxZ - minZ;
                double topZThreshold = minZ + wallHeight * 0.85;
                foreach (var fi in faceInfos)
                {
                    if (fi.FaceType != FaceType.FormworkRequired) continue;
                    if (fi.Normal == null) continue;
                    double nz = fi.Normal.Z;
                    // 上向き斜面のみ対象 (水平面は既存ルールで処理済み)
                    if (nz <= 0.1 || nz >= 0.99) continue;

                    // 面中心 z が壁の上位 15% 以内にあることを要求
                    double faceCenterZ = TryGetFaceCenterZ(fi.Face);
                    if (double.IsNaN(faceCenterZ) || faceCenterZ < topZThreshold)
                    {
                        FormworkDebugLog.Log(
                            $"  [SlopedWallTop-Skip] E{elem.Id.IntValue()} face nz={nz:F3} " +
                            $"centerZmm={faceCenterZ * 304.8:F0} topThr={topZThreshold * 304.8:F0} " +
                            $"(壁中央〜下の側面のためスキップ)");
                        continue;
                    }

                    double widthFeet = ComputeFaceHorizontalShortDim(fi.Face);
                    if (widthFeet >= thresholdFeet)
                    {
                        fi.FaceType = FaceType.DeductedTop;
                        FormworkDebugLog.Log(
                            $"  [SlopedWallTop] E{elem.Id.IntValue()} face nz={nz:F3} " +
                            $"widthMm={widthFeet * 304.8:F1} centerZmm={faceCenterZ * 304.8:F0} → DeductedTop");
                    }
                }
            }

            // BoundingBox を世界座標で計算する (ホスト・リンクとも統一)
            BoundingBoxXYZ bb = ComputeWorldBoundingBox(finalSolids);

            if (isLinkSweep && FormworkDebugLog.Enabled)
            {
                int req = 0, top = 0, bot = 0, con = 0, inc = 0, err = 0;
                foreach (var f in faceInfos)
                {
                    switch (f.FaceType)
                    {
                        case FaceType.FormworkRequired: req++; break;
                        case FaceType.DeductedTop: top++; break;
                        case FaceType.DeductedBottom: bot++; break;
                        case FaceType.DeductedContact: con++; break;
                        case FaceType.Inclined: inc++; break;
                        case FaceType.Error: err++; break;
                    }
                }
                FormworkDebugLog.Log(
                    $"  [LinkSweepDiag] Pass1-classified E{elem.Id.IntValue()} src='{src.SourceName}' " +
                    $"surrogateId={src.SurrogateId} totalFaces={faceInfos.Count} " +
                    $"Req={req} Top={top} Bot={bot} Con={con} Inc={inc} Err={err}");
            }

            return new ContactFaceDetector.ElementFacesContext
            {
                ElementId = src.SurrogateId,
                Category = ElementCollector.ToCategoryGroup(elem),
                CategoryName = elem.Category?.Name ?? string.Empty,
                BB = bb,
                Faces = faceInfos,
                IsWallSweep = elem is WallSweep,
            };
        }

        /// <summary>
        /// 床(スラブ)のボイド/開口の内部にある面 (縦面・下向き面) を型枠不要として除外する。
        ///
        /// 判定方法: 床の最大の上向き水平面 (主たる天端) を取得し、その面の XY 投影に
        /// 対する Project 結果を用いて「面外向きの近傍点が床の天端範囲内にあるか」を判定する。
        ///
        ///   - 縦面: 面中心から法線方向 (外向き) に少し移動した点が天端面上にあれば、
        ///     その縦面は天端の下に「潜り込む」位置にある = ボイド/開口の内側 → 除外
        ///   - 下向き面: 面中心が天端面の真下にあって、かつ床の最下面 (DeductedBottom 候補)
        ///     より高い位置にあれば、ボイド/開口の天井部分 → 除外
        ///
        /// 完全な貫通開口の場合、天端面の Project が貫通穴の真ん中で失敗するが、
        /// その場合は床の XY バウンディングボックスとの距離でフォールバック判定する。
        /// </summary>
        private static void ExcludeSlabInteriorVoidFaces(
            Element elem,
            List<FaceClassifier.FaceInfo> faceInfos,
            List<Solid> finalSolids)
        {
            // 最大の上向き水平面 (天端の代表) を取得
            Face topFace = null;
            double topArea = 0;
            foreach (var s in finalSolids)
            {
                if (s == null) continue;
                foreach (Face f in s.Faces)
                {
                    XYZ n = FaceClassifier.GetFaceNormal(f);
                    if (n == null || n.Z <= 0.99) continue;
                    double a = 0;
                    try { a = f.Area; } catch { }
                    if (a > topArea)
                    {
                        topArea = a;
                        topFace = f;
                    }
                }
            }

            // 床の XY バウンディングボックスをエッジから計算 (天端 Project の補助)
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            foreach (var s in finalSolids)
            {
                if (s == null) continue;
                foreach (Edge e in s.Edges)
                {
                    var c = e.AsCurve();
                    if (c == null) continue;
                    for (int i = 0; i <= 1; i++)
                    {
                        var p = c.GetEndPoint(i);
                        if (p.X < minX) minX = p.X;
                        if (p.X > maxX) maxX = p.X;
                        if (p.Y < minY) minY = p.Y;
                        if (p.Y > maxY) maxY = p.Y;
                    }
                }
            }
            if (minX == double.MaxValue) return;

            // 外周からのマージン: 外周縦面は外向きオフセット後に BBox 外に出るが、
            // 開口縦面は BBox 内部に留まる。マージンは外周面の誤検出防止用 (50mm)。
            double margin = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Millimeters);
            // 縦面を外向きにオフセットする距離 (外周面が BBox 境界を少し超えるよう)
            double normalOffset = UnitUtils.ConvertToInternalUnits(10, UnitTypeId.Millimeters);

            FormworkDebugLog.Log(
                $"  [SlabVoidExclude] E{elem.Id.IntValue()} start: " +
                $"topFace={(topFace != null ? "found" : "null")} " +
                $"BBox X=[{minX * 304.8:F0},{maxX * 304.8:F0}]mm " +
                $"Y=[{minY * 304.8:F0},{maxY * 304.8:F0}]mm " +
                $"margin={margin * 304.8:F0}mm offset={normalOffset * 304.8:F0}mm");

            int excludedCount = 0;
            foreach (var fi in faceInfos)
            {
                if (fi.FaceType != FaceType.FormworkRequired) continue;
                if (fi.Normal == null) continue;

                double nz = fi.Normal.Z;
                // 縦面のみ処理する。下向き面 (床底面) は常に型枠必要のため除外しない。
                // 理由: 下向き面の中心は定義上 BBox 内部にあるため、BBox 判定が常に TRUE に
                //       なってしまい、床底面まで型枠不要と誤判定される。
                bool isVertical = Math.Abs(nz) < 0.1;
                if (!isVertical) continue;

                // 面中心を取得
                XYZ center = null;
                try
                {
                    var bb = fi.Face.GetBoundingBox();
                    if (bb != null)
                        center = fi.Face.Evaluate((bb.Min + bb.Max) * 0.5);
                }
                catch { }
                if (center == null) continue;

                // 縦面を外向きに少し移動した位置で判定する:
                //   - 外周縦面: 外向きオフセット後に BBox 外へ出る → 除外されない
                //   - 開口縦面: 外向きオフセット後も BBox 内部に留まる → 除外される
                double checkX = center.X + fi.Normal.X * normalOffset;
                double checkY = center.Y + fi.Normal.Y * normalOffset;

                // 判定: 外向きオフセット後の XY 位置が床 BBox の内側 (マージン分) に収まるか
                bool insideByBBox =
                    checkX > minX + margin && checkX < maxX - margin &&
                    checkY > minY + margin && checkY < maxY - margin;

                // 補助判定: 天端面への Project が成功 = XY 位置が床材料の上にある (開口の上ではない)
                // ただし貫通開口では topFace に穴があるため Project が失敗する場合もある。
                // insideByBBox が TRUE であれば補助判定不要。FALSE の場合のみ使用。
                bool insideByTopProject = false;
                if (!insideByBBox && topFace != null)
                {
                    try
                    {
                        var probe = new XYZ(checkX, checkY, center.Z + 1.0);
                        var ir = topFace.Project(probe);
                        if (ir != null) insideByTopProject = true;
                    }
                    catch { }
                }

                if (insideByBBox || insideByTopProject)
                {
                    FormworkDebugLog.Log(
                        $"  [SlabVoidExclude]   exclude face n=({fi.Normal.X:F0},{fi.Normal.Y:F0},{fi.Normal.Z:F0}) " +
                        $"center=({center.X * 304.8:F0},{center.Y * 304.8:F0})mm " +
                        $"check=({checkX * 304.8:F0},{checkY * 304.8:F0})mm " +
                        $"bbox={insideByBBox} proj={insideByTopProject}");
                    fi.FaceType = FaceType.DeductedTop;
                    excludedCount++;
                }
            }

            FormworkDebugLog.Log(
                $"  [SlabVoidExclude] E{elem.Id.IntValue()} done: excluded={excludedCount}");
        }

        /// <summary>
        /// 面の中心点 (UV bounding box の中点を Evaluate) の Z 座標を取得する。
        /// 取得失敗時は double.NaN を返す。
        /// </summary>
        private static double TryGetFaceCenterZ(Face face)
        {
            if (face == null) return double.NaN;
            try
            {
                var bb = face.GetBoundingBox();
                if (bb == null) return double.NaN;
                var midUV = (bb.Min + bb.Max) * 0.5;
                var p = face.Evaluate(midUV);
                return p?.Z ?? double.NaN;
            }
            catch { return double.NaN; }
        }

        /// <summary>
        /// 変換適用後の Solid 群からワールド座標系の BoundingBox を計算する。
        /// </summary>
        private static BoundingBoxXYZ ComputeWorldBoundingBox(IEnumerable<Solid> solids)
        {
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            bool any = false;
            foreach (var s in solids)
            {
                if (s == null) continue;
                foreach (Edge e in s.Edges)
                {
                    var c = e.AsCurve();
                    if (c == null) continue;
                    for (int i = 0; i <= 1; i++)
                    {
                        var p = c.GetEndPoint(i);
                        if (p.X < minX) minX = p.X;
                        if (p.Y < minY) minY = p.Y;
                        if (p.Z < minZ) minZ = p.Z;
                        if (p.X > maxX) maxX = p.X;
                        if (p.Y > maxY) maxY = p.Y;
                        if (p.Z > maxZ) maxZ = p.Z;
                        any = true;
                    }
                }
            }
            if (!any) return null;
            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ),
            };
        }

        private ElementResult BuildElementResult(
            ElementSource src,
            ContactFaceDetector.ElementFacesContext ctx,
            Dictionary<int, OpeningProcessor.OpeningDelta> openingMap)
        {
            var elem = src.Element;
            double formwork = 0, dedTop = 0, dedBottom = 0, dedContact = 0, inclined = 0;

            foreach (var fi in ctx.Faces)
            {
                double aM2 = FeetSqToM2(fi.Area);
                switch (fi.FaceType)
                {
                    case FaceType.FormworkRequired:
                        {
                            // 部分接触の面積を控除する。
                            // Naive sum (ContactArea を単純合計) は接触領域同士が重なって
                            // いると過大評価され、面積が 0 にクランプされてしまう。
                            // Clipper を使った 2D 矩形差分で正確な残面積を計算する。
                            string clipperStatus;
                            double effectiveFeetSq = ComputeAndSetEffectiveArea(fi, out clipperStatus);

                            // 部分接触の詳細をログ (問題追跡用)
                            if (fi.PartialContacts.Count > 0 && FormworkDebugLog.Enabled)
                            {
                                var sb = new System.Text.StringBuilder();
                                sb.Append($"  [FaceDiag] E{ctx.ElementId} face[{ctx.Faces.IndexOf(fi)}] ");
                                sb.Append($"area={fi.Area:F4} partials={fi.PartialContacts.Count} [");
                                foreach (var pc in fi.PartialContacts)
                                    sb.Append($" E{pc.OtherElementId}:a={pc.ContactArea:F4}");
                                sb.Append($" ] {clipperStatus} eff={effectiveFeetSq:F4}");
                                FormworkDebugLog.Log(sb.ToString());
                            }

                            formwork += fi.EffectiveAreaM2;
                            // 部分接触分も控除面積として集計に加える
                            double partialDed = Math.Max(0, fi.Area - effectiveFeetSq);
                            dedContact += FeetSqToM2(partialDed);
                        }
                        break;
                    case FaceType.DeductedTop: dedTop += aM2; break;
                    case FaceType.DeductedBottom: dedBottom += aM2; break;
                    case FaceType.DeductedContact: dedContact += aM2; break;
                    case FaceType.Inclined: inclined += aM2; break;
                }
            }

            double openingDeducted = 0;
            double openingAdded = 0;
            if (openingMap.TryGetValue(ctx.ElementId, out var od))
            {
                openingDeducted = FeetSqToM2(od.DeductedArea);
                openingAdded = FeetSqToM2(od.AddedEdgeArea);
                formwork = Math.Max(0, formwork - openingDeducted + openingAdded);
            }

            bool hasPartialContact = false;
            foreach (var fi in ctx.Faces)
            {
                if (fi.FaceType == FaceType.FormworkRequired && fi.PartialContacts.Count > 0)
                { hasPartialContact = true; break; }
            }

            if (src.IsLinked && elem is WallSweep && FormworkDebugLog.Enabled)
            {
                FormworkDebugLog.Log(
                    $"  [LinkSweepDiag] Pass3-result E{elem.Id.IntValue()} src='{src.SourceName}' " +
                    $"FormworkArea={formwork:F4}m² dedTop={dedTop:F4} dedBot={dedBottom:F4} " +
                    $"dedCon={dedContact:F4} openingDed={openingDeducted:F4} openingAdd={openingAdded:F4} " +
                    $"hasPartial={hasPartialContact}");
            }

            return new ElementResult
            {
                ElementId = ctx.ElementId,
                ElementName = elem.Name ?? string.Empty,
                Category = ElementCollector.ToCategoryGroup(elem),
                CategoryName = elem.Category?.Name ?? string.Empty,
                SourceName = src.SourceName,
                Zone = _settings.GroupByZone
                    ? NormalizeParamValue(ElementCollector.GetParameterString(elem, _settings.ZoneParameterName))
                    : string.Empty,
                FormworkType = _settings.GroupByFormworkType
                    ? NormalizeParamValue(ElementCollector.GetParameterString(elem, _settings.FormworkTypeParameterName))
                    : string.Empty,
                FormworkArea = formwork,
                DeductedTopArea = dedTop,
                DeductedBottomArea = dedBottom,
                DeductedContactArea = dedContact,
                InclinedArea = inclined,
                OpeningAreaDeducted = openingDeducted,
                OpeningEdgeAreaAdded = openingAdded,
                HasPartialContact = hasPartialContact,
            };
        }

        /// <summary>
        /// 可視化のため、分類後の面を要素IDごとに提供する。
        /// 幾何学的な接触検出を使うので View3D は不要。
        /// </summary>
        internal static Dictionary<int, List<FaceClassifier.FaceInfo>> RecomputeFaces(
            Document doc,
            FormworkResult result,
            FormworkSettings settings)
        {
            var map = new Dictionary<int, List<FaceClassifier.FaceInfo>>();

            var contexts = new List<ContactFaceDetector.ElementFacesContext>();
            var registry = result.SourceRegistry as ElementSourceRegistry;

            foreach (var er in result.ElementResults)
            {
                ElementSource src = registry?.Get(er.ElementId);
                Element elem = src?.Element;
                Transform xform = src?.Transform ?? Transform.Identity;
                if (elem == null)
                {
                    // フォールバック: 旧来通り ID 解決 (リンクなし時の互換)
                    try { elem = doc.GetElement(new ElementId(er.ElementId)); } catch { }
                }
                if (elem == null) continue;

                var solids = SolidUnionProcessor.GetSolids(elem, xform);
                if (solids.Count == 0) continue;
                var unioned = SolidUnionProcessor.Union(solids);
                var final = unioned != null ? new List<Solid> { unioned } : solids;
                var (minZ, maxZ) = FaceClassifier.GetZRange(final);
                var faces = FaceClassifier.ClassifyAll(final, minZ, maxZ);

                bool isFoundation = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Foundation;
                bool isSlab = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Slab;
                bool isWall = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Wall;
                bool isRoof = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Roof;

                if (!isFoundation)
                {
                    foreach (var f in faces)
                    {
                        if (f.FaceType == FaceType.DeductedBottom)
                            f.FaceType = FaceType.FormworkRequired;
                    }
                }

                if (isSlab || isFoundation || isWall || isRoof)
                {
                    double upTol = isRoof ? 0.1 : 0.99;
                    foreach (var f in faces)
                    {
                        if (f.Normal != null && f.Normal.Z > upTol &&
                            f.FaceType == FaceType.FormworkRequired)
                        {
                            f.FaceType = FaceType.DeductedTop;
                        }
                    }
                }

                // 壁の斜めの天端も控除 (BuildElementResult と同じロジック)
                if (isWall && settings.SlopedWallTopWidthThresholdMm > 0)
                {
                    double thresholdFeet = UnitUtils.ConvertToInternalUnits(
                        settings.SlopedWallTopWidthThresholdMm, UnitTypeId.Millimeters);
                    double wallHeight = maxZ - minZ;
                    double topZThreshold = minZ + wallHeight * 0.85;
                    foreach (var f in faces)
                    {
                        if (f.FaceType != FaceType.FormworkRequired) continue;
                        if (f.Normal == null) continue;
                        double nz = f.Normal.Z;
                        if (nz <= 0.1 || nz >= 0.99) continue;
                        double faceCenterZ = TryGetFaceCenterZ(f.Face);
                        if (double.IsNaN(faceCenterZ) || faceCenterZ < topZThreshold) continue;
                        double widthFeet = ComputeFaceHorizontalShortDim(f.Face);
                        if (widthFeet >= thresholdFeet)
                            f.FaceType = FaceType.DeductedTop;
                    }
                }

                // スラブのボイド/開口内部の縦面除外は廃止 (ClassifyElementFaces と同じ理由)。
                // ボイドの内側面はコンクリート打設時の型枠が必要なため FormworkRequired のまま残す。

                BoundingBoxXYZ bb = ComputeWorldBoundingBox(final);

                contexts.Add(new ContactFaceDetector.ElementFacesContext
                {
                    ElementId = er.ElementId,
                    Category = ElementCollector.ToCategoryGroup(elem),
                    CategoryName = elem.Category?.Name ?? string.Empty,
                    BB = bb,
                    Faces = faces,
                    IsWallSweep = elem is WallSweep,
                });
            }

            ContactFaceDetector.RefineContactFaces(contexts);

            // 各面の有効面積を計算して FaceInfo に保存。
            // RecomputeFaces は新しい FaceInfo インスタンスを作るため、BuildElementResult
            // 側で設定した EffectiveAreaM2 はここには来ない。可視化で使うために再計算する。
            foreach (var ctx in contexts)
            {
                foreach (var fi in ctx.Faces)
                {
                    if (fi.FaceType == FaceType.FormworkRequired)
                    {
                        string _;
                        ComputeAndSetEffectiveArea(fi, out _);
                    }
                }
                map[ctx.ElementId] = ctx.Faces;
            }

            return map;
        }

        /// <summary>
        /// FormworkRequired 面の有効面積 (部分接触控除後) を計算し、fi.EffectiveAreaM2 に保存する。
        /// 戻り値は feet² 単位の有効面積 (ログ出力用)、out clipperStatus は診断文字列。
        ///
        /// 残面積が face 全体の 1% 未満になった場合は FaceType を DeductedContact に降格する
        /// (FormworkVisualizer はこれを見て DirectShape を作らない)。
        /// </summary>
        private static double ComputeAndSetEffectiveArea(
            FaceClassifier.FaceInfo fi, out string clipperStatus)
        {
            clipperStatus = "no-partial";
            double effectiveFeetSq = fi.Area;
            if (fi.PartialContacts.Count > 0)
            {
                var clip = PartialContactClipper.TryClip(fi);
                if (clip.Success && clip.Solids.Count > 0)
                {
                    double area = 0;
                    foreach (var s in clip.Solids)
                        area += s.Volume / PartialContactClipper.ThicknessFeet;
                    effectiveFeetSq = area;
                    clipperStatus = "clipper-OK";
                }
                else if (fi.PartialContacts.Count == 1)
                {
                    // 単一 partial の場合は ContactArea をそのまま信頼 (95% cap は不要)。
                    // 95% cap は複数 partial の naive sum 過大評価対策のため。
                    var pc = fi.PartialContacts[0];
                    effectiveFeetSq = Math.Max(0, fi.Area - pc.ContactArea);
                    clipperStatus = $"single-partial(partial={pc.ContactArea:F4}";
                    if (clip.FailReason != null) clipperStatus += $",clipper-FAIL={clip.FailReason}";
                    clipperStatus += ")";
                }
                else
                {
                    // 複数 partial: naive sum + 95% cap (重なり過大評価対策)
                    double partial = 0;
                    foreach (var pc in fi.PartialContacts) partial += pc.ContactArea;
                    double rawSum = partial;
                    partial = Math.Min(partial, fi.Area * 0.95);
                    effectiveFeetSq = Math.Max(0, fi.Area - partial);
                    clipperStatus = $"clipper-FAIL({clip.FailReason}) rawSum={rawSum:F4}";
                }
            }

            // 残面積が face 全体の 1% 未満、または部分接触合計が face の 95% 以上なら
            // 「ほぼ全面接触」として接触面に降格する。
            //
            // 例: スラブに梁が完全埋もれている場合、Partial が複数 (主梁面積 + 微小な
            // 残端) で合計が face にほぼ一致するケース。
            // 既存の 95% キャップ branch だと残面積 5% (≈ 2.7 ft² ≈ 0.25m²) が残り、
            // 不要な型枠 DirectShape として可視化されてしまう。
            //
            // 降格すると FormworkVisualizer は DirectShape を作らず、集計表は
            // formwork → dedContact に振替えるため総合計は維持される。
            if (fi.PartialContacts.Count > 0)
            {
                double partialSum = 0;
                foreach (var pc in fi.PartialContacts) partialSum += pc.ContactArea;
                bool tinyResidual = effectiveFeetSq < fi.Area * 0.01;
                bool nearFullCoverage = partialSum >= fi.Area * 0.95;
                if (tinyResidual || nearFullCoverage)
                {
                    effectiveFeetSq = 0;
                    fi.FaceType = FaceType.DeductedContact;
                    clipperStatus += $" demoted-to-contact(partialSum={partialSum:F4}/face={fi.Area:F4})";
                }
            }

            fi.EffectiveAreaM2 = FeetSqToM2(effectiveFeetSq);
            return effectiveFeetSq;
        }

        private static string NormalizeParamValue(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "未設定";
            return s.Trim();
        }

        private static double FeetSqToM2(double feetSq)
        {
            return UnitUtils.ConvertFromInternalUnits(feetSq, UnitTypeId.SquareMeters);
        }

        private static void Aggregate(FormworkResult result)
        {
            double total = 0, deducted = 0, inclined = 0;
            var byCat = new Dictionary<CategoryGroup, CategoryResult>();
            var byZone = new Dictionary<string, ZoneResult>();
            var byType = new Dictionary<string, FormworkTypeResult>();

            foreach (var er in result.ElementResults)
            {
                total += er.FormworkArea;
                deducted += er.DeductedTopArea + er.DeductedBottomArea + er.DeductedContactArea;
                inclined += er.InclinedArea;

                if (!byCat.TryGetValue(er.Category, out var c))
                {
                    c = new CategoryResult { Category = er.Category, CategoryName = er.CategoryName };
                    byCat[er.Category] = c;
                }
                c.FormworkArea += er.FormworkArea;
                c.DeductedArea += er.DeductedTopArea + er.DeductedBottomArea + er.DeductedContactArea;
                c.ElementCount++;

                if (!string.IsNullOrEmpty(er.Zone))
                {
                    if (!byZone.TryGetValue(er.Zone, out var z))
                    {
                        z = new ZoneResult { Zone = er.Zone };
                        byZone[er.Zone] = z;
                    }
                    z.FormworkArea += er.FormworkArea;
                    z.ElementCount++;
                }

                if (!string.IsNullOrEmpty(er.FormworkType))
                {
                    if (!byType.TryGetValue(er.FormworkType, out var t))
                    {
                        t = new FormworkTypeResult { FormworkType = er.FormworkType };
                        byType[er.FormworkType] = t;
                    }
                    t.FormworkArea += er.FormworkArea;
                    t.ElementCount++;
                }
            }

            result.TotalFormworkArea = total;
            result.TotalDeductedArea = deducted;
            result.InclinedFaceArea = inclined;
            result.CategoryResults = byCat.Values.OrderBy(c => c.Category).ToList();
            result.ZoneResults = byZone.Values.OrderBy(z => z.Zone).ToList();
            result.TypeResults = byType.Values.OrderBy(t => t.FormworkType).ToList();
        }
    }
}

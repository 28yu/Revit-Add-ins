using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Engine;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Output
{
    /// <summary>
    /// 型枠面を DirectShape として作成し、View Filter で色分けする。
    /// - 分類キー毎に DirectShapeType を作成（名前で識別可能）
    /// - 色付けは View Filter ベース（ユーザーが実行後に編集可能）
    /// - 元の躯体要素は 20% 透過 + RGB(94,94,94) で落ち着いて表示
    /// - ビュー名から日時を省略（同名が存在する場合は削除してから再作成）
    /// </summary>
    internal static class FormworkVisualizer
    {
        internal const string AnalysisViewName = "3D_型枠数量";
        internal const string AnalysisViewPrefix = "3D_型枠数量 - ";
        // 旧名 (以前のバージョンで作成されたビューを検出するためのフォールバック)
        internal const string LegacyAnalysisViewName = "型枠分析";
        internal const string LegacyAnalysisViewPrefix = "型枠分析 - ";
        internal const string Legacy2AnalysisViewName = "型枠数量算出";
        internal const string Legacy2AnalysisViewPrefix = "型枠数量算出 - ";
        // さらに古いバージョン ("**型枠：SourceViewName" 形式)
        internal const string Legacy3AnalysisViewPrefix = "**型枠：";

        /// <summary>
        /// ソースビュー名からこのビューに対応する解析ビュー名を生成する。
        /// ソース名なし: "3D_型枠数量"
        /// ソース名あり: "3D_型枠数量 - {sourceViewName}"
        /// </summary>
        internal static string BuildAnalysisViewName(string sourceViewName)
        {
            if (string.IsNullOrEmpty(sourceViewName)) return AnalysisViewName;
            return AnalysisViewPrefix + sourceViewName;
        }

        /// <summary>名前が型枠分析ビュー (新旧全パターン) のパターンに一致するか判定する。
        /// 注: Legacy3 ("**型枠：") は当アドインの管理対象ではない別系統の旧ビューのため
        /// ここには含めない（含めると型枠数量集計シートに余計なビューが貼り付けられる）。
        /// 非表示フィルタの除外目的では HideAllFormworkShapesInOtherViews 内で別途扱う。
        /// </summary>
        internal static bool IsAnalysisViewName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            return name == AnalysisViewName
                || name.StartsWith(AnalysisViewPrefix)
                || name == LegacyAnalysisViewName
                || name.StartsWith(LegacyAnalysisViewPrefix)
                || name == Legacy2AnalysisViewName
                || name.StartsWith(Legacy2AnalysisViewPrefix);
        }
        internal class VisualizerResult
        {
            public View3D AnalysisView;
            public List<ElementId> CreatedShapeIds = new List<ElementId>();
            public Dictionary<string, (byte R, byte G, byte B)> KeyColors
                = new Dictionary<string, (byte, byte, byte)>();
            // このソースビューの DirectShape を識別するマーカー値
            public string MarkerValue;
        }

        private static readonly Dictionary<CategoryGroup, (byte R, byte G, byte B)> _categoryColors
            = new Dictionary<CategoryGroup, (byte, byte, byte)>
            {
                { CategoryGroup.Column,     (60, 120, 240) },
                { CategoryGroup.Beam,       (60, 200, 100) },
                { CategoryGroup.Wall,       (255, 160, 60) },
                { CategoryGroup.Slab,       (250, 220, 60) },
                { CategoryGroup.Foundation, (160, 100, 220) },
                { CategoryGroup.Stairs,     (60, 200, 220) },
                { CategoryGroup.Roof,       (220, 100, 140) },
                { CategoryGroup.Other,      (180, 180, 180) },
            };

        // 型枠不要として除外された要素のオレンジ系統色 (躯体グレー・型枠色と区別)
        private static readonly (byte R, byte G, byte B) _excludedColor = (255, 145, 30);

        private static readonly (byte R, byte G, byte B)[] _autoPalette = new (byte, byte, byte)[]
        {
            (230, 80, 80),   (80, 160, 230), (60, 200, 100), (240, 200, 60),
            (180, 100, 220), (255, 130, 170), (80, 210, 210), (255, 160, 60),
            (170, 220, 80), (60, 170, 170), (220, 100, 200), (120, 120, 220),
        };

        internal static VisualizerResult CreateVisualization(
            Document doc,
            FormworkResult result,
            FormworkSettings settings,
            View sourceView = null,
            bool updateMode = false)
        {
            var vr = new VisualizerResult();

            // ソースビュー名を含む解析ビュー名を生成 (ソース毎に独立した分析ビューを作成)
            string sourceViewName = sourceView?.Name;
            string analysisName = BuildAnalysisViewName(sourceViewName);
            string markerValue = FormworkParameterManager.MarkerValue;
            // [8] ホストモデルの表示名 (実際のファイル名 or タイトル)
            string hostDisplayName = GetDocumentDisplayName(doc);

            // 更新モード: 既存の分析ビューを検索して再利用する (シート上のビューポート参照を保つ)。
            View3D view = null;
            bool reusedView = false;
            if (updateMode)
            {
                view = FormworkOutputFinder.FindAnalysisView(doc, sourceViewName);
                if (view != null)
                {
                    reusedView = true;
                    FormworkDebugLog.Log(
                        $"  [Visual] update mode: reusing existing analysis view Id={view.Id.IntValue()} Name='{view.Name}'");
                }
            }

            if (!reusedView)
            {
                // 既存の同名ビューを削除 (このソースビューのもののみ)
                DeleteViewByName(doc, analysisName);
            }

            // 既存の型枠 DirectShape を削除。
            // ソースビュー名が指定されている場合は、そのビュー由来の DirectShape のみ削除し、
            // 他のソースビューで作成した DirectShape はそのまま残す（[2] 機能）。
            // ソースビュー名が空 (互換モード) の場合は全 DirectShape を削除する（従来動作）。
            CleanupExistingFormworkShapes(doc, sourceViewName);

            if (!reusedView)
            {
                // ソースが 3D ビューの場合は View.Duplicate でソースビューの全設定を複製する。
                // これにより以下が自動的に継承される:
                //   - V/G → Revitリンク タブの「表示設定」(ByHostView / ByLinkedView)
                //   - 各リンクに割り当てられた "Linked View Id"
                //   - カテゴリ非表示、フィルタ、ワークセット可視性、セクションボックス等
                // Revit 2022/2023 では GetLinkOverrides/SetLinkOverrides API が無いため、
                // Duplicate が ByLinkedView 設定を継承する唯一の確実な方法。
                view = null;
                if (sourceView is View3D srcV3DDup)
                {
                    try
                    {
                        var dupId = srcV3DDup.Duplicate(ViewDuplicateOption.Duplicate);
                        if (dupId != null && dupId != ElementId.InvalidElementId)
                        {
                            view = doc.GetElement(dupId) as View3D;
                            FormworkDebugLog.Log(
                                $"  [Visual] view duplicated from source='{srcV3DDup.Name}' newId={dupId.IntValue()}");
                        }
                    }
                    catch (Exception exDup)
                    {
                        FormworkDebugLog.Log($"  [Visual] Duplicate EX: {exDup.Message} → CreateIsometric fallback");
                    }
                }

                if (view == null)
                {
                    var view3DType = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewFamilyType))
                        .Cast<ViewFamilyType>()
                        .FirstOrDefault(v => v.ViewFamily == ViewFamily.ThreeDimensional);
                    if (view3DType == null) return vr;

                    view = View3D.CreateIsometric(doc, view3DType.Id);
                }

                try { view.Name = analysisName; } catch { }
            }

            // 出力タグを書き込む (リネーム耐性のための識別子)
            try
            {
                FormworkParameterManager.SetOutputTag(
                    view,
                    FormworkParameterManager.OutputKindAnalysisView,
                    sourceViewName ?? string.Empty);
            }
            catch (Exception ex) { FormworkDebugLog.Log($"  [Visual] SetOutputTag EX: {ex.Message}"); }

            vr.AnalysisView = view;
            vr.MarkerValue = markerValue;

            if (!reusedView)
            {
                // ビューテンプレートを解除する（テンプレートが適用されていると
                // カテゴリ/フィルタ設定の変更がブロックされる場合がある）。
                try
                {
                    if (view.ViewTemplateId != null && view.ViewTemplateId != ElementId.InvalidElementId)
                    {
                        view.ViewTemplateId = ElementId.InvalidElementId;
                        FormworkDebugLog.Log("  [Visual] view template detached");
                    }
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Visual] detach template EX: {ex.Message}");
                }

                // 表示スタイルを Shading に明示設定 (Realistic では OverrideGraphicSettings の
                // 透過・色オーバーライドが期待通りに反映されない場合があるため)。
                try { view.DisplayStyle = DisplayStyle.Shading; }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Visual] DisplayStyle set EX: {ex.Message}");
                }
                try { view.DetailLevel = ViewDetailLevel.Fine; } catch { }

                // ディシプリンを Coordination (調整) に設定して全カテゴリ (構造・電気等) を
                // 等しく表示できるようにする。型枠 DS のカテゴリ (NurseCallDevices) は
                // Electrical discipline 所属のため、Structural のままだとデフォルトで
                // 表示されないリスクがある。
                try
                {
                    var disciplineParam = view.get_Parameter(BuiltInParameter.VIEW_DISCIPLINE);
                    if (disciplineParam != null && !disciplineParam.IsReadOnly)
                    {
                        // Coordination = 4095 (全ディシプリンビット ON)
                        disciplineParam.Set(4095);
                        FormworkDebugLog.Log("  [Visual] View discipline set to Coordination");
                    }
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Visual] discipline set EX: {ex.Message}");
                }

                // 尺度を 1/200 に設定
                try { view.Scale = 200; }
                catch (Exception ex) { FormworkDebugLog.Log($"  [Visual] Scale set EX: {ex.Message}"); }

                // 視点をビューキューブの青い角（右前上からのアイソメトリック）に設定。
                // ソースが 3D ビューの場合はその視点を継承し、ソースと同じ見え方にする。
                if (sourceView is View3D srcV3DForOrient)
                    CopyOrientationIfPossible(srcV3DForOrient, view);
                else
                    SetIsometricOrientation(view);
            }

            if (!reusedView)
            {
                // 視点を固定アイソメトリックに設定し、ロックする（ユーザーが誤って回転しないよう）
                try { view.SaveOrientationAndLock(); }
                catch (Exception ex) { FormworkDebugLog.Log($"  [Visual] SaveOrientationAndLock EX: {ex.Message}"); }
            }

            // 接触検出込みで面を再計算（幾何学的検査なので rayView 不要）
            var facesByElement = FormworkCalcEngine.RecomputeFaces(doc, result, settings);

            // [4] ソースが 3D ビューなら、後段でソースフィルタをコピーする必要があるかを記録する。
            // フィルタは型枠フィルタ適用後に追加して優先度を低くする（型枠フィルタが上位）。
            bool shouldCopySourceFilters = false;

            FormworkDebugLog.Log(
                $"  [Visual] CreateVisualization: sourceView='{sourceView?.Name}' " +
                $"type={sourceView?.GetType().Name} reusedView={reusedView} " +
                $"target='{view?.Name}'");

            if (!reusedView)
            {
                // [3] ソースが 3D ビューの場合 (CurrentView / SelectedViews モード) は、
                // そのソースビューの切断ボックス (位置・有効/無効) をそのまま継承する。
                // ソースに切断ボックスが無い場合は型枠要素の BoundingBox から自動算出することで
                // 分析ビューが全体表示にならないようにする。
                // EntireProject 等は要素 BoundingBox から自動算出。
                // 更新モードでは既存ビューのセクションボックスを保持する。
                bool useSourceSectionBox = sourceView is View3D &&
                    (settings.Scope == CalculationScope.CurrentView ||
                     settings.Scope == CalculationScope.SelectedViews);
                if (useSourceSectionBox)
                {
                    var srcV3D = (View3D)sourceView;
                    if (srcV3D.IsSectionBoxActive)
                    {
                        // ソースに切断ボックスあり → コピーを試みる。
                        // コピー失敗時は要素BBoxからフォールバック。
                        bool copied = CopySectionBoxFromSource(srcV3D, view);
                        if (!copied)
                        {
                            FormworkDebugLog.Log("  [Visual] SectionBox copy failed → fallback to EnableSectionBox");
                            EnableSectionBox(doc, view, result);
                        }
                    }
                    else
                    {
                        // ソースに切断ボックスなし → 型枠要素の BoundingBox から切断ボックスを算出して設定。
                        // これにより解析ビューがモデル全体表示にならず、型枠エリアにフォーカスされる。
                        // EnableSectionBox は要素BBoxが取得できない場合は切断ボックスを有効化しないため安全。
                        FormworkDebugLog.Log(
                            "  [Visual] source IsSectionBoxActive=False → compute section box from element BBoxes");
                        EnableSectionBox(doc, view, result);
                    }
                }
                else
                    EnableSectionBox(doc, view, result);

                // [4] ソースが 3D ビューならその V/G 上書き (カテゴリの表示/色) を継承する。
                // フィルタのコピーは ApplyColorFilters 後に実施し、型枠フィルタを高優先度に保つ。
                // ソースが 3D ビューでない場合は、ノイズになる切断ボックスのアウトラインと
                // レベル線のみ非表示化。
                if (sourceView is View3D)
                {
                    CopyCategoryVisibilityAndOverrides(doc, sourceView, view);
                    CopyWorksetVisibility(doc, sourceView, view);
                    shouldCopySourceFilters = true;
                }
                else
                    HideClutterCategories(view);
            }

            // 型枠 DS のカテゴリ (FormworkCategory) を必ず表示状態にする。
            // CopyCategoryVisibilityAndOverrides の後に実行することで、ソースビューが
            // 同カテゴリを非表示にしていても型枠 DirectShape が確実に表示される。
            // 更新モードでも防御的に実行 (idempotent)。
            try
            {
                var fwCatId = new ElementId(FormworkParameterManager.FormworkCategory);
                if (view.CanCategoryBeHidden(fwCatId))
                    view.SetCategoryHidden(fwCatId, false);
                FormworkDebugLog.Log("  [Visual] FormworkCategory set visible");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] FormworkCategory set EX: {ex.Message}");
            }

            // 元躯体要素を 20% 透過 + RGB(94,94,94) で表示
            // 更新モードでも新たに追加された要素にも適用する (idempotent)。
            ApplySourceElementAppearance(doc, view, result);

            // 分類キーに基づく色割当
            var keyAssignment = AssignColors(result, settings);

            // 控除面表示時は「控除面」キーを追加（単色グレー）
            if (settings.ShowDeductedFaces)
                keyAssignment["控除面"] = (180, 180, 180);

            // 除外要素がある場合、オレンジ系のキーを追加（View Filter で色付け、既定で非表示）
            if (result.ExcludedResults != null && result.ExcludedResults.Count > 0)
                keyAssignment[FormworkParameterManager.ExcludedGroupKey] = _excludedColor;

            vr.KeyColors = new Dictionary<string, (byte, byte, byte)>(keyAssignment);

            // 分類キー毎に DirectShapeType を作成（名前で識別可能にする）
            var typeIdByKey = new Dictionary<string, ElementId>();
            var catLabelByKey = new Dictionary<string, string>();
            foreach (var er in result.ElementResults)
            {
                string key = GetKey(er, settings);
                if (typeIdByKey.ContainsKey(key)) continue;

                string typeName = BuildTypeName(settings, key, er);
                var dsType = GetOrCreateDirectShapeType(doc, typeName);
                if (dsType != null)
                {
                    typeIdByKey[key] = dsType.Id;
                    catLabelByKey[key] = CategoryLabel(er.Category);
                }
            }

            // ソースビューに実際に表示されている要素IDを収集し、非表示要素のDirectShape作成を省く。
            // FilteredElementCollector(doc, viewId) はカテゴリ・ワークセット・フィルタ・要素個別非表示を
            // 全て反映した「視点上で見えている要素」のみを返す。
            // EntireProject スコープは全要素対象なのでフィルタしない。
            HashSet<int> visibleInSourceIds = null;
            if (sourceView != null && settings.Scope != CalculationScope.EntireProject)
            {
                try
                {
                    visibleInSourceIds = new FilteredElementCollector(doc, sourceView.Id)
                        .WhereElementIsNotElementType()
                        .ToElementIds()
                        .Select(eid => eid.IntValue())
                        .ToHashSet();
                    FormworkDebugLog.Log(
                        $"  [Visual] visibleInSourceIds: {visibleInSourceIds.Count} (scope={settings.Scope})");
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Visual] visibleInSourceIds EX: {ex.Message}");
                    visibleInSourceIds = null;
                }
            }

            var registryForVisibility = result.SourceRegistry as Engine.ElementSourceRegistry;

            int skippedInvisible = 0;
            foreach (var er in result.ElementResults)
            {
                // ソースビューで非表示のホスト要素はDirectShapeを作らない。
                // リンク要素の可視判定はElementCollector側(TryGetByLinkedViewId等)で実施済み。
                // FilteredElementCollector(doc, viewId)はリンク内の個別要素を返さないため
                // Visualizerレベルではリンク要素の可視判定はできない。
                if (visibleInSourceIds != null)
                {
                    var srcInfo = registryForVisibility?.Get(er.ElementId);
                    if (srcInfo != null && !srcInfo.IsLinked)
                    {
                        // ホスト要素のみ: ElementId でチェック
                        if (!visibleInSourceIds.Contains(er.ElementId))
                        {
                            skippedInvisible++;
                            continue;
                        }
                    }
                }

                if (!facesByElement.TryGetValue(er.ElementId, out var faces))
                {
                    // [LinkSweepDiag] facesByElement に該当 ID が無いと DirectShape が作られない
                    var regChk = result.SourceRegistry as Engine.ElementSourceRegistry;
                    var sChk = regChk?.Get(er.ElementId);
                    if (sChk != null && sChk.IsLinked && sChk.Element is WallSweep
                        && FormworkDebugLog.Enabled)
                    {
                        FormworkDebugLog.Log(
                            $"  [LinkSweepDiag] Visualizer-SKIP E{sChk.Element.Id.IntValue()} " +
                            $"src='{sChk.SourceName}' surrogateId={er.ElementId} " +
                            $"reason='facesByElement に該当 ID 無し (RecomputeFaces で取れていない)'");
                    }
                    continue;
                }

                string key = GetKey(er, settings);
                typeIdByKey.TryGetValue(key, out var typeId);
                catLabelByKey.TryGetValue(key, out var catLabel);
                catLabel = catLabel ?? CategoryLabel(er.Category);

                // 要素のレベル名を取得 (リンク要素はリンクドキュメント経由)
                string levelName = string.Empty;
                Engine.ElementSource srcForDiag = null;
                try
                {
                    var registry = result.SourceRegistry as Engine.ElementSourceRegistry;
                    var src = registry?.Get(er.ElementId);
                    srcForDiag = src;
                    var srcElem = src?.Element ?? doc.GetElement(new ElementId(er.ElementId));
                    levelName = Engine.ElementCollector.GetElementLevelName(srcElem);
                }
                catch { }

                bool isLinkSweepDiag =
                    srcForDiag != null && srcForDiag.IsLinked && srcForDiag.Element is WallSweep;
                int linkSweepDsCreated = 0;

                // 各 FormworkRequired 面の有効面積 (fi.EffectiveAreaM2) をスケーリング係数で
                // 按分して各 DirectShape に乗せる。スケーリングは「面の素の面積合計」と
                // 「実際の formwork 面積 (= 素の合計 - 開口控除 + 開口縁加算)」の比。
                //
                // 旧実装は openingDelta を最初の DS に一発で乗せていたが、開口控除が
                // 巨大な壁 (例: 工作物擁壁の大開口) では最初の DS の面積が負値 (例: -8.3m²)
                // になる不具合があった。スケーリングなら個々の DS は負にならず、合計も
                // er.FormworkArea にぴったり一致する。
                double sumEffectiveFaceAreasM2 = 0;
                foreach (var fi in faces)
                {
                    if (fi.FaceType == FaceType.FormworkRequired)
                        sumEffectiveFaceAreasM2 += fi.EffectiveAreaM2;
                }
                double areaScale = sumEffectiveFaceAreasM2 > 1e-9
                    ? er.FormworkArea / sumEffectiveFaceAreasM2
                    : 1.0;
                // スケール係数が負になる理屈は無いが防御的にクランプ
                if (areaScale < 0) areaScale = 0;
                // 合計一致保証用: 実際に DS パラメータに書き込んだ面積の合計と最初の DS の ID
                double assignedAreaM2 = 0;
                ElementId firstFormworkDsId = null;
                // DirectShape (Clipper 分割なら最初の 1 個) に割り当てる。
                // これにより、ユーザーが個々の DirectShape を選択しても正しい面積が表示される。

                foreach (var fi in faces)
                {
                    bool visible =
                        fi.FaceType == FaceType.FormworkRequired ||
                        (settings.ShowDeductedFaces && IsDeducted(fi.FaceType));
                    if (!visible) continue;

                    bool faceHasPartialContact =
                        fi.FaceType == FaceType.FormworkRequired && fi.PartialContacts.Count > 0;

                    // Phase 2: 部分接触がある面は矩形差分で厳密な形状を作る。
                    //   Clipper 成功 → 半透明オーバーライドは不要 (形状が既に正確)
                    //   Clipper 失敗 → 3D Boolean Difference fallback を試行
                    //                 (face が非矩形/notched でも対応可能)
                    //   両方失敗 → 従来通り面全体を作成 + 半透明オーバーライド
                    List<Solid> partSolids = null;
                    bool clipperUsed = false;
                    if (faceHasPartialContact)
                    {
                        var clip = PartialContactClipper.TryClip(fi);
                        if (clip.Success && clip.Solids.Count > 0)
                        {
                            partSolids = clip.Solids;
                            clipperUsed = true;
                        }
                        else
                        {
                            // Clipper 失敗時の 3D Boolean Difference fallback。
                            // 非矩形 face (interior-points / has-holes 等) や fully-clipped
                            // ケースでも、face 薄板から PartialContact 領域を Boolean 差分で
                            // くり抜くことで正確な形状を作れる。
                            Solid carved = TryBuildCarvedFaceSolid(fi);
                            if (carved != null)
                            {
                                partSolids = new List<Solid> { carved };
                                clipperUsed = true;
                                FormworkDebugLog.Log(
                                    $"Boolean diff fallback OK: E{er.ElementId} face area={fi.Area:F4} " +
                                    $"clipperReason={clip.FailReason}");
                            }
                            else
                            {
                                FormworkDebugLog.Log(
                                    $"Clipper fallback: E{er.ElementId} face type={fi.FaceType} " +
                                    $"reason={clip.FailReason}");
                            }
                        }
                    }

                    if (partSolids == null)
                    {
                        Solid thin = CreateThinSolidFromFace(fi);
                        if (thin == null)
                        {
                            FormworkDebugLog.Log(
                                $"  [Visual] SKIP E{er.ElementId} face area={fi.Area:F4} " +
                                $"type={fi.FaceType} (CreateThinSolid returned null)");
                            continue;
                        }
                        partSolids = new List<Solid> { thin };
                    }

                    bool isFormworkRequired = (fi.FaceType == FaceType.FormworkRequired);

                    // 同一面の DirectShape ピース間で面積を按分する: piece.Volume / sum を share とする。
                    // 単一ピース (Boolean fallback / no-partials 経路) は share=1.0。
                    // クリッパー分割で複数ピースになる場合に各ピースが幾何学的サイズに比例した
                    // 面積を持つようにし、「2 個目以降が 0m²」になる旧設計を解消する。
                    double pieceVolumeSum = 0;
                    if (isFormworkRequired)
                    {
                        foreach (var s in partSolids)
                        {
                            try { pieceVolumeSum += s?.Volume ?? 0; } catch { }
                        }
                    }

                    foreach (var solid in partSolids)
                    {
                        ElementId dsId = null;
                        double areaM2ThisDs = 0.0;
                        try
                        {
                            var catOst = new ElementId(FormworkParameterManager.FormworkCategory);
                            var ds = DirectShape.CreateElement(doc, catOst);
                            ds.ApplicationId = "Tools28";
                            ds.ApplicationDataId = "Formwork";
                            ds.SetShape(new GeometryObject[] { solid });

                            if (typeId != null && typeId != ElementId.InvalidElementId)
                            {
                                try { ds.SetTypeId(typeId); } catch { }
                            }

                            try
                            {
                                // 面単位の有効面積 × areaScale を、ピースの体積比で按分する。
                                // 開口調整 (OpeningAreaDeducted/OpeningEdgeAreaAdded) は areaScale に
                                // 既に織り込み済みのため個別加算は不要。
                                double areaM2 = 0.0;
                                if (isFormworkRequired)
                                {
                                    double pieceVol = 0;
                                    try { pieceVol = solid?.Volume ?? 0; } catch { }
                                    double pieceShare = pieceVolumeSum > 1e-12
                                        ? pieceVol / pieceVolumeSum
                                        : 1.0;
                                    areaM2 = fi.EffectiveAreaM2 * areaScale * pieceShare;
                                }
                                string filterKey = IsDeducted(fi.FaceType) ? "控除面" : key;
                                string srcDisplay = er.SourceName == ElementSourceRegistry.HostSourceName
                                    ? hostDisplayName : er.SourceName;
                                FormworkParameterManager.SetInstanceValues(
                                    ds, catLabel, levelName, filterKey, areaM2,
                                    faceHasPartialContact,
                                    sourceName: srcDisplay,
                                    sourceViewName: sourceViewName);
                                areaM2ThisDs = areaM2;
                            }
                            catch { }

                            dsId = ds.Id;
                        }
                        catch { continue; }

                        if (dsId != null)
                        {
                            vr.CreatedShapeIds.Add(dsId);
                            if (isLinkSweepDiag && isFormworkRequired) linkSweepDsCreated++;
                            if (isFormworkRequired)
                            {
                                assignedAreaM2 += areaM2ThisDs;
                                if (firstFormworkDsId == null) firstFormworkDsId = dsId;
                            }
                            // Clipper でクリップ済みなら形状が既に正確なので半透明不要。
                            // Clipper 失敗時のみ「一部控除されている」ことを視覚化するが、
                            // 部分接触の合計面積が面全体の 5% 以下なら誤差として無視し、
                            // 半透明にしない (ユーザー視点で「ほぼ全面が型枠」のため)。
                            if (faceHasPartialContact && !clipperUsed
                                && IsSignificantPartialContact(fi))
                            {
                                try { ApplyPartialContactOverride(view, dsId); } catch { }
                            }
                        }
                    }
                }

                // 合計一致の保証: CreateThinSolid 失敗等でスキップされた面の面積を
                // 最初の FormworkRequired DirectShape で補正し、Revit 集計表と Excel の
                // 合計値が常に一致するようにする。
                if (firstFormworkDsId != null)
                {
                    double discrepancy = er.FormworkArea - assignedAreaM2;
                    if (Math.Abs(discrepancy) > 1e-6)
                    {
                        try
                        {
                            var dsFix = doc.GetElement(firstFormworkDsId);
                            if (dsFix != null)
                            {
                                var pArea = dsFix.LookupParameter(FormworkParameterManager.ParamArea);
                                if (pArea != null && !pArea.IsReadOnly
                                    && pArea.StorageType == StorageType.Double)
                                {
                                    double curFt2 = pArea.AsDouble();
                                    double curM2 = UnitUtils.ConvertFromInternalUnits(
                                        curFt2, UnitTypeId.SquareMeters);
                                    double newM2 = curM2 + discrepancy;
                                    pArea.Set(UnitUtils.ConvertToInternalUnits(
                                        newM2, UnitTypeId.SquareMeters));
                                    FormworkDebugLog.Log(
                                        $"  [Visual] area reconcile E{er.ElementId}: " +
                                        $"assigned={assignedAreaM2:F4} expected={er.FormworkArea:F4} " +
                                        $"delta={discrepancy:F4}m²");
                                }
                            }
                        }
                        catch { }
                    }
                }

                if (isLinkSweepDiag && FormworkDebugLog.Enabled)
                {
                    FormworkDebugLog.Log(
                        $"  [LinkSweepDiag] Visualizer-done E{srcForDiag.Element.Id.IntValue()} " +
                        $"src='{srcForDiag.SourceName}' surrogateId={er.ElementId} " +
                        $"FormworkArea={er.FormworkArea:F4}m² faceCount={faces.Count} " +
                        $"createdFormworkDS={linkSweepDsCreated}");
                }
            }

            int formworkShapesCount = vr.CreatedShapeIds.Count;
            FormworkDebugLog.Log(
                $"  [Visual] formwork DirectShapes created: {formworkShapesCount} " +
                $"(elements={result.ElementResults.Count} skippedInvisible={skippedInvisible})");

            // 新しく作成した DirectShape のジオメトリを Revit に認識させるため Regenerate
            try
            {
                doc.Regenerate();
                FormworkDebugLog.Log("  [Visual] doc.Regenerate() called after formwork creation");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] doc.Regenerate() EX: {ex.Message}");
            }

            // 診断: 作成されたDirectShapeのカテゴリ別件数 + 先頭5個のBBox
            DiagnoseFormworkShapes(doc, vr.CreatedShapeIds, result);

            // 除外要素 (鉄骨・デッキスラブ) を別マーカーの DirectShape として作成（オレンジ）
            CreateExcludedShapes(doc, result, vr, sourceViewName);
            int excludedShapesCount = vr.CreatedShapeIds.Count - formworkShapesCount;
            FormworkDebugLog.Log(
                $"  [Visual] total DirectShapes: {vr.CreatedShapeIds.Count} " +
                $"(formwork={formworkShapesCount} excluded={excludedShapesCount})");

            // ワークセット可視性コピー後に型枠 DirectShape のワークセットを強制表示する。
            // ソースビューがホストのワークセットを非表示にしている場合、CopyWorksetVisibility で
            // その設定が引き継がれ DirectShape が見えなくなるのを防ぐ。
            if (doc.IsWorkshared)
                EnsureFormworkWorksetsVisible(doc, view, vr.CreatedShapeIds);

            // ソースビューから Duplicate で継承されたフィルタを一旦全て削除する。
            // Duplicate で引き継いだソースフィルタはフィルタリスト先頭（最高優先度）に位置するが、
            // 型枠フィルタが最高優先度でないとソースフィルタの visible=false 設定が
            // 型枠 DirectShape（OST_GenericModel）にも適用されて非表示になってしまう。
            // 削除したフィルタは後段 CopyFilterSettings で再追加し、型枠フィルタより低優先度にする。
            if (shouldCopySourceFilters)
            {
                try
                {
                    var inheritedFilters = view.GetFilters();
                    if (inheritedFilters != null && inheritedFilters.Count > 0)
                    {
                        foreach (var ifid in inheritedFilters)
                        {
                            try { view.RemoveFilter(ifid); } catch { }
                        }
                        FormworkDebugLog.Log(
                            $"  [Visual] cleared {inheritedFilters.Count} inherited filters " +
                            $"(will re-add after formwork filters for correct priority)");
                    }
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Visual] clearInheritedFilters EX: {ex.Message}");
                }
            }

            // View Filter による色分けを適用（個別要素オーバーライドは使わない）
            // ★ この後に追加されるフィルタ（ソースフィルタ）より高優先度になる。
            // 更新モードでは既存ビューのフィルタ設定 (ユーザーが手動で調整した色等を含む) を
            // 保持するため再適用しない。
            if (!reusedView)
            {
                try
                {
                    FormworkFilterManager.ApplyColorFilters(doc, view, keyAssignment);
                }
                catch { }
            }
            else
            {
                FormworkDebugLog.Log("  [Visual] update mode: skip ApplyColorFilters (preserve existing filter settings)");
            }

            // ソースビューのフィルタ設定を分析ビューに追加する。
            // ソースフィルタが visible=false で型枠 DS にマッチする場合は、派生フィルタに
            // 置き換えて型枠 DS を確実に表示する（CopyFilterSettings 内で処理）。
            if (shouldCopySourceFilters && sourceView != null)
            {
                try { CopyFilterSettings(doc, sourceView, view, vr.CreatedShapeIds); }
                catch { }
            }

            // View Filter 適用後も FormworkCategory が確実に表示状態であることを確認
            try
            {
                var fwId2 = new ElementId(FormworkParameterManager.FormworkCategory);
                if (view.CanCategoryBeHidden(fwId2))
                    view.SetCategoryHidden(fwId2, false);
            }
            catch { }

            // 解析ビューの最終状態をログ出力 (表示問題の診断用)
            DiagnoseAnalysisViewState(doc, view);

            return vr;
        }

        /// <summary>
        /// 解析ビューの最終的な可視性関連状態をログ出力する。
        /// 「DSは作成されたが表示されない」問題が発生したときの原因切り分け用。
        /// </summary>
        private static void DiagnoseAnalysisViewState(Document doc, View3D view)
        {
            if (view == null) return;
            try
            {
                FormworkDebugLog.Log($"  [Visual:ViewState] === '{view.Name}' final state ===");

                // SectionBox
                try
                {
                    bool sbActive = view.IsSectionBoxActive;
                    var sb = view.GetSectionBox();
                    string sbInfo = "null";
                    if (sb != null)
                    {
                        var tf = sb.Transform;
                        bool tfIdentity = tf == null || tf.IsIdentity;
                        sbInfo = $"min=({sb.Min.X:F2},{sb.Min.Y:F2},{sb.Min.Z:F2}) " +
                                 $"max=({sb.Max.X:F2},{sb.Max.Y:F2},{sb.Max.Z:F2}) " +
                                 $"tfIdentity={tfIdentity}";
                        if (!tfIdentity && tf != null)
                            sbInfo += $" tfOrigin=({tf.Origin.X:F2},{tf.Origin.Y:F2},{tf.Origin.Z:F2})";
                    }
                    FormworkDebugLog.Log($"    SectionBox: active={sbActive} {sbInfo}");
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"    SectionBox EX: {ex.Message}");
                }

                // CropBox
                try
                {
                    bool cropActive = view.CropBoxActive;
                    bool cropVisible = view.CropBoxVisible;
                    var cb = view.CropBox;
                    string cbInfo = "null";
                    if (cb != null)
                    {
                        cbInfo = $"min=({cb.Min.X:F2},{cb.Min.Y:F2},{cb.Min.Z:F2}) " +
                                 $"max=({cb.Max.X:F2},{cb.Max.Y:F2},{cb.Max.Z:F2})";
                    }
                    FormworkDebugLog.Log(
                        $"    CropBox: active={cropActive} visible={cropVisible} {cbInfo}");
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"    CropBox EX: {ex.Message}");
                }

                // ViewTemplate
                try
                {
                    var tplId = view.ViewTemplateId;
                    string tplName = "(none)";
                    if (tplId != null && tplId != ElementId.InvalidElementId)
                    {
                        var tpl = doc.GetElement(tplId) as View;
                        tplName = tpl?.Name ?? "(unknown)";
                    }
                    FormworkDebugLog.Log($"    ViewTemplate: {tplName}");
                }
                catch { }

                // FormworkCategory state (型枠 DS が属するカテゴリ)
                try
                {
                    var fwId = new ElementId(FormworkParameterManager.FormworkCategory);
                    bool fwHidden = view.GetCategoryHidden(fwId);
                    var ogs = view.GetCategoryOverrides(fwId);
                    string ogsInfo = "null";
                    if (ogs != null)
                    {
                        int trans = ogs.Transparency;
                        bool fgVisible = ogs.IsSurfaceForegroundPatternVisible;
                        bool bgVisible = ogs.IsSurfaceBackgroundPatternVisible;
                        int lw = ogs.ProjectionLineWeight;
                        var halftone = ogs.Halftone;
                        ogsInfo = $"trans={trans} fgVis={fgVisible} bgVis={bgVisible} " +
                                  $"lineWeight={lw} halftone={halftone}";
                    }
                    FormworkDebugLog.Log(
                        $"    FormworkCategory: hidden={fwHidden} ogs={{ {ogsInfo} }}");
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"    FormworkCategory state EX: {ex.Message}");
                }

                // Workset 28Tools_型枠
                try
                {
                    if (doc.IsWorkshared)
                    {
                        var allWs = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset);
                        var fwWs = allWs.FirstOrDefault(ws => ws.Name == "28Tools_型枠");
                        if (fwWs != null)
                        {
                            var wsVis = view.GetWorksetVisibility(fwWs.Id);
                            FormworkDebugLog.Log(
                                $"    Workset '28Tools_型枠': vis={wsVis}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"    Workset state EX: {ex.Message}");
                }

                // Filters
                try
                {
                    var filterIds = view.GetFilters();
                    FormworkDebugLog.Log($"    Filters: total={filterIds?.Count ?? 0}");
                    if (filterIds != null)
                    {
                        int idx = 0;
                        foreach (var fid in filterIds)
                        {
                            string fName = "(?)";
                            bool fVisible = true;
                            try
                            {
                                var fEl = doc.GetElement(fid);
                                fName = fEl?.Name ?? "(null)";
                                fVisible = view.GetFilterVisibility(fid);
                            }
                            catch { }
                            FormworkDebugLog.Log(
                                $"      [{idx}] '{fName}' visible={fVisible}");
                            idx++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"    Filters EX: {ex.Message}");
                }

                FormworkDebugLog.Log($"  [Visual:ViewState] === end ===");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual:ViewState] EX: {ex.Message}");
            }
        }

        /// <summary>
        /// 作成された formwork DirectShape を診断ログに出力する:
        /// - 区分パラメータ毎の件数
        /// - 区分毎に 1 個サンプリングして BBox / Solid 数 / 体積をログ
        /// </summary>
        private static void DiagnoseFormworkShapes(
            Document doc, ICollection<ElementId> shapeIds, FormworkResult result)
        {
            try
            {
                // 最初の3個のDirectShapeのワークセットIDを確認（SetActiveWorksetId有効性診断）
                int wsCheckCount = 0;
                foreach (var id in shapeIds)
                {
                    if (wsCheckCount >= 3) break;
                    var elem = doc.GetElement(id);
                    if (elem == null) continue;
                    try
                    {
                        var wsParam = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                        int wsId = wsParam != null ? wsParam.AsInteger() : -1;
                        bool isRO = wsParam?.IsReadOnly ?? true;
                        FormworkDebugLog.Log(
                            $"  [Visual:WsDiag] DS id={id.IntValue()} wsId={wsId} paramIsReadOnly={isRO}");
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"  [Visual:WsDiag] DS id={id.IntValue()} EX={ex.Message}");
                    }
                    wsCheckCount++;
                }

                var byKey = new Dictionary<string, int>();
                var sampledKeys = new HashSet<string>();
                foreach (var id in shapeIds)
                {
                    var elem = doc.GetElement(id);
                    if (elem == null) continue;
                    string key = "";
                    try
                    {
                        var p = elem.LookupParameter(FormworkParameterManager.ParamGroupKey);
                        if (p != null && p.StorageType == StorageType.String)
                            key = p.AsString() ?? "";
                    }
                    catch { }
                    if (!byKey.ContainsKey(key)) byKey[key] = 0;
                    byKey[key]++;

                    // 区分毎に 1 個サンプリング
                    if (!sampledKeys.Contains(key))
                    {
                        sampledKeys.Add(key);
                        BoundingBoxXYZ bb = null;
                        try { bb = elem.get_BoundingBox(null); } catch { }

                        int solidCount = 0;
                        double totalVol = 0;
                        try
                        {
                            var opts = new Options
                            {
                                ComputeReferences = false,
                                IncludeNonVisibleObjects = false,
                                DetailLevel = ViewDetailLevel.Fine,
                            };
                            var geom = elem.get_Geometry(opts);
                            if (geom != null)
                            {
                                foreach (GeometryObject obj in geom)
                                {
                                    if (obj is Solid s)
                                    {
                                        solidCount++;
                                        totalVol += s.Volume;
                                    }
                                }
                            }
                        }
                        catch { }

                        string bbStr = bb != null
                            ? $"BB=[{bb.Min.X:F2},{bb.Min.Y:F2},{bb.Min.Z:F2}]-[{bb.Max.X:F2},{bb.Max.Y:F2},{bb.Max.Z:F2}]"
                            : "BB=null";
                        FormworkDebugLog.Log(
                            $"  [Visual:Diag] sample key='{key}' id={id.IntValue()} {bbStr} " +
                            $"solids={solidCount} vol={totalVol:F6}");
                    }
                }
                FormworkDebugLog.Log("  [Visual:Diag] DirectShapes by 区分:");
                foreach (var kv in byKey.OrderBy(k => k.Key))
                    FormworkDebugLog.Log($"    '{kv.Key}': {kv.Value}");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual:Diag] EX: {ex.GetType().Name}: {ex.Message}");
            }
        }

        /// <summary>
        /// 型枠不要として除外された要素 (鉄骨・デッキスラブ) の DirectShape を作成する。
        /// 元要素の Solid をそのまま流用してオレンジ色で表示する。
        /// 集計表からは別マーカー値で除外される (ScheduleCreator のフィルタが MarkerValue のみ通す)。
        /// View Filter は既定で非表示なので、ユーザーが手動で ON にしないと見えない。
        /// </summary>
        private static void CreateExcludedShapes(
            Document doc,
            FormworkResult result,
            VisualizerResult vr,
            string sourceViewName)
        {
            if (result.ExcludedResults == null || result.ExcludedResults.Count == 0)
                return;

            int created = 0;
            string hostDisplayName = GetDocumentDisplayName(doc);
            var registry = result.SourceRegistry as Engine.ElementSourceRegistry;
            foreach (var ex in result.ExcludedResults)
            {
                Element src = null;
                Transform xform = Transform.Identity;
                var elemSrc = registry?.Get(ex.ElementId);
                if (elemSrc != null)
                {
                    src = elemSrc.Element;
                    xform = elemSrc.Transform;
                }
                if (src == null)
                {
                    try { src = doc.GetElement(new ElementId(ex.ElementId)); } catch { }
                }
                if (src == null) continue;

                var solids = SolidUnionProcessor.GetSolids(src, xform);
                if (solids.Count == 0) continue;

                // Union を試みる（複数 Solid を 1 シェイプにまとめる）。失敗時は個別に出す。
                var unioned = SolidUnionProcessor.Union(solids);
                var emit = unioned != null ? new List<Solid> { unioned } : solids;

                string levelName = string.Empty;
                try { levelName = Engine.ElementCollector.GetElementLevelName(src); } catch { }

                string label = ExcludedKindLabel(ex.Kind);

                bool firstShape = true;
                foreach (var solid in emit)
                {
                    try
                    {
                        var catOst = new ElementId(FormworkParameterManager.FormworkCategory);
                        var ds = DirectShape.CreateElement(doc, catOst);
                        ds.ApplicationId = "Tools28";
                        ds.ApplicationDataId = "FormworkExcluded";
                        ds.SetShape(new GeometryObject[] { solid });

                        try
                        {
                            string srcDisplay = ex.SourceName == ElementSourceRegistry.HostSourceName
                                ? hostDisplayName : ex.SourceName;
                            FormworkParameterManager.SetInstanceValues(
                                ds,
                                FormworkParameterManager.MarkerValueExcluded,
                                label,
                                levelName,
                                FormworkParameterManager.ExcludedGroupKey,
                                0.0,
                                false,
                                sourceName: srcDisplay,
                                sourceViewName: sourceViewName);
                        }
                        catch { }

                        vr.CreatedShapeIds.Add(ds.Id);
                        if (firstShape) created++;
                        firstShape = false;
                    }
                    catch { }
                }
            }

            FormworkDebugLog.Log($"  [ExcludedShapes] excluded element shapes created: {created}");
        }

        private static string ExcludedKindLabel(ExclusionKind kind)
        {
            switch (kind)
            {
                case ExclusionKind.Steel: return FormworkParameterManager.SteelExcludedLabel;
                case ExclusionKind.DeckSlab: return FormworkParameterManager.DeckSlabExcludedLabel;
                case ExclusionKind.WallSweep: return FormworkParameterManager.WallSweepExcludedLabel;
                case ExclusionKind.SteelStair: return FormworkParameterManager.SteelStairExcludedLabel;
                case ExclusionKind.LgsWall: return FormworkParameterManager.LgsWallExcludedLabel;
                default: return FormworkParameterManager.SteelExcludedLabel;
            }
        }

        /// <summary>
        /// 解析用ビュー以外の既存ビューで DirectShape を非表示にする。
        /// </summary>
        internal static void HideInOtherViews(Document doc, ICollection<ElementId> shapeIds, ElementId analysisViewId)
        {
            if (shapeIds == null || shapeIds.Count == 0) return;

            var allViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Id != analysisViewId && !v.IsTemplate)
                .ToList();

            var idsList = shapeIds.ToList();

            int hideSuccess = 0, hideFail = 0;
            string lastErr = null;
            string failedViewSample = null;
            foreach (var v in allViews)
            {
                if (v is ViewSchedule) continue;
                if (v.ViewType == ViewType.Legend) continue;

                try
                {
                    v.HideElements(idsList);
                    hideSuccess++;
                }
                catch (Exception ex)
                {
                    hideFail++;
                    lastErr = ex.Message;
                    if (failedViewSample == null) failedViewSample = v.Name;
                }
            }
            FormworkDebugLog.Log(
                $"  [Hide] HideInOtherViews: shapes={idsList.Count} keepInViewId={analysisViewId.IntValue()} " +
                $"success={hideSuccess} fail={hideFail}" +
                (hideFail > 0 ? $" lastErr='{lastErr}' failedView='{failedViewSample}'" : ""));
        }

        /// <summary>
        /// 全型枠 DirectShape を「対応する分析ビューでだけ表示」状態に揃える。
        ///   - 非分析ビュー (平面・断面・他 3D ビュー等):
        ///       ① ビューフィルタ "28T_型枠_全非表示" を追加して visible=false に設定（主手段）
        ///          ビューにテンプレートが適用されている場合はテンプレートにも同フィルタを適用
        ///       ② フィルタ適用が失敗したビューは HideElements() で要素レベル非表示（フォールバック）
        ///   - 各分析ビュー: 自分のソースビューに紐づくシェイプのみ表示。それ以外
        ///     (他ソースの分析ビュー由来 + タグなし) は HideElements() で非表示
        /// これにより、過去実行の残りや別ソースの DirectShape が紛れ込むのを防ぐ。
        /// </summary>
        internal static void HideAllFormworkShapesInOtherViews(Document doc)
        {
            // ─── Step 1: 分析ビュー一覧 ──────────────────────────────────────────
            var sourceByAnalysisId = new Dictionary<int, string>();
            var analysisViewIds = new HashSet<int>();
            var analysisViews = Engine.FormworkOutputFinder.FindAllAnalysisViews(doc);
            foreach (var av in analysisViews)
            {
                analysisViewIds.Add(av.Id.IntValue());
                string srcViewName = FormworkParameterManager.GetRelatedSourceView(av);
                if (string.IsNullOrEmpty(srcViewName))
                {
                    if (av.Name.StartsWith(AnalysisViewPrefix))
                        srcViewName = av.Name.Substring(AnalysisViewPrefix.Length);
                    else if (av.Name.StartsWith(LegacyAnalysisViewPrefix))
                        srcViewName = av.Name.Substring(LegacyAnalysisViewPrefix.Length);
                    else if (av.Name.StartsWith(Legacy2AnalysisViewPrefix))
                        srcViewName = av.Name.Substring(Legacy2AnalysisViewPrefix.Length);
                    else if (av.Name.StartsWith(Legacy3AnalysisViewPrefix))
                        srcViewName = av.Name.Substring(Legacy3AnalysisViewPrefix.Length);
                    else if (av.Name == AnalysisViewName || av.Name == LegacyAnalysisViewName
                        || av.Name == Legacy2AnalysisViewName)
                        srcViewName = string.Empty;
                    else
                        srcViewName = string.Empty;
                }
                sourceByAnalysisId[av.Id.IntValue()] = srcViewName ?? string.Empty;
            }
            FormworkDebugLog.Log($"  [Hide] analysisViews found: {analysisViewIds.Count}");

            // Legacy3 ("**型枠：") 系の古い分析ビューは IsAnalysisViewName から外しているが
            // (シート貼付対象から除外するため)、それらに対する非表示フィルタ適用は避けたい
            // (古いビューでも型枠表示を保護する)。ここでスキップ対象集合に追加する。
            int legacy3Count = 0;
            var allViews3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate && v.Name.StartsWith(Legacy3AnalysisViewPrefix));
            foreach (var lv in allViews3D)
            {
                if (analysisViewIds.Add(lv.Id.IntValue())) legacy3Count++;
            }
            FormworkDebugLog.Log($"  [Hide] legacy3 analysisViews added to skip: {legacy3Count}");

            // ─── Step 2: 全型枠 DirectShape を収集 ───────────────────────────────
            // 新カテゴリ (FormworkCategory) と旧カテゴリ (LegacyFormworkCategory) の両方を走査。
            var allShapeIds = new List<ElementId>();
            var shapesByView = new Dictionary<string, List<ElementId>>();
            var untaggedShapes = new List<ElementId>();
            ElementId formworkMarkerParamId = null;

            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(DirectShape))
                .WherePasses(new LogicalOrFilter(
                    new ElementCategoryFilter(FormworkParameterManager.FormworkCategory),
                    new ElementCategoryFilter(FormworkParameterManager.LegacyFormworkCategory)));
            foreach (Element e in collector)
            {
                try
                {
                    var pm = e.LookupParameter(FormworkParameterManager.ParamMarker);
                    if (pm == null || pm.StorageType != StorageType.String) continue;
                    string mv = pm.AsString();
                    if (string.IsNullOrEmpty(mv) ||
                        !mv.StartsWith(FormworkParameterManager.MarkerValue)) continue;

                    allShapeIds.Add(e.Id);
                    if (formworkMarkerParamId == null) formworkMarkerParamId = pm.Id;

                    var pv = e.LookupParameter(FormworkParameterManager.ParamSourceView);
                    string sv = pv?.AsString() ?? string.Empty;
                    if (string.IsNullOrEmpty(sv))
                    {
                        untaggedShapes.Add(e.Id);
                    }
                    else
                    {
                        if (!shapesByView.TryGetValue(sv, out var list))
                        {
                            list = new List<ElementId>();
                            shapesByView[sv] = list;
                        }
                        list.Add(e.Id);
                    }
                }
                catch { }
            }
            FormworkDebugLog.Log(
                $"  [Hide] total formwork shapes={allShapeIds.Count} " +
                $"untagged={untaggedShapes.Count} taggedGroups={shapesByView.Count} " +
                $"markerParamId={formworkMarkerParamId?.IntValue()}");

            if (allShapeIds.Count == 0) return;

            // ─── ワークセット方式（ワークシェアプロジェクト推奨） ───────────────
            // 型枠 DS は全て "28Tools_型枠" ワークセットに属する。
            // WorksetDefaultVisibilitySettings.SetWorksetVisibility(wsId, false) で
            // グローバルデフォルトを Hidden に設定すれば、UseGlobalSetting のビューは
            // 全て自動的に型枠ワークセットを非表示にする（新規作成ビューも自動対応）。
            // 各ビュー個別に明示的 Hidden/Visible を設定して、テンプレートやレガシー
            // ビューの可視性設定を確実に上書きする。
            if (doc.IsWorkshared)
            {
                if (HideFormworkViaWorkset(doc, analysisViewIds, sourceByAnalysisId,
                    shapesByView, untaggedShapes))
                {
                    return; // 完了
                }
                // 失敗時はフィルタ方式へフォールバック
            }

            // ─── Step 3: ビューフィルタ「全非表示」を作成/取得（非ワークシェア用） ──
            // 非分析ビューに適用する ParameterFilterElement を作成しておく。
            ElementId hideFilterId = null;
            if (formworkMarkerParamId != null)
            {
                try
                {
                    hideFilterId = GetOrCreateHideAllFormworkFilter(doc, formworkMarkerParamId);
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Hide] hideFilter create EX: {ex.Message}");
                }
            }
            FormworkDebugLog.Log($"  [Hide] hideFilterId={hideFilterId?.IntValue()}");

            // ─── Step 4: 全ビューを走査 ──────────────────────────────────────────
            // 【非表示化の方針】
            //   テンプレートありビュー → ビューテンプレートのみにフィルタを適用
            //                            (個別ビューには適用しない → V/G ダイアログが汚れない)
            //   テンプレートなしビュー → HideElements（フィルタ不使用 → V/G に影響なし）
            //   分析ビュー           → 他ソースのシェイプのみ HideElements
            var excludedViewTypes = new HashSet<ViewType>
            {
                ViewType.Legend,
                ViewType.Schedule,
                ViewType.Internal,
                ViewType.ProjectBrowser,
                ViewType.SystemBrowser,
            };
            var allViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && !(v is ViewSchedule)
                    && !excludedViewTypes.Contains(v.ViewType))
                .ToList();

            // テンプレートへの適用済みIDを追跡（同一テンプレートに重複適用しない）
            var appliedTemplateIds = new HashSet<int>();
            int filterHideCount = 0, elemHideCount = 0;

            foreach (var v in allViews)
            {
                bool isAnalysis = analysisViewIds.Contains(v.Id.IntValue());

                if (isAnalysis)
                {
                    // 分析ビュー: 自ソース以外のシェイプを要素レベルで非表示
                    sourceByAnalysisId.TryGetValue(v.Id.IntValue(), out string ownSrc);
                    var toHide = new List<ElementId>();
                    foreach (var kv in shapesByView)
                    {
                        if (kv.Key != ownSrc) toHide.AddRange(kv.Value);
                    }
                    toHide.AddRange(untaggedShapes);
                    if (toHide.Count > 0)
                        HideElementsSafe(v, toHide, ref elemHideCount);
                    continue;
                }

                // 非分析ビュー ────────────────────────────────────────────────────
                bool hasTemplate = v.ViewTemplateId != null
                    && v.ViewTemplateId != ElementId.InvalidElementId;

                if (hasTemplate && hideFilterId != null)
                {
                    // テンプレートありビュー: フィルタをテンプレートのみに適用
                    int tmplId = v.ViewTemplateId.IntValue();
                    if (!appliedTemplateIds.Contains(tmplId))
                    {
                        try
                        {
                            var tmpl = doc.GetElement(v.ViewTemplateId) as View;
                            if (tmpl != null)
                            {
                                SetFilterAtHighestPriority(tmpl, hideFilterId);
                                tmpl.SetFilterVisibility(hideFilterId, false);
                                appliedTemplateIds.Add(tmplId);
                                filterHideCount++;
                                FormworkDebugLog.Log(
                                    $"  [Hide] template='{tmpl.Name}' filterOK");
                            }
                        }
                        catch (Exception exTmpl)
                        {
                            FormworkDebugLog.Log(
                                $"  [Hide] view='{v.Name}' templateFilter EX: {exTmpl.Message}");
                            // テンプレートへの適用失敗 → ビュー直接に HideElements
                            HideElementsSafe(v, allShapeIds, ref elemHideCount);
                        }
                    }
                    // 同テンプレートの2件目以降はスキップ（テンプレートで一括適用済み）
                }
                else
                {
                    // テンプレートなしビュー: HideElements（V/G に影響なし）
                    HideElementsSafe(v, allShapeIds, ref elemHideCount);
                }
            }

            FormworkDebugLog.Log(
                $"  [Hide] done: views={allViews.Count} templateFilter={filterHideCount} " +
                $"elemHide={elemHideCount} appliedTemplates={appliedTemplateIds.Count}");
        }

        /// <summary>
        /// 指定フィルタをビューのフィルタリスト先頭（最高優先度）に配置する。
        /// Revit の filter visibility は「最初にマッチしたフィルタ」が勝つ仕様のため、
        /// 既存の高優先度フィルタに型枠 DS がマッチして visible=true を返してしまうと
        /// 末尾の非表示フィルタが負ける。これを防ぐため先頭に移動する。
        ///
        /// View.SetOrderOnFilters は Revit 2022+ にしかないため、全バージョン対応のため
        /// 既存「他」フィルタを一旦 Remove してから再 Add する方式で並び替える。
        /// 「他」フィルタの上書設定 (overrides / visibility) は保存してから復元する。
        /// </summary>
        private static void SetFilterAtHighestPriority(View v, ElementId filterId)
        {
            var current = v.GetFilters().ToList();
            // 既に先頭にある場合は何もしない
            if (current.Count > 0 && current[0] == filterId) return;

            // 非表示フィルタが未追加なら追加
            if (!current.Contains(filterId))
            {
                v.AddFilter(filterId);
                current = v.GetFilters().ToList();
            }

            // 「他」フィルタを順序保ったまま取得
            var others = current.Where(f => f != filterId).ToList();

            // 他フィルタの設定を保存
            var savedOverrides = new Dictionary<ElementId, OverrideGraphicSettings>();
            var savedVisibility = new Dictionary<ElementId, bool>();
            foreach (var f in others)
            {
                try { savedOverrides[f] = v.GetFilterOverrides(f); } catch { }
                try { savedVisibility[f] = v.GetFilterVisibility(f); }
                catch { savedVisibility[f] = true; }
            }

            // 他フィルタを一旦すべて削除（filterId だけが残る → 先頭）
            foreach (var f in others)
            {
                try { v.RemoveFilter(f); } catch { }
            }

            // 他フィルタを元の順序で再追加（filterId の後ろに来る）
            foreach (var f in others)
            {
                try
                {
                    v.AddFilter(f);
                    if (savedOverrides.TryGetValue(f, out var og) && og != null)
                    {
                        try { v.SetFilterOverrides(f, og); } catch { }
                    }
                    if (savedVisibility.TryGetValue(f, out var vis))
                    {
                        try { v.SetFilterVisibility(f, vis); } catch { }
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// 型枠 DirectShape を全て非表示にするビューフィルタを作成または取得する。
        /// ルール: 28Tools_FormworkMarker パラメータが "28Tools_Formwork" で始まる要素
        /// カテゴリ: FormworkCategory (新) + LegacyFormworkCategory (旧 DS 残骸対応)
        /// </summary>
        private static ElementId GetOrCreateHideAllFormworkFilter(
            Document doc, ElementId markerParamId)
        {
            const string filterName = "28T_型枠_全非表示";

            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>()
                .FirstOrDefault(f => f.Name == filterName);
            if (existing != null)
            {
                // 旧バージョン作成のフィルタが OST_GenericModel のみを対象としている場合は
                // 作り直す (新カテゴリを含めるため)
                var eCats = existing.GetCategories();
                int newCatInt = (int)FormworkParameterManager.FormworkCategory;
                bool hasNew = false;
                if (eCats != null)
                    foreach (var c in eCats) if (c.IntValue() == newCatInt) { hasNew = true; break; }
                if (hasNew)
                {
                    FormworkDebugLog.Log(
                        $"  [Hide] hideFilter '{filterName}' reused id={existing.Id.IntValue()}");
                    return existing.Id;
                }
                try { doc.Delete(existing.Id); } catch { }
                FormworkDebugLog.Log(
                    $"  [Hide] hideFilter '{filterName}' deleted (legacy GenericModel-only) → recreating");
            }

            var cats = new List<ElementId>
            {
                new ElementId(FormworkParameterManager.FormworkCategory),
                new ElementId(FormworkParameterManager.LegacyFormworkCategory),
            };
#if REVIT2026
            FilterRule rule = ParameterFilterRuleFactory.CreateBeginsWithRule(
                markerParamId, FormworkParameterManager.MarkerValue);
#else
            FilterRule rule = ParameterFilterRuleFactory.CreateBeginsWithRule(
                markerParamId, FormworkParameterManager.MarkerValue, false);
#endif
            var ef = new ElementParameterFilter(rule);
            var newFilter = ParameterFilterElement.Create(doc, filterName, cats, ef);
            FormworkDebugLog.Log(
                $"  [Hide] hideFilter '{filterName}' created id={newFilter.Id.IntValue()}");
            return newFilter.Id;
        }

        /// <summary>
        /// 型枠ワークセット "28Tools_型枠" のグローバルデフォルト可視性を Hidden に設定し、
        /// 全ビューおよびビューテンプレートのワークセット可視性を明示的に設定する。
        ///   - 非分析ビュー / 非分析テンプレート → Hidden
        ///   - 分析ビュー / 分析ビュー用テンプレート → Visible
        /// 分析ビュー内では、HideElements で「他ソース由来 + タグなし」シェイプを非表示。
        /// 成功時 true。ワークセット未取得や例外時は false（呼出し側でフォールバック）。
        /// </summary>
        private static bool HideFormworkViaWorkset(
            Document doc,
            HashSet<int> analysisViewIds,
            Dictionary<int, string> sourceByAnalysisId,
            Dictionary<string, List<ElementId>> shapesByView,
            List<ElementId> untaggedShapes)
        {
            try
            {
                var wsId = GetOrCreateFormworkWorkset(doc);
                if (wsId == null)
                {
                    FormworkDebugLog.Log("  [Hide][WS] formwork workset not available");
                    return false;
                }

                // (1) グローバルデフォルト可視性を Hidden に
                // GetOrCreateFormworkWorkset でも設定しているが、ここでも確実に適用する。
                // (初回作成直後の TX コミット前でも機能するよう二重保険)
                TrySetWorksetGlobalVisibility(doc, wsId, "28Tools_型枠(HidePhase)", isNew: false);

                // (2) 分析ビューが使うテンプレート ID を収集
                //     （それらテンプレートには Visible を設定して分析ビューの可視性を維持）
                var analysisTemplateIds = new HashSet<int>();
                foreach (var avId in analysisViewIds)
                {
                    try
                    {
                        var av = doc.GetElement(new ElementId(avId)) as View;
                        if (av?.ViewTemplateId != null &&
                            av.ViewTemplateId != ElementId.InvalidElementId)
                        {
                            analysisTemplateIds.Add(av.ViewTemplateId.IntValue());
                        }
                    }
                    catch { }
                }

                // (3) 全ビュー（テンプレート含む）のワークセット可視性を明示設定
                var excludedViewTypes = new HashSet<ViewType>
                {
                    ViewType.Legend,
                    ViewType.Schedule,
                    ViewType.Internal,
                    ViewType.ProjectBrowser,
                    ViewType.SystemBrowser,
                };
                var allViewsAndTemplates = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !(v is ViewSchedule)
                        && !excludedViewTypes.Contains(v.ViewType))
                    .ToList();

                int hiddenCount = 0, visibleCount = 0, skipped = 0;
                foreach (var v in allViewsAndTemplates)
                {
                    bool isAnalysisView = analysisViewIds.Contains(v.Id.IntValue());
                    bool isAnalysisTemplate = v.IsTemplate
                        && analysisTemplateIds.Contains(v.Id.IntValue());

                    WorksetVisibility target = (isAnalysisView || isAnalysisTemplate)
                        ? WorksetVisibility.Visible
                        : WorksetVisibility.Hidden;

                    try
                    {
                        var cur = v.GetWorksetVisibility(wsId);
                        if (cur != target)
                        {
                            v.SetWorksetVisibility(wsId, target);
                        }
                        if (target == WorksetVisibility.Hidden) hiddenCount++;
                        else visibleCount++;
                    }
                    catch
                    {
                        skipped++;
                    }
                }

                // (4) 分析ビュー内で「他ソース由来 + タグなし」シェイプを HideElements
                foreach (var avIdInt in analysisViewIds)
                {
                    try
                    {
                        var av = doc.GetElement(new ElementId(avIdInt)) as View;
                        if (av == null) continue;
                        sourceByAnalysisId.TryGetValue(avIdInt, out string ownSrc);
                        var toHide = new List<ElementId>();
                        foreach (var kv in shapesByView)
                        {
                            if (kv.Key != ownSrc) toHide.AddRange(kv.Value);
                        }
                        toHide.AddRange(untaggedShapes);
                        if (toHide.Count == 0) continue;
                        int dummy = 0;
                        HideElementsSafe(av, toHide, ref dummy);
                    }
                    catch { }
                }

                FormworkDebugLog.Log(
                    $"  [Hide][WS] done: hiddenViews={hiddenCount} visibleViews={visibleCount} " +
                    $"skipped={skipped} analysisTemplates={analysisTemplateIds.Count}");
                return true;
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Hide][WS] EX (fallback to filter): {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// HideElements をバッチで実行し、失敗時に1件ずつ再試行する。
        /// </summary>
        private static void HideElementsSafe(
            View v, List<ElementId> toHide, ref int counter)
        {
            if (toHide == null || toHide.Count == 0) return;
            try
            {
                v.HideElements(toHide);
                counter += toHide.Count;
                FormworkDebugLog.Log(
                    $"  [Hide] view='{v.Name}'({v.ViewType}) elemHide OK count={toHide.Count}");
            }
            catch (Exception ex)
            {
                int ok = 0;
                foreach (var id in toHide)
                {
                    try
                    {
                        v.HideElements(new List<ElementId> { id });
                        ok++;
                    }
                    catch { }
                }
                counter += ok;
                FormworkDebugLog.Log(
                    $"  [Hide] view='{v.Name}'({v.ViewType}) " +
                    $"batchEX='{ex.Message}' perItem={ok}/{toHide.Count}");
            }
        }

        private static DirectShapeType GetOrCreateDirectShapeType(Document doc, string typeName)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(DirectShapeType))
                .Cast<DirectShapeType>()
                .FirstOrDefault(t => t.Name == typeName);
            if (existing != null) return existing;

            try
            {
                var catId = new ElementId(FormworkParameterManager.FormworkCategory);
                return DirectShapeType.Create(doc, typeName, catId);
            }
            catch
            {
                return null;
            }
        }

        private static string BuildTypeName(FormworkSettings s, string key, ElementResult er)
        {
            switch (s.ColorScheme)
            {
                case ColorSchemeType.ByZone:
                    return $"型枠_工区_{(string.IsNullOrEmpty(er.Zone) ? "未設定" : er.Zone)}";
                case ColorSchemeType.ByFormworkType:
                    return $"型枠_種別_{(string.IsNullOrEmpty(er.FormworkType) ? "未設定" : er.FormworkType)}";
                default:
                    return $"型枠_{CategoryLabel(er.Category)}";
            }
        }

        private static string CategoryLabel(CategoryGroup cg)
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
                default: return "その他";
            }
        }

        private static bool IsDeducted(FaceType t)
        {
            return t == FaceType.DeductedTop ||
                   t == FaceType.DeductedBottom ||
                   t == FaceType.DeductedContact;
        }

        private static string GetKey(ElementResult er, FormworkSettings s)
        {
            switch (s.ColorScheme)
            {
                case ColorSchemeType.ByZone: return string.IsNullOrEmpty(er.Zone) ? "未設定" : er.Zone;
                case ColorSchemeType.ByFormworkType: return string.IsNullOrEmpty(er.FormworkType) ? "未設定" : er.FormworkType;
                default: return CategoryLabel(er.Category);
            }
        }

        private static Dictionary<string, (byte R, byte G, byte B)> AssignColors(
            FormworkResult result, FormworkSettings s)
        {
            var map = new Dictionary<string, (byte, byte, byte)>();
            if (s.ColorScheme == ColorSchemeType.ByCategory)
            {
                foreach (var kv in _categoryColors)
                    map[CategoryLabel(kv.Key)] = kv.Value;
                return map;
            }

            var keys = new HashSet<string>();
            foreach (var er in result.ElementResults) keys.Add(GetKey(er, s));

            int i = 0;
            foreach (var k in keys.OrderBy(x => x))
            {
                map[k] = _autoPalette[i % _autoPalette.Length];
                i++;
            }
            return map;
        }

        /// <summary>
        /// ソースビューが 3D ビューの場合、その視点（EyePosition / ForwardDirection / UpDirection）を
        /// 新しい解析ビューにコピーする。3D ビュー以外の場合は何もしない（既定のアイソメトリックを維持）。
        /// </summary>
        /// <summary>
        /// ビューキューブの青い角（前・上・右の交差する角）からのアイソメトリック視点を設定する。
        /// Revit の標準プロジェクト方位では:
        ///   - 右 (右側面) = +X、左 (左側面) = -X
        ///   - 前 (正面)   = -Y、後 (背面)   = +Y
        ///   - 上 (上面)   = +Z、下 (下面)   = -Z
        /// 「前・上・右」の交差する角はカメラ位置 (+X, -Y, +Z) に対応する。
        /// ForwardDirection: そこから原点を見る方向 = (-1, +1, -1)。
        /// UpDirection: 前方と直交し概ね上向き (-1, +1, 2) normalized。
        /// </summary>
        private static void SetIsometricOrientation(View3D view)
        {
            if (view == null) return;
            try
            {
                var forward = new XYZ(-1, 1, -1);
                var up = new XYZ(-1, 1, 2);
                var eye = new XYZ(100, -100, 100);
                var orient = new ViewOrientation3D(eye, up, forward);
                view.SetOrientation(orient);
                FormworkDebugLog.Log("  [Visual] isometric orientation set (front-top-right corner)");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] SetIsometricOrientation EX: {ex.Message}");
            }
        }

        private static void CopyOrientationIfPossible(View sourceView, View3D targetView)
        {
            if (!(sourceView is View3D src) || targetView == null) return;
            try
            {
                var orient = src.GetOrientation();
                if (orient != null)
                    targetView.SetOrientation(orient);
            }
            catch { }
        }

        /// <summary>
        /// CurrentView スコープでソースが 3D ビューのとき、ソースの切断ボックス位置と
        /// 有効状態をコピーする。切断ボックスが正常にコピーされた場合は true を返す。
        /// ソースの切断ボックスが非アクティブな場合は false を返し、
        /// false を返した場合、切断ボックスなし状態のままにする（ソース非アクティブを尊重）。
        /// </summary>
        private static bool CopySectionBoxFromSource(View3D source, View3D target)
        {
            try
            {
                bool srcActive = source.IsSectionBoxActive;
                FormworkDebugLog.Log(
                    $"  [Visual] CopySectionBoxFromSource: source='{source?.Name}' IsSectionBoxActive={srcActive}");
                if (!srcActive)
                {
                    FormworkDebugLog.Log("  [Visual] CopySectionBoxFromSource: ソース非アクティブ → 解析ビューも切断ボックスなし");
                    return false;
                }
                var sb = source.GetSectionBox();
                if (sb != null)
                {
                    // 型枠 DirectShape (面から 15mm 押し出し) が境界でクリップされないよう
                    // ソースの切断ボックスに小さなマージン (約 30mm) を追加する。
                    double margin = 0.1; // ft
                    var copy = new BoundingBoxXYZ
                    {
                        Min = new XYZ(sb.Min.X - margin, sb.Min.Y - margin, sb.Min.Z - margin),
                        Max = new XYZ(sb.Max.X + margin, sb.Max.Y + margin, sb.Max.Z + margin),
                        Transform = sb.Transform,
                    };
                    target.SetSectionBox(copy);
                    FormworkDebugLog.Log(
                        $"  [Visual] CopySectionBoxFromSource: copied sb min=({sb.Min.X:F2},{sb.Min.Y:F2},{sb.Min.Z:F2}) " +
                        $"max=({sb.Max.X:F2},{sb.Max.Y:F2},{sb.Max.Z:F2})");
                }
                target.IsSectionBoxActive = true;
                return true;
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] CopySectionBoxFromSource EX: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// ソースビューの V/G 上書き設定 (モデル/注釈カテゴリの表示/非表示・色) を
        /// ターゲット 3D ビューにコピーする。ターゲットで非表示にできないカテゴリは
        /// スキップ (例: 平面ビュー専用のカテゴリ)。
        ///
        /// 解析ビューに必要なカテゴリ (FormworkCategory = 型枠 DirectShape) は強制的に表示にする。
        /// 切断ボックスのアウトライン・レベル線は視覚的ノイズなので強制非表示。
        /// </summary>
        private static void CopyCategoryVisibilityAndOverrides(
            Document doc, View source, View3D target)
        {
            int catTotal = 0, catCopied = 0, catSkipped = 0, catErrors = 0;
            int hiddenCopied = 0;
            FormworkDebugLog.Log($"  [Visual] CopyCategoryVisibility 開始: source='{source?.Name}' target='{target?.Name}'");
            // 型枠 DirectShape カテゴリ (新 + 旧)。
            // ソースが当該カテゴリに透過 100% / SurfaceForegroundPatternVisible=false /
            // ProjectionLineWeight=0 等を設定していると、コピー後の解析ビューで
            // 型枠 DS が「存在するが見えない」状態になる。
            // カテゴリ可視性・V/G ともコピー対象から除外する（後段で明示クリーン化する）。
            var fwCatIdInt = (int)FormworkParameterManager.FormworkCategory;
            var legacyFwCatIdInt = (int)FormworkParameterManager.LegacyFormworkCategory;

            try
            {
                var allCategories = doc.Settings.Categories;
                foreach (Category cat in allCategories)
                {
                    if (cat == null) continue;
                    catTotal++;
                    var catId = cat.Id;
                    if (!target.CanCategoryBeHidden(catId)) { catSkipped++; continue; }
                    bool isFormworkCategory =
                        catId.IntValue() == fwCatIdInt || catId.IntValue() == legacyFwCatIdInt;
                    try
                    {
                        // 表示/非表示
                        bool srcHidden = source.GetCategoryHidden(catId);
                        target.SetCategoryHidden(catId, srcHidden);
                        if (srcHidden) hiddenCopied++;
                        catCopied++;

                        // V/G 上書き (色・線種・透過等)
                        // source.IsCategoryOverridable のチェックは外す:
                        // View Template ロック時に false を返す場合があるが、
                        // GetCategoryOverrides では実効値を取得できる。
                        // 型枠カテゴリ (新/旧) はコピーせず、後段で空 OGS を当てる。
                        if (!isFormworkCategory && target.IsCategoryOverridable(catId))
                        {
                            try
                            {
                                var ogs = source.GetCategoryOverrides(catId);
                                if (ogs != null)
                                    target.SetCategoryOverrides(catId, ogs);
                            }
                            catch { }
                        }
                    }
                    catch { catErrors++; /* per-category failures are non-fatal */ }

                    // サブカテゴリも同様にコピー (OST_GenericModel のサブは型枠 DS に
                    // 影響しないため通常コピーで OK)
                    foreach (Category sub in cat.SubCategories)
                    {
                        if (sub == null) continue;
                        var subId = sub.Id;
                        if (!target.CanCategoryBeHidden(subId)) continue;
                        try
                        {
                            bool srcHidden = source.GetCategoryHidden(subId);
                            target.SetCategoryHidden(subId, srcHidden);
                            if (target.IsCategoryOverridable(subId))
                            {
                                try
                                {
                                    var ogs = source.GetCategoryOverrides(subId);
                                    if (ogs != null)
                                        target.SetCategoryOverrides(subId, ogs);
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] CopyCategoryVisibility EX: {ex.Message}");
            }
            FormworkDebugLog.Log(
                $"  [Visual] CopyCategoryVisibility 完了: total={catTotal} copied={catCopied} " +
                $"hiddenCopied={hiddenCopied} skipped={catSkipped} errors={catErrors}");

            // ※フィルタのコピーは CreateVisualization 内で ApplyColorFilters 後に実施する。
            //   ソースフィルタを型枠フィルタより後に追加することで、型枠フィルタの優先度が
            //   高くなり、ソースフィルタが OST_GenericModel を非表示にしていても
            //   型枠 DirectShape が確実に表示・色分けされる。

            // 型枠 DirectShape (FormworkCategory) は型枠表示の主役なので強制的に表示状態に上書き。
            // V/G オーバーライドも空にリセットする (ソース側に透過 100% やパターン不可視等が
            // 設定されていると、Color フィルタを適用しても型枠 DS が見えなくなるため)。
            // 旧カテゴリ (LegacyFormworkCategory) も同様に処理する (旧 DS がまだ残っている場合の保険)。
            foreach (var bic in new[] { FormworkParameterManager.FormworkCategory,
                                        FormworkParameterManager.LegacyFormworkCategory })
            {
                try
                {
                    var fwId = new ElementId(bic);
                    if (target.CanCategoryBeHidden(fwId))
                        target.SetCategoryHidden(fwId, false);
                    if (target.IsCategoryOverridable(fwId))
                    {
                        target.SetCategoryOverrides(fwId, new OverrideGraphicSettings());
                        FormworkDebugLog.Log($"  [Visual] {bic} V/G overrides reset");
                    }
                }
                catch { }
            }

            // 切断ボックスのアウトライン・レベル線は解析ビューで邪魔なので強制非表示
            HideClutterCategories(target);
        }

        /// <summary>
        /// ソースビューのワークセット可視性設定をターゲットビューにコピーする。
        /// V/G 上書き > ワークセット タブの設定に相当する。
        /// </summary>
        private static void CopyWorksetVisibility(Document doc, View source, View3D target)
        {
            if (!doc.IsWorkshared) return;
            try
            {
                var worksets = new FilteredWorksetCollector(doc)
                    .OfKind(WorksetKind.UserWorkset);
                int copied = 0, errors = 0;
                foreach (var ws in worksets)
                {
                    try
                    {
                        WorksetVisibility vis = source.GetWorksetVisibility(ws.Id);
                        target.SetWorksetVisibility(ws.Id, vis);
                        copied++;
                    }
                    catch { errors++; }
                }
                FormworkDebugLog.Log(
                    $"  [Visual] CopyWorksetVisibility: copied={copied} errors={errors}");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] CopyWorksetVisibility EX: {ex.Message}");
            }
        }

        /// <summary>
        /// 型枠 DirectShape が属するワークセットをビューで強制的に表示状態にする。
        /// CopyWorksetVisibility でソースビューのワークセット非表示設定が引き継がれた場合に
        /// 専用ワークセット「28Tools_型枠」を解析ビューで強制表示する。
        /// DirectShape はアクティブワークセットを事前変更して作成済みのため、
        /// ここではワークセットの可視性設定のみ行う（移動処理は不要）。
        /// </summary>
        private static void EnsureFormworkWorksetsVisible(
            Document doc, View3D view, IList<ElementId> shapeIds)
        {
            if (shapeIds == null || shapeIds.Count == 0) return;
            try
            {
                var wsId = GetOrCreateFormworkWorkset(doc);
                if (wsId == null) return;
                view.SetWorksetVisibility(wsId, WorksetVisibility.Visible);
                // 設定が実際に反映されたか確認
                var actualVis = view.GetWorksetVisibility(wsId);
                FormworkDebugLog.Log(
                    $"  [Visual] EnsureFormworkWorksetsVisible: wsId={wsId.IntegerValue} shapeCount={shapeIds.Count} setVis=Visible actualVis={actualVis}");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] EnsureFormworkWorksetsVisible EX: {ex.Message}");
            }
        }

        /// <summary>
        /// 専用ワークセット「28Tools_型枠」を検索または作成してIDを返す。
        /// </summary>
        internal static WorksetId GetOrCreateFormworkWorkset(Document doc)
        {
            const string wsName = "28Tools_型枠";
            var allWs = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset);
            var existing = allWs.FirstOrDefault(ws => ws.Name == wsName);
            if (existing != null)
            {
                // 既存ワークセットでも「全ビューに表示」を確実にオフにする
                TrySetWorksetGlobalVisibility(doc, existing.Id, wsName, isNew: false);
                return existing.Id;
            }

            var newWs = Workset.Create(doc, wsName);
            FormworkDebugLog.Log($"  [Visual] 専用ワークセット作成: '{wsName}' id={newWs.Id.IntegerValue}");
            // 新規ワークセットはデフォルトで「全ビューに表示」=true になるため即座にオフに設定
            TrySetWorksetGlobalVisibility(doc, newWs.Id, wsName, isNew: true);
            return newWs.Id;
        }

        private static void TrySetWorksetGlobalVisibility(
            Document doc, WorksetId wsId, string wsName, bool isNew)
        {
            try
            {
                var defaults = WorksetDefaultVisibilitySettings
                    .GetWorksetDefaultVisibilitySettings(doc);
                defaults.SetWorksetVisibility(wsId, false);
                FormworkDebugLog.Log(
                    $"  [Visual] WS '{wsName}' 全ビュー非表示に設定 " +
                    $"(isNew={isNew}, wsId={wsId.IntegerValue})");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log(
                    $"  [Visual] WS '{wsName}' 全ビュー非表示 設定失敗 EX: {ex.Message}");
            }
        }

        /// <summary>
        /// ソースビューのフィルタ設定を分析ビューにコピーする。
        ///
        /// 派生フィルタは作らず、既存ソースフィルタをそのまま分析ビューに参照させる。
        /// フィルタ名・内容・ID・表示設定 (visibility / overrides) を完全にソースから引き継ぐ。
        ///
        /// 注意: ソースで visible=false かつ型枠 DS にマッチするフィルタの場合、
        /// 分析ビューでも型枠 DS が非表示になる可能性がある (Revit 仕様: 一致する
        /// どれかのフィルタが visible=false なら要素は隠れる)。
        /// この場合は警告ログを出すだけで動作は変更しない。ユーザーが必要に応じて
        /// 分析ビュー側で当該フィルタを手動で visible=true に切り替える、または
        /// ソースフィルタのカテゴリから OST_GenericModel を外すことで解消できる。
        /// </summary>
        private static void CopyFilterSettings(Document doc, View source, View3D target,
            ICollection<ElementId> formworkShapeIds = null)
        {
            try
            {
                var srcFilterIds = source.GetFilters();
                if (srcFilterIds == null || srcFilterIds.Count == 0) return;

                // 型枠 DS のサンプルを取得 (ソースフィルタが型枠 DS を隠す可能性の警告判定用)
                ElementId sampleFormworkId = null;
                if (formworkShapeIds != null)
                {
                    foreach (var sid in formworkShapeIds)
                    {
                        if (sid == null || sid == ElementId.InvalidElementId) continue;
                        sampleFormworkId = sid;
                        break;
                    }
                }

                int copied = 0, warnHidesFormwork = 0;
                var existingTargetFilters = target.GetFilters();
                foreach (var fid in srcFilterIds)
                {
                    if (fid == null || fid == ElementId.InvalidElementId) continue;
                    try
                    {
                        bool srcVis = source.GetFilterVisibility(fid);

                        // ソースフィルタをそのまま分析ビューに参照させる (新規フィルタは作らない)
                        if (!existingTargetFilters.Contains(fid))
                        {
                            target.AddFilter(fid);
                            existingTargetFilters = target.GetFilters();
                        }
                        var ogs = source.GetFilterOverrides(fid);
                        if (ogs != null) target.SetFilterOverrides(fid, ogs);

                        // 表示設定もソースの値を完全に引き継ぐ。
                        target.SetFilterVisibility(fid, srcVis);

                        // visible=false が型枠 DS を隠してしまう可能性のあるフィルタには警告ログを出す。
                        if (!srcVis && sampleFormworkId != null
                            && FilterMatchesFormwork(doc, fid, sampleFormworkId))
                        {
                            warnHidesFormwork++;
                            string fName = (doc.GetElement(fid) as ParameterFilterElement)?.Name ?? "?";
                            FormworkDebugLog.Log(
                                $"  [Visual] ⚠️ filter fid={fid.IntValue()} '{fName}' visible=false " +
                                $"matches formwork DS → 分析ビューで型枠 DS が隠れる可能性あり " +
                                $"(必要なら当該フィルタを手動で表示=ONに切替)");
                        }
                        copied++;
                    }
                    catch (Exception exFilter)
                    {
                        FormworkDebugLog.Log(
                            $"  [Visual] CopyFilter fid={fid.IntValue()} EX: {exFilter.Message}");
                    }
                }
                FormworkDebugLog.Log(
                    $"  [Visual] CopyFilterSettings: copied={copied}/{srcFilterIds.Count} " +
                    $"(warnHidesFormwork={warnHidesFormwork})");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] CopyFilterSettings EX: {ex.Message}");
            }
        }

        /// <summary>
        /// 指定フィルタが型枠 DS サンプルにマッチするか判定する。
        /// PassesFilter で実際の要素に対して評価する。ルール無し（カテゴリのみ）の
        /// フィルタは全カテゴリ要素に一致するため true 扱い。
        /// </summary>
        private static bool FilterMatchesFormwork(Document doc, ElementId fid, ElementId sampleFormworkId)
        {
            try
            {
                var pfe = doc.GetElement(fid) as ParameterFilterElement;
                if (pfe == null) return true; // 不明な場合は安全側に倒す
                // カテゴリチェック: 型枠 DS のカテゴリ (新/旧) を含むかどうか
                var cats = pfe.GetCategories();
                if (cats == null) return false;
                var newCatInt = (int)FormworkParameterManager.FormworkCategory;
                var legacyCatInt = (int)FormworkParameterManager.LegacyFormworkCategory;
                bool hasGm = false;
                foreach (var cid in cats)
                {
                    if (cid.IntValue() == newCatInt || cid.IntValue() == legacyCatInt)
                    { hasGm = true; break; }
                }
                if (!hasGm) return false;

                var ef = pfe.GetElementFilter();
                if (ef == null) return true; // カテゴリのみ → 当該カテゴリ全要素に一致
                return ef.PassesFilter(doc, sampleFormworkId);
            }
            catch
            {
                return true; // 例外時は安全側
            }
        }

        private static ParameterFilterElement FindFilterByExactName(Document doc, string name)
        {
            try
            {
                return new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .FirstOrDefault(f => f.Name == name);
            }
            catch { return null; }
        }

        private static void DeleteExistingParamFilterByName(Document doc, string name)
        {
            try
            {
                var existing = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .FirstOrDefault(f => f.Name == name);
                if (existing != null)
                {
                    try { doc.Delete(existing.Id); } catch { }
                }
            }
            catch { }
        }

        /// <summary>
        /// 旧実装で生成された派生フィルタを削除する:
        ///   - 旧旧 ("28T_FW_*_GM" / "28T_FW_*_Other"): 数値ID命名の派生
        ///   - 旧   ("*_型枠除外GM" / "*_型枠除外"):   元名派生の派生
        ///
        /// 現行の CopyFilterSettings は派生フィルタを作らずソースフィルタを直接参照させる
        /// ため、これらは全て不要になった。トランザクション内で呼び出すこと。
        /// </summary>
        internal static void CleanupLegacySplitFilters(Document doc)
        {
            try
            {
                var toDelete = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .Where(f =>
                        (f.Name.StartsWith("28T_FW_") &&
                            (f.Name.EndsWith("_GM") || f.Name.EndsWith("_Other"))) ||
                        f.Name.EndsWith("_型枠除外GM") ||
                        f.Name.EndsWith("_型枠除外"))
                    .Select(f => f.Id)
                    .ToList();
                if (toDelete.Count == 0) return;
                foreach (var id in toDelete)
                {
                    try { doc.Delete(id); } catch { }
                }
                FormworkDebugLog.Log($"  [Visual] CleanupLegacySplitFilters: deleted={toDelete.Count}");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] CleanupLegacySplitFilters EX: {ex.Message}");
            }
        }

        /// <summary>
        /// 3D ビューでセクションボックスのアウトラインとレベル線を非表示にする。
        /// 解析ビューでは型枠面の色分けが主役なので、これらが視覚的ノイズになるのを防ぐ。
        /// </summary>
        private static void HideClutterCategories(View3D view)
        {
            var hideCats = new[]
            {
                BuiltInCategory.OST_SectionBox, // 切断ボックスのアウトライン
                BuiltInCategory.OST_Levels,     // レベル線
            };
            foreach (var bic in hideCats)
            {
                try
                {
                    var catId = new ElementId(bic);
                    if (view.CanCategoryBeHidden(catId))
                        view.SetCategoryHidden(catId, true);
                }
                catch { }
            }
        }

        /// <summary>
        /// 解析対象要素の BoundingBox を全て囲む形でセクションボックスを設定し有効化。
        /// </summary>
        private static void EnableSectionBox(Document doc, View3D view, FormworkResult result)
        {
            try
            {
                XYZ minP = null, maxP = null;
                var ids = new List<int>();
                foreach (var er in result.ElementResults) ids.Add(er.ElementId);
                if (result.ExcludedResults != null)
                    foreach (var ex in result.ExcludedResults) ids.Add(ex.ElementId);

                var registry = result.SourceRegistry as Engine.ElementSourceRegistry;
                foreach (var id in ids)
                {
                    Element elem = null;
                    Transform xform = Transform.Identity;
                    var src = registry?.Get(id);
                    if (src != null) { elem = src.Element; xform = src.Transform; }
                    if (elem == null)
                    {
                        try { elem = doc.GetElement(new ElementId(id)); } catch { }
                    }
                    if (elem == null) continue;
                    BoundingBoxXYZ bb = null;
                    try { bb = elem.get_BoundingBox(null); } catch { }
                    if (bb == null) continue;
                    // リンク要素の BoundingBox はリンクローカル座標なのでホスト座標に変換
                    if (xform != null && !xform.IsIdentity)
                        bb = TransformBoundingBox(bb, xform);
                    if (minP == null)
                    {
                        minP = bb.Min;
                        maxP = bb.Max;
                    }
                    else
                    {
                        minP = new XYZ(System.Math.Min(minP.X, bb.Min.X),
                                       System.Math.Min(minP.Y, bb.Min.Y),
                                       System.Math.Min(minP.Z, bb.Min.Z));
                        maxP = new XYZ(System.Math.Max(maxP.X, bb.Max.X),
                                       System.Math.Max(maxP.Y, bb.Max.Y),
                                       System.Math.Max(maxP.Z, bb.Max.Z));
                    }
                }

                if (minP != null && maxP != null)
                {
                    double margin = 1.0; // 1ft 余裕
                    var sb = new BoundingBoxXYZ
                    {
                        Min = new XYZ(minP.X - margin, minP.Y - margin, minP.Z - margin),
                        Max = new XYZ(maxP.X + margin, maxP.Y + margin, maxP.Z + margin),
                    };
                    view.SetSectionBox(sb);
                    view.IsSectionBoxActive = true;
                    FormworkDebugLog.Log(
                        $"  [Visual] EnableSectionBox: set BB " +
                        $"min=({minP.X:F2},{minP.Y:F2},{minP.Z:F2}) " +
                        $"max=({maxP.X:F2},{maxP.Y:F2},{maxP.Z:F2})");
                }
                else
                {
                    // 要素のBBoxが取得できなかった場合は切断ボックスを有効化しない。
                    // 空のBBoxを有効化すると型枠DSがクリップアウトされる。
                    FormworkDebugLog.Log(
                        "  [Visual] EnableSectionBox: no elements with BBox found → section box NOT activated");
                }
            }
            catch (Exception exSb)
            {
                FormworkDebugLog.Log($"  [Visual] EnableSectionBox EX: {exSb.Message}");
            }
        }

        /// <summary>
        /// 28Tools_FormworkMarker が "28Tools_Formwork" で始まる既存 DirectShape を全て削除する
        /// （再実行時の累積を防ぐ）。MarkerValue / MarkerValueExcluded / 旧バージョンの
        /// "28Tools_Formwork_Steel" 等を全てカバーする。
        /// </summary>
        internal static void CleanupExistingFormworkShapes(Document doc, string sourceViewNameFilter = null)
        {
            var toDelete = new List<ElementId>();
            int total = 0, taggedMatch = 0, untagged = 0, otherView = 0;
            // 新カテゴリ (FormworkCategory) と旧カテゴリ (LegacyFormworkCategory) の両方を走査。
            // 旧バージョンで作成された GenericModel DS も対象にしてマイグレーション。
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(DirectShape))
                .WherePasses(new LogicalOrFilter(
                    new ElementCategoryFilter(FormworkParameterManager.FormworkCategory),
                    new ElementCategoryFilter(FormworkParameterManager.LegacyFormworkCategory)));
            foreach (Element e in collector)
            {
                try
                {
                    var p = e.LookupParameter(FormworkParameterManager.ParamMarker);
                    if (p == null || p.StorageType != StorageType.String) continue;
                    string val = p.AsString();
                    if (string.IsNullOrEmpty(val) ||
                        !val.StartsWith(FormworkParameterManager.MarkerValue)) continue;

                    total++;

                    // ソースビューフィルタが指定されている場合: 一致するもの + タグなし (旧形式・移行)
                    // を削除。一致しない (他のソースビュー) はそのまま残す。
                    if (!string.IsNullOrEmpty(sourceViewNameFilter))
                    {
                        var pv = e.LookupParameter(FormworkParameterManager.ParamSourceView);
                        string viewVal = pv?.AsString() ?? string.Empty;
                        if (string.IsNullOrEmpty(viewVal))
                        {
                            // タグなし (旧バージョン・パラメータ未バインドの場合) → 削除
                            untagged++;
                            toDelete.Add(e.Id);
                            continue;
                        }
                        if (viewVal != sourceViewNameFilter)
                        {
                            otherView++;
                            continue;
                        }
                        taggedMatch++;
                    }

                    toDelete.Add(e.Id);
                }
                catch { }
            }
            foreach (var id in toDelete)
            {
                try { doc.Delete(id); } catch { }
            }
            FormworkDebugLog.Log(
                $"  [Visual] CleanupExistingFormworkShapes filter='{sourceViewNameFilter ?? "(all)"}' " +
                $"total={total} match={taggedMatch} untagged={untagged} otherView={otherView} deleted={toDelete.Count}");
        }

        /// <summary>
        /// 指定名のビューが既に存在すれば削除する（再実行時の衝突を回避）。
        /// </summary>
        private static void DeleteViewByName(Document doc, string name)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && v.Name == name)
                .Select(v => v.Id)
                .ToList();
            foreach (var id in existing)
            {
                try { doc.Delete(id); } catch { }
            }
        }

        /// <summary>
        /// 元躯体要素を落ち着いた灰色 (RGB 94,94,94) + 20% 透過で表示。
        /// 塗潰しパターンが無い場合でもサーフェスの色味は設定される。
        /// </summary>
        private static void ApplySourceElementAppearance(Document doc, View3D view, FormworkResult result)
        {
            try
            {
                var gray = new Color(94, 94, 94);
                var ogs = new OverrideGraphicSettings();
                ogs.SetSurfaceTransparency(50);  // formwork が前景で目立つよう半透明
                ogs.SetSurfaceForegroundPatternColor(gray);
                ogs.SetSurfaceBackgroundPatternColor(gray);
                ogs.SetProjectionLineColor(gray);

                ElementId solidFillId = GetDraftingSolidFillPatternId(doc);
                if (solidFillId != null && solidFillId != ElementId.InvalidElementId)
                {
                    ogs.SetSurfaceForegroundPatternId(solidFillId);
                    ogs.SetSurfaceForegroundPatternVisible(true);
                }

                // リンク要素はホスト側からは個別に override 不可なのでスキップ。
                // ホスト要素のみ半透明グレーで表示。
                var registry = result.SourceRegistry as Engine.ElementSourceRegistry;
                foreach (var er in result.ElementResults)
                {
                    try
                    {
                        var src = registry?.Get(er.ElementId);
                        if (src != null && src.IsLinked) continue;
                        var id = new ElementId(er.ElementId);
                        if (doc.GetElement(id) != null)
                            view.SetElementOverrides(id, ogs);
                    }
                    catch { }
                }
            }
            catch { }
        }

        /// <summary>
        /// 部分接触がある面の DirectShape を半透明(50%)で表示する。
        /// View Filter による色分けの上からオーバーライドされる。
        /// </summary>
        private static void ApplyPartialContactOverride(View3D view, ElementId dsId)
        {
            var ogs = new OverrideGraphicSettings();
            ogs.SetSurfaceTransparency(50);
            view.SetElementOverrides(dsId, ogs);
        }

        /// <summary>
        /// 部分接触の合計面積が面全体の 5% を超えるかを判定する。
        /// 5% 以下なら誤差として扱い、半透明オーバーライドを適用しない
        /// (ほぼ全面が型枠なのに 1 面だけ半透明だとユーザーが混乱するため)。
        /// </summary>
        private static bool IsSignificantPartialContact(FaceClassifier.FaceInfo fi)
        {
            if (fi == null || fi.Face == null || fi.PartialContacts == null
                || fi.PartialContacts.Count == 0) return false;
            double faceArea = fi.Area;
            if (faceArea <= 1e-6) return false;
            double partialSum = 0;
            foreach (var pc in fi.PartialContacts) partialSum += pc.ContactArea;
            return (partialSum / faceArea) > 0.05;
        }

        /// <summary>
        /// BoundingBoxXYZ をトランスフォーム適用後のワールド AABB に変換する (8 頂点を変換して再計算)。
        /// リンク要素の BB をホスト座標系に正規化するのに使用。
        /// </summary>
        private static BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bb, Transform xform)
        {
            if (bb == null) return null;
            if (xform == null || xform.IsIdentity) return bb;
            var corners = new XYZ[]
            {
                new XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
                new XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
                new XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
                new XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
                new XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
                new XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
                new XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
                new XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
            };
            double mnX = double.MaxValue, mnY = double.MaxValue, mnZ = double.MaxValue;
            double mxX = double.MinValue, mxY = double.MinValue, mxZ = double.MinValue;
            foreach (var c in corners)
            {
                var p = xform.OfPoint(c);
                if (p.X < mnX) mnX = p.X;
                if (p.Y < mnY) mnY = p.Y;
                if (p.Z < mnZ) mnZ = p.Z;
                if (p.X > mxX) mxX = p.X;
                if (p.Y > mxY) mxY = p.Y;
                if (p.Z > mxZ) mxZ = p.Z;
            }
            return new BoundingBoxXYZ
            {
                Min = new XYZ(mnX, mnY, mnZ),
                Max = new XYZ(mxX, mxY, mxZ),
            };
        }

        private static ElementId GetDraftingSolidFillPatternId(Document doc)
        {
            var fps = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>();
            foreach (var fp in fps)
            {
                var p = fp.GetFillPattern();
                if (p == null) continue;
                if (p.IsSolidFill && p.Target == FillPatternTarget.Drafting) return fp.Id;
            }
            foreach (var fp in fps)
            {
                var p = fp.GetFillPattern();
                if (p == null) continue;
                if (p.IsSolidFill) return fp.Id;
            }
            return ElementId.InvalidElementId;
        }

        private static Solid CreateThinSolidFromFace(FaceClassifier.FaceInfo fi)
        {
            try
            {
                IList<CurveLoop> edges;
                // 直径 ≤300mm の小さな穴 (内側ループ) はフィルタして塞ぐ。
                // FormworkRequired 平面のみ対象 (曲面は ExcludeCurvedFormworkFaces で除外済み)。
                if (fi.FaceType == FaceType.FormworkRequired && fi.Face is PlanarFace)
                {
                    var (filtered, _, _) = Engine.SmallHoleFiller.Process(fi.Face);
                    edges = filtered;
                }
                else
                {
                    edges = fi.Face.GetEdgesAsCurveLoops();
                }
                if (edges == null || edges.Count == 0)
                {
                    FormworkDebugLog.Log($"    [ThinSolid] no-edges area={fi.Area:F4}");
                    return null;
                }

                XYZ n = fi.Normal?.Normalize();
                if (n == null)
                {
                    FormworkDebugLog.Log($"    [ThinSolid] no-normal area={fi.Area:F4}");
                    return null;
                }

                double thickness = 0.05;  // 約 15mm (視認性向上のため厚めに)
                string ex1Msg = null, ex2Msg = null, ex3Msg = null, ex4Msg = null;

                try
                {
                    return GeometryCreationUtilities.CreateExtrusionGeometry(edges, n, thickness);
                }
                catch (Exception ex1) { ex1Msg = ex1.Message; }

                try
                {
                    return GeometryCreationUtilities.CreateExtrusionGeometry(edges, n.Negate(), thickness);
                }
                catch (Exception ex2) { ex2Msg = ex2.Message; }

                // フォールバック: CurveLoop の向きを反転して再試行
                // (Revit が GetEdgesAsCurveLoops で返した向きが face normal と
                //  整合していないケースの救済)
                List<CurveLoop> reversed = null;
                try
                {
                    reversed = new List<CurveLoop>();
                    foreach (CurveLoop loop in edges)
                    {
                        var curves = new List<Curve>();
                        foreach (Curve c in loop) curves.Add(c.CreateReversed());
                        curves.Reverse();
                        reversed.Add(CurveLoop.Create(curves));
                    }
                }
                catch (Exception ex) { ex3Msg = "build-reversed:" + ex.Message; }

                if (reversed != null)
                {
                    try
                    {
                        return GeometryCreationUtilities.CreateExtrusionGeometry(reversed, n, thickness);
                    }
                    catch (Exception ex3) { ex3Msg = ex3.Message; }

                    try
                    {
                        return GeometryCreationUtilities.CreateExtrusionGeometry(reversed, n.Negate(), thickness);
                    }
                    catch (Exception ex4) { ex4Msg = ex4.Message; }
                }

                // フォールバック5/6: エッジを Tessellate で点列化し、PlanarFace 平面に
                // スナップした直線ポリラインで CurveLoop を再構築。
                // - extrudeProj 失敗 (微小な曲線/トレランス不整合)
                // - Non-planar CurveLoop (浮動小数点誤差で平面外)
                // の両方を救済。
                string ex5Msg = null, ex6Msg = null;
                List<CurveLoop> rebuilt = TryRebuildLoopsAsPolyline(edges, fi.Face as PlanarFace, n);
                if (rebuilt != null && rebuilt.Count > 0)
                {
                    try
                    {
                        return GeometryCreationUtilities.CreateExtrusionGeometry(rebuilt, n, thickness);
                    }
                    catch (Exception ex5) { ex5Msg = ex5.Message; }

                    try
                    {
                        return GeometryCreationUtilities.CreateExtrusionGeometry(rebuilt, n.Negate(), thickness);
                    }
                    catch (Exception ex6) { ex6Msg = ex6.Message; }
                }
                else
                {
                    ex5Msg = "rebuild-null";
                }

                FormworkDebugLog.Log(
                    $"    [ThinSolid] EXTRUDE_FAIL area={fi.Area:F4} loops={edges.Count} " +
                    $"n=({n.X:F3},{n.Y:F3},{n.Z:F3}) " +
                    $"ex1={ex1Msg} ex2={ex2Msg} ex3={ex3Msg} ex4={ex4Msg} ex5={ex5Msg} ex6={ex6Msg}");
                return null;
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"    [ThinSolid] OUTER_EX area={fi?.Area:F4} ex={ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// CurveLoop を Tessellate で点列化し、PlanarFace 平面 (pf != null の場合)
        /// にスナップした直線ポリラインで再構築する。
        /// 微小な曲線・浮動小数点誤差で extrudeProj が失敗するケースの最終救済策。
        /// pf == null (非PlanarFace) でも、点列を tangent plane (normal n) に投影して
        /// 平面性を強制する。
        /// </summary>
        private static List<CurveLoop> TryRebuildLoopsAsPolyline(
            IList<CurveLoop> srcLoops, PlanarFace pf, XYZ n)
        {
            if (srcLoops == null || n == null) return null;
            try
            {
                XYZ planeOrigin;
                if (pf != null && pf.Origin != null)
                {
                    planeOrigin = pf.Origin;
                }
                else
                {
                    XYZ first = null;
                    foreach (CurveLoop l in srcLoops)
                    {
                        foreach (Curve c in l) { first = c.GetEndPoint(0); break; }
                        if (first != null) break;
                    }
                    if (first == null)
                    {
                        FormworkDebugLog.Log($"      [Rebuild] null: no-first-point");
                        return null;
                    }
                    planeOrigin = first;
                }

                var result = new List<CurveLoop>();
                int li = 0;
                foreach (CurveLoop loop in srcLoops)
                {
                    var pts = new List<XYZ>();
                    int ci = 0;
                    foreach (Curve c in loop)
                    {
                        IList<XYZ> tess;
                        try { tess = c.Tessellate(); } catch (Exception ex) { tess = null; FormworkDebugLog.Log($"      [Rebuild] tess-ex loop={li} curve={ci} ex={ex.Message}"); }
                        if (tess == null || tess.Count < 2)
                        {
                            FormworkDebugLog.Log($"      [Rebuild] null: tess-fail loop={li} curve={ci} tess={tess?.Count}");
                            return null;
                        }
                        for (int i = 0; i < tess.Count - 1; i++) pts.Add(tess[i]);
                        ci++;
                    }
                    FormworkDebugLog.Log($"      [Rebuild] loop={li} pts={pts.Count} pf={pf != null} origin=({planeOrigin.X:F3},{planeOrigin.Y:F3},{planeOrigin.Z:F3})");
                    if (pts.Count < 3)
                    {
                        FormworkDebugLog.Log($"      [Rebuild] null: pts<3");
                        return null;
                    }

                    var snapped = new List<XYZ>(pts.Count);
                    foreach (XYZ p in pts)
                    {
                        double d = (p - planeOrigin).DotProduct(n);
                        snapped.Add(p - n * d);
                    }

                    // Revit の ShortCurveTolerance (≈0.00256 ft) より大きく取る必要あり。
                    // 0.003 ft ≈ 0.9mm で Line.CreateBound 拒否を回避。
                    const double tol = 0.003;
                    var cleaned = new List<XYZ>();
                    cleaned.Add(snapped[0]);
                    for (int i = 1; i < snapped.Count; i++)
                    {
                        if (snapped[i].DistanceTo(cleaned[cleaned.Count - 1]) > tol)
                            cleaned.Add(snapped[i]);
                    }
                    while (cleaned.Count >= 2 &&
                        cleaned[cleaned.Count - 1].DistanceTo(cleaned[0]) <= tol)
                    {
                        cleaned.RemoveAt(cleaned.Count - 1);
                    }
                    FormworkDebugLog.Log($"      [Rebuild] loop={li} cleaned={cleaned.Count}");
                    if (cleaned.Count < 3)
                    {
                        FormworkDebugLog.Log($"      [Rebuild] null: cleaned<3");
                        return null;
                    }

                    var lines = new List<Curve>();
                    for (int i = 0; i < cleaned.Count; i++)
                    {
                        XYZ p0 = cleaned[i];
                        XYZ p1 = cleaned[(i + 1) % cleaned.Count];
                        if (p0.DistanceTo(p1) <= tol) continue;
                        try { lines.Add(Line.CreateBound(p0, p1)); }
                        catch (Exception ex)
                        {
                            FormworkDebugLog.Log($"      [Rebuild] null: line-create-ex i={i} ex={ex.Message}");
                            return null;
                        }
                    }
                    FormworkDebugLog.Log($"      [Rebuild] loop={li} lines={lines.Count}");
                    if (lines.Count < 3)
                    {
                        FormworkDebugLog.Log($"      [Rebuild] null: lines<3");
                        return null;
                    }

                    CurveLoop newLoop;
                    try { newLoop = CurveLoop.Create(lines); }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"      [Rebuild] null: CurveLoop.Create-ex ex={ex.Message}");
                        return null;
                    }
                    result.Add(newLoop);
                    li++;
                }
                if (result.Count == 0) FormworkDebugLog.Log($"      [Rebuild] null: result-empty");
                return result.Count > 0 ? result : null;
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"      [Rebuild] null: outer-catch ex={ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Clipper が失敗したケースの 3D Boolean Difference fallback。
        /// face 全体の薄板 Solid を作り、各 PartialContact の UvBoundsOnA を
        /// 「貫通する角材」として Boolean 差分でくり抜く。
        /// 非矩形 face (interior-points / has-holes) でも 3D Boolean なら対応可能。
        ///
        /// 1 個も差分できなかった場合 (UvBoundsOnA がすべて null 等) は null を返し、
        /// 呼び出し側は従来の半透明 fallback に流れる。
        /// </summary>
        private static Solid TryBuildCarvedFaceSolid(FaceClassifier.FaceInfo fi)
        {
            if (fi == null || fi.Face == null || fi.Normal == null) return null;
            if (fi.PartialContacts == null || fi.PartialContacts.Count == 0) return null;

            Solid baseSolid = CreateThinSolidFromFace(fi);
            if (baseSolid == null) return null;

            XYZ normal = fi.Normal.Normalize();
            Solid currentSolid = baseSolid;
            int subtractCount = 0;

            foreach (var pc in fi.PartialContacts)
            {
                if (FormworkDebugLog.Enabled)
                    FormworkDebugLog.Log($"  [BoolDiff] pc E{pc.OtherElementId}/f{pc.OtherFaceIndex} UvBoundsOnA={(pc.UvBoundsOnA == null ? "NULL" : $"[U:{pc.UvBoundsOnA.Min.U:F3}..{pc.UvBoundsOnA.Max.U:F3},V:{pc.UvBoundsOnA.Min.V:F3}..{pc.UvBoundsOnA.Max.V:F3}]")} hasContactFaceB={(pc.ContactFaceB != null)}");

                // ポリゴンカッター（ContactFaceB の実形状から押し出し）を最優先で使用する。
                // これにより L 字形など非矩形の接触面を UV-AABB 矩形でオーバーカットする問題を回避する。
                Solid cutter = null;
                if (pc.ContactFaceB != null)
                {
                    cutter = BuildCutterFromContactFacePolygon(pc.ContactFaceB, normal);
                    if (FormworkDebugLog.Enabled)
                        FormworkDebugLog.Log($"    [BoolDiff] polygon-cutter={(cutter != null ? "OK" : "null")}");
                }
                // フォールバック: ポリゴンカッターが失敗した場合に UV 矩形カッターを使用
                if (cutter == null && pc.UvBoundsOnA != null)
                {
                    cutter = BuildCutterFromUvRect(fi.Face, pc.UvBoundsOnA, normal);
                    if (FormworkDebugLog.Enabled)
                        FormworkDebugLog.Log($"    [BoolDiff] uvRect-cutter-fallback={(cutter != null ? "OK" : "null")}");
                }
                if (cutter == null) continue;

                try
                {
                    double volBefore = currentSolid.Volume;
                    Solid result = BooleanOperationsUtils.ExecuteBooleanOperation(
                        currentSolid, cutter, BooleanOperationsType.Difference);
                    double volAfter = result != null ? result.Volume : -1.0;
                    if (FormworkDebugLog.Enabled)
                        FormworkDebugLog.Log($"    [BoolDiff] subtract volBefore={volBefore:F6} volAfter={volAfter:F6} delta={volBefore-volAfter:F6}");
                    if (result != null && result.Volume > 1e-9)
                    {
                        currentSolid = result;
                        subtractCount++;
                    }
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log(
                        $"  [BoolDiff] EX face area={fi.Area:F4} pc.area={pc.ContactArea:F4} ex={ex.Message}");
                }
            }

            if (subtractCount == 0) return null;
            if (currentSolid.Volume < 1e-9) return null;

            return currentSolid;
        }

        /// <summary>
        /// 接触面 B の実際のポリゴン形状から「貫通カッター」Solid を生成する。
        /// UV-AABB 矩形カッターと異なり、L 字形などの非矩形接触面に対して
        /// オーバーカットを起こさない。
        /// </summary>
        private static Solid BuildCutterFromContactFacePolygon(Face contactFaceB, XYZ faceANormal)
        {
            IList<CurveLoop> loops = null;
            try { loops = contactFaceB.GetEdgesAsCurveLoops(); } catch { return null; }
            if (loops == null || loops.Count == 0) return null;

            // 面 A の法線方向に prePad だけ手前からカッターを始め、thickness だけ押し出す。
            // baseSolid (厚み 0.05 ft) を確実に包含するよう設定する。
            const double prePad = 0.10;
            const double thickness = 0.20;

            XYZ dir = faceANormal.Normalize();
            var transform = Transform.CreateTranslation(dir.Multiply(-prePad));

            // 接触面 B の法線が押し出し方向(face A 法線) と反対向きの場合
            // (= 通常の接触ペア)、GetEdgesAsCurveLoops が返すループは
            // 押し出し方向から見ると CW になる。CreateExtrusionGeometry は
            // CCW を要求するため、その場合は事前にループを反転する。
            XYZ faceBNormal = null;
            try
            {
                var pf = contactFaceB as PlanarFace;
                if (pf != null) faceBNormal = pf.FaceNormal.Normalize();
            }
            catch { }
            bool needsReverse = (faceBNormal != null && faceBNormal.DotProduct(dir) < 0);

            var translatedLoops = new List<CurveLoop>(loops.Count);
            foreach (var l in loops)
            {
                CurveLoop tl = null;
                try { tl = CurveLoop.CreateViaTransform(l, transform); } catch { return null; }
                if (tl == null) return null;
                if (needsReverse)
                {
                    try
                    {
                        var curves = new List<Curve>();
                        foreach (var c in tl) curves.Add(c.CreateReversed());
                        curves.Reverse();
                        tl = CurveLoop.Create(curves);
                    }
                    catch { return null; }
                }
                translatedLoops.Add(tl);
            }

            Solid cutter = null;
            try
            {
                cutter = GeometryCreationUtilities.CreateExtrusionGeometry(translatedLoops, dir, thickness);
            }
            catch (Exception ex)
            {
                if (FormworkDebugLog.Enabled)
                    FormworkDebugLog.Log($"    [PolyCutter] EXTRUDE-EX needsReverse={needsReverse} ex={ex.Message}");

                // 反転して再試行
                try
                {
                    var altLoops = new List<CurveLoop>(translatedLoops.Count);
                    foreach (var l in translatedLoops)
                    {
                        var curves = new List<Curve>();
                        foreach (var c in l) curves.Add(c.CreateReversed());
                        curves.Reverse();
                        altLoops.Add(CurveLoop.Create(curves));
                    }
                    cutter = GeometryCreationUtilities.CreateExtrusionGeometry(altLoops, dir, thickness);
                }
                catch { return null; }
            }

            if (FormworkDebugLog.Enabled && cutter != null)
                FormworkDebugLog.Log($"    [PolyCutter] vol={cutter.Volume:F6} faceBNormal=({(faceBNormal != null ? $"{faceBNormal.X:F2},{faceBNormal.Y:F2},{faceBNormal.Z:F2}" : "null")}) dir=({dir.X:F2},{dir.Y:F2},{dir.Z:F2}) needsReverse={needsReverse} loops={translatedLoops.Count}");
            return cutter;
        }

        /// <summary>
        /// face の UV 矩形から「面の前後を貫通する」厚みのある角材 Solid を作る。
        /// baseSolid (厚み 0.05) を確実に貫通するよう、面の前 prePad ft から
        /// thickness ft 厚に伸ばすことで baseSolid の法線範囲 [0, 0.05] を包含する。
        /// </summary>
        private static Solid BuildCutterFromUvRect(Face face, BoundingBoxUV uvBounds, XYZ normal)
        {
            try
            {
                XYZ p00, p10, p11, p01;
                try
                {
                    p00 = face.Evaluate(uvBounds.Min);
                    p10 = face.Evaluate(new UV(uvBounds.Max.U, uvBounds.Min.V));
                    p11 = face.Evaluate(uvBounds.Max);
                    p01 = face.Evaluate(new UV(uvBounds.Min.U, uvBounds.Max.V));
                }
                catch { return null; }
                if (p00 == null || p10 == null || p11 == null || p01 == null) return null;

                const double minEdge = 1e-4;
                if (p00.DistanceTo(p10) < minEdge) return null;
                if (p10.DistanceTo(p11) < minEdge) return null;
                if (p11.DistanceTo(p01) < minEdge) return null;
                if (p01.DistanceTo(p00) < minEdge) return null;

                // 面の前 prePad ft から後ろに thickness ft 伸ばす
                // → 範囲 [-prePad, thickness - prePad] in normal 方向
                // → baseSolid 範囲 [0, 0.05] を包含するため貫通する
                const double prePad = 0.10;     // 30mm 手前から
                const double thickness = 0.20;  // 60mm extrude

                XYZ offset = normal.Multiply(-prePad);
                XYZ q00 = p00 + offset;
                XYZ q10 = p10 + offset;
                XYZ q11 = p11 + offset;
                XYZ q01 = p01 + offset;

                // CCW (viewed from +normal) で CurveLoop 構築
                var loop = new CurveLoop();
                loop.Append(Line.CreateBound(q00, q10));
                loop.Append(Line.CreateBound(q10, q11));
                loop.Append(Line.CreateBound(q11, q01));
                loop.Append(Line.CreateBound(q01, q00));

                try
                {
                    return GeometryCreationUtilities.CreateExtrusionGeometry(
                        new[] { loop }, normal, thickness);
                }
                catch
                {
                    // CW 方向だった場合: ループを反転して再試行
                    try
                    {
                        var revLoop = new CurveLoop();
                        revLoop.Append(Line.CreateBound(q00, q01));
                        revLoop.Append(Line.CreateBound(q01, q11));
                        revLoop.Append(Line.CreateBound(q11, q10));
                        revLoop.Append(Line.CreateBound(q10, q00));
                        return GeometryCreationUtilities.CreateExtrusionGeometry(
                            new[] { revLoop }, normal, thickness);
                    }
                    catch { return null; }
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// ドキュメントの表示名を返す。パスがあればファイル名(拡張子なし)、なければタイトル。
        /// </summary>
        private static string GetDocumentDisplayName(Document doc)
        {
            try
            {
                if (doc == null) return ElementSourceRegistry.HostSourceName;
                var path = doc.PathName;
                if (!string.IsNullOrEmpty(path))
                    return Path.GetFileNameWithoutExtension(path);
                return !string.IsNullOrEmpty(doc.Title)
                    ? doc.Title : ElementSourceRegistry.HostSourceName;
            }
            catch { return ElementSourceRegistry.HostSourceName; }
        }
    }
}

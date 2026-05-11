using System;
using System.Collections.Generic;
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
        internal const string AnalysisViewName = "型枠分析";
        internal class VisualizerResult
        {
            public View3D AnalysisView;
            public List<ElementId> CreatedShapeIds = new List<ElementId>();
            public Dictionary<string, (byte R, byte G, byte B)> KeyColors
                = new Dictionary<string, (byte, byte, byte)>();
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
            View sourceView = null)
        {
            var vr = new VisualizerResult();

            // 既存の同名ビューを削除
            DeleteViewByName(doc, AnalysisViewName);

            // 既存の型枠 DirectShape を削除（再実行時の累積を防ぐ）
            CleanupExistingFormworkShapes(doc);

            var view3DType = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(v => v.ViewFamily == ViewFamily.ThreeDimensional);
            if (view3DType == null) return vr;

            var view = View3D.CreateIsometric(doc, view3DType.Id);
            try { view.Name = AnalysisViewName; } catch { }
            vr.AnalysisView = view;

            // 表示スタイルを Shading に明示設定 (Realistic では OverrideGraphicSettings の
            // 透過・色オーバーライドが期待通りに反映されない場合があるため)。
            try { view.DisplayStyle = DisplayStyle.Shading; }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] DisplayStyle set EX: {ex.Message}");
            }
            try { view.DetailLevel = ViewDetailLevel.Fine; } catch { }

            // 尺度を 1/200 に設定
            try { view.Scale = 200; }
            catch (Exception ex) { FormworkDebugLog.Log($"  [Visual] Scale set EX: {ex.Message}"); }

            // 視点をビューキューブの青い角（右前上からのアイソメトリック）に設定
            SetIsometricOrientation(view);

            // OST_GenericModel (DirectShape のカテゴリ) を明示的に表示状態にする。
            // ビューテンプレートやデフォルト設定で非表示になっている場合への防御。
            try
            {
                var gmCatId = new ElementId(BuiltInCategory.OST_GenericModel);
                if (view.CanCategoryBeHidden(gmCatId))
                {
                    bool wasHidden = view.GetCategoryHidden(gmCatId);
                    view.SetCategoryHidden(gmCatId, false);
                    FormworkDebugLog.Log(
                        $"  [Visual] OST_GenericModel category was hidden={wasHidden} → set visible");
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] OST_GenericModel category set EX: {ex.Message}");
            }

            // 視点を固定アイソメトリックに設定し、ロックする（ユーザーが誤って回転しないよう）
            try { view.SaveOrientationAndLock(); }
            catch (Exception ex) { FormworkDebugLog.Log($"  [Visual] SaveOrientationAndLock EX: {ex.Message}"); }

            // 接触検出込みで面を再計算（幾何学的検査なので rayView 不要）
            var facesByElement = FormworkCalcEngine.RecomputeFaces(doc, result, settings);

            // [3] CurrentView スコープでソースが 3D ビューの場合は、ソースの切断ボックス
            // (位置・有効/無効) をそのまま継承する。それ以外は要素 BoundingBox から自動算出。
            bool useSourceSectionBox =
                settings.Scope == CalculationScope.CurrentView && sourceView is View3D;
            if (useSourceSectionBox)
                CopySectionBoxFromSource((View3D)sourceView, view);
            else
                EnableSectionBox(doc, view, result);

            // [4] CurrentView スコープでソースビューが指定されていれば、V/G 上書き
            // (モデル/注釈カテゴリの表示/非表示・色等) をターゲット 3D ビューに継承。
            // それ以外はノイズになる切断ボックスのアウトラインとレベル線を非表示化。
            if (settings.Scope == CalculationScope.CurrentView && sourceView != null)
                CopyCategoryVisibilityAndOverrides(doc, sourceView, view);
            else
                HideClutterCategories(view);

            // 元躯体要素を 20% 透過 + RGB(94,94,94) で表示
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

            foreach (var er in result.ElementResults)
            {
                if (!facesByElement.TryGetValue(er.ElementId, out var faces)) continue;

                string key = GetKey(er, settings);
                typeIdByKey.TryGetValue(key, out var typeId);
                catLabelByKey.TryGetValue(key, out var catLabel);
                catLabel = catLabel ?? CategoryLabel(er.Category);

                // 要素のレベル名を取得
                string levelName = string.Empty;
                try
                {
                    var srcElem = doc.GetElement(new ElementId(er.ElementId));
                    levelName = Engine.ElementCollector.GetElementLevelName(srcElem);
                }
                catch { }

                // 各 FormworkRequired 面の有効面積 (fi.EffectiveAreaM2) を、その面に対応する
                // DirectShape (Clipper 分割なら最初の 1 個) に割り当てる。
                // これにより、ユーザーが個々の DirectShape を選択しても正しい面積が表示される。
                //
                // 開口部の調整 (OpeningEdgeAreaAdded - OpeningAreaDeducted) は要素単位の値で
                // 面単位に分配できないため、最初の FormworkRequired DirectShape に乗せる。
                // 結果として:
                //   sum(全 DirectShape の面積) = sum(fi.EffectiveAreaM2) + 開口調整
                //                              = er.FormworkArea
                // 集計表の総合計は維持される。
                double openingDelta = er.OpeningEdgeAreaAdded - er.OpeningAreaDeducted;
                bool firstFormworkAssigned = false;
                // 合計一致保証用: 実際に DS パラメータに書き込んだ面積の合計と最初の DS の ID
                double assignedAreaM2 = 0;
                ElementId firstFormworkDsId = null;

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
                    // 同一面の中で最初のピース (Clipper 分割時の先頭) のみに面積を乗せる
                    bool firstPieceForThisFace = true;

                    foreach (var solid in partSolids)
                    {
                        ElementId dsId = null;
                        double areaM2ThisDs = 0.0;
                        try
                        {
                            var catOst = new ElementId(BuiltInCategory.OST_GenericModel);
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
                                // 面単位の有効面積を、その面の最初の DirectShape ピースに割り当てる。
                                // Clipper 分割で複数ピースになった場合、2 個目以降は 0。
                                // 開口調整は要素全体の最初の FormworkRequired DirectShape に乗せる。
                                double areaM2 = 0.0;
                                if (isFormworkRequired && firstPieceForThisFace)
                                {
                                    areaM2 = fi.EffectiveAreaM2;
                                    if (!firstFormworkAssigned)
                                    {
                                        areaM2 += openingDelta;
                                        firstFormworkAssigned = true;
                                    }
                                    firstPieceForThisFace = false;
                                }
                                string filterKey = IsDeducted(fi.FaceType) ? "控除面" : key;
                                FormworkParameterManager.SetInstanceValues(
                                    ds, catLabel, levelName, filterKey, areaM2, faceHasPartialContact);
                                areaM2ThisDs = areaM2;
                            }
                            catch { }

                            dsId = ds.Id;
                        }
                        catch { continue; }

                        if (dsId != null)
                        {
                            vr.CreatedShapeIds.Add(dsId);
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
            }

            int formworkShapesCount = vr.CreatedShapeIds.Count;
            FormworkDebugLog.Log(
                $"  [Visual] formwork DirectShapes created: {formworkShapesCount} " +
                $"(elements={result.ElementResults.Count})");

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
            CreateExcludedShapes(doc, result, vr);
            int excludedShapesCount = vr.CreatedShapeIds.Count - formworkShapesCount;
            FormworkDebugLog.Log(
                $"  [Visual] total DirectShapes: {vr.CreatedShapeIds.Count} " +
                $"(formwork={formworkShapesCount} excluded={excludedShapesCount})");

            // View Filter による色分けを適用（個別要素オーバーライドは使わない）
            try
            {
                FormworkFilterManager.ApplyColorFilters(doc, view, keyAssignment);
            }
            catch { }

            return vr;
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
                            $"  [Visual:Diag] sample key='{key}' id={id.IntegerValue} {bbStr} " +
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
            VisualizerResult vr)
        {
            if (result.ExcludedResults == null || result.ExcludedResults.Count == 0)
                return;

            int created = 0;
            foreach (var ex in result.ExcludedResults)
            {
                Element src = null;
                try { src = doc.GetElement(new ElementId(ex.ElementId)); } catch { }
                if (src == null) continue;

                var solids = SolidUnionProcessor.GetSolids(src);
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
                        var catOst = new ElementId(BuiltInCategory.OST_GenericModel);
                        var ds = DirectShape.CreateElement(doc, catOst);
                        ds.ApplicationId = "Tools28";
                        ds.ApplicationDataId = "FormworkExcluded";
                        ds.SetShape(new GeometryObject[] { solid });

                        try
                        {
                            FormworkParameterManager.SetInstanceValues(
                                ds,
                                FormworkParameterManager.MarkerValueExcluded,
                                label,
                                levelName,
                                FormworkParameterManager.ExcludedGroupKey,
                                0.0,
                                false);
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

            foreach (var v in allViews)
            {
                if (v is ViewSchedule) continue;
                if (v.ViewType == ViewType.Legend) continue;

                try
                {
                    v.HideElements(idsList);
                }
                catch { }
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
                var catId = new ElementId(BuiltInCategory.OST_GenericModel);
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
                default: return "その他";
            }
        }

        private static bool IsDeducted(FaceType t)
        {
            return t == FaceType.DeductedTop ||
                   t == FaceType.DeductedBottom ||
                   t == FaceType.DeductedContact ||
                   t == FaceType.DeductedBelowGL;
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
        /// 有効/無効状態をそのままコピーする。これによりユーザーがソースビューで設定した
        /// 切断範囲が解析ビューに引き継がれる。
        /// </summary>
        private static void CopySectionBoxFromSource(View3D source, View3D target)
        {
            try
            {
                var sb = source.GetSectionBox();
                if (sb != null)
                {
                    var copy = new BoundingBoxXYZ
                    {
                        Min = sb.Min,
                        Max = sb.Max,
                        Transform = sb.Transform,
                    };
                    target.SetSectionBox(copy);
                }
                target.IsSectionBoxActive = source.IsSectionBoxActive;
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Visual] CopySectionBoxFromSource EX: {ex.Message}");
            }
        }

        /// <summary>
        /// ソースビューの V/G 上書き設定 (モデル/注釈カテゴリの表示/非表示・色) を
        /// ターゲット 3D ビューにコピーする。ターゲットで非表示にできないカテゴリは
        /// スキップ (例: 平面ビュー専用のカテゴリ)。
        ///
        /// 解析ビューに必要なカテゴリ (OST_GenericModel = DirectShape) は強制的に表示にする。
        /// 切断ボックスのアウトライン・レベル線は視覚的ノイズなので強制非表示。
        /// </summary>
        private static void CopyCategoryVisibilityAndOverrides(
            Document doc, View source, View3D target)
        {
            try
            {
                var allCategories = doc.Settings.Categories;
                foreach (Category cat in allCategories)
                {
                    if (cat == null) continue;
                    var catId = cat.Id;
                    if (!target.CanCategoryBeHidden(catId)) continue;
                    try
                    {
                        // 表示/非表示
                        bool srcHidden = source.GetCategoryHidden(catId);
                        target.SetCategoryHidden(catId, srcHidden);

                        // V/G 上書き (色・線種・透過等)
                        if (target.IsCategoryOverridable(catId)
                            && source.IsCategoryOverridable(catId))
                        {
                            var ogs = source.GetCategoryOverrides(catId);
                            if (ogs != null)
                                target.SetCategoryOverrides(catId, ogs);
                        }
                    }
                    catch { /* per-category failures are non-fatal */ }

                    // サブカテゴリも同様にコピー
                    foreach (Category sub in cat.SubCategories)
                    {
                        if (sub == null) continue;
                        var subId = sub.Id;
                        if (!target.CanCategoryBeHidden(subId)) continue;
                        try
                        {
                            bool srcHidden = source.GetCategoryHidden(subId);
                            target.SetCategoryHidden(subId, srcHidden);
                            if (target.IsCategoryOverridable(subId)
                                && source.IsCategoryOverridable(subId))
                            {
                                var ogs = source.GetCategoryOverrides(subId);
                                if (ogs != null)
                                    target.SetCategoryOverrides(subId, ogs);
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

            // DirectShape (OST_GenericModel) は型枠表示の主役なので強制的に表示状態に上書き
            try
            {
                var gmId = new ElementId(BuiltInCategory.OST_GenericModel);
                if (target.CanCategoryBeHidden(gmId))
                    target.SetCategoryHidden(gmId, false);
            }
            catch { }

            // 切断ボックスのアウトライン・レベル線は解析ビューで邪魔なので強制非表示
            HideClutterCategories(target);
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

                foreach (var id in ids)
                {
                    var elem = doc.GetElement(new ElementId(id));
                    if (elem == null) continue;
                    BoundingBoxXYZ bb = null;
                    try { bb = elem.get_BoundingBox(null); } catch { }
                    if (bb == null) continue;
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
                }
                view.IsSectionBoxActive = true;
            }
            catch { }
        }

        /// <summary>
        /// 28Tools_FormworkMarker が "28Tools_Formwork" で始まる既存 DirectShape を全て削除する
        /// （再実行時の累積を防ぐ）。MarkerValue / MarkerValueExcluded / 旧バージョンの
        /// "28Tools_Formwork_Steel" 等を全てカバーする。
        /// </summary>
        private static void CleanupExistingFormworkShapes(Document doc)
        {
            var toDelete = new List<ElementId>();
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(DirectShape))
                .OfCategory(BuiltInCategory.OST_GenericModel);
            foreach (Element e in collector)
            {
                try
                {
                    var p = e.LookupParameter(FormworkParameterManager.ParamMarker);
                    if (p != null && p.StorageType == StorageType.String)
                    {
                        string val = p.AsString();
                        if (!string.IsNullOrEmpty(val) &&
                            val.StartsWith(FormworkParameterManager.MarkerValue))
                        {
                            toDelete.Add(e.Id);
                        }
                    }
                }
                catch { }
            }
            foreach (var id in toDelete)
            {
                try { doc.Delete(id); } catch { }
            }
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

                foreach (var er in result.ElementResults)
                {
                    try
                    {
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
                var edges = fi.Face.GetEdgesAsCurveLoops();
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
    }
}

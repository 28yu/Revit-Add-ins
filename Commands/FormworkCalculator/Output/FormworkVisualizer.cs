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

            // ソースが 3D ビューなら視点（カメラの向き）を継承する
            CopyOrientationIfPossible(sourceView, view);

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

                foreach (var fi in faces)
                {
                    bool visible =
                        fi.FaceType == FaceType.FormworkRequired ||
                        (settings.ShowDeductedFaces && IsDeducted(fi.FaceType));
                    if (!visible) continue;

                    bool faceHasPartialContact =
                        fi.FaceType == FaceType.FormworkRequired && fi.PartialContacts.Count > 0;

                    // Phase 2: 部分接触がある面は矩形差分で厳密な形状を作る。
                    //   Clipper が成功 → 半透明オーバーライドは不要 (形状が既に正確)
                    //   Clipper が失敗 → 従来通り面全体を作成 + 半透明オーバーライド (Phase 1)
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
                            FormworkDebugLog.Log(
                                $"Clipper fallback: E{er.ElementId} face type={fi.FaceType} " +
                                $"reason={clip.FailReason}");
                        }
                    }

                    if (partSolids == null)
                    {
                        Solid thin = CreateThinSolidFromFace(fi);
                        if (thin == null) continue;
                        partSolids = new List<Solid> { thin };
                    }

                    bool isFormworkRequired = (fi.FaceType == FaceType.FormworkRequired);
                    // 同一面の中で最初のピース (Clipper 分割時の先頭) のみに面積を乗せる
                    bool firstPieceForThisFace = true;

                    foreach (var solid in partSolids)
                    {
                        ElementId dsId = null;
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
                            }
                            catch { }

                            dsId = ds.Id;
                        }
                        catch { continue; }

                        if (dsId != null)
                        {
                            vr.CreatedShapeIds.Add(dsId);
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
                if (edges == null || edges.Count == 0) return null;

                XYZ n = fi.Normal?.Normalize();
                if (n == null) return null;

                double thickness = 0.05;  // 約 15mm (視認性向上のため厚めに)
                try
                {
                    return GeometryCreationUtilities.CreateExtrusionGeometry(edges, n, thickness);
                }
                catch
                {
                    return GeometryCreationUtilities.CreateExtrusionGeometry(edges, n.Negate(), thickness);
                }
            }
            catch
            {
                return null;
            }
        }
    }
}

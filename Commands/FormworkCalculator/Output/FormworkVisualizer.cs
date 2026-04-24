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

        private static readonly (byte R, byte G, byte B)[] _autoPalette = new (byte, byte, byte)[]
        {
            (230, 80, 80),   (80, 160, 230), (60, 200, 100), (240, 200, 60),
            (180, 100, 220), (255, 130, 170), (80, 210, 210), (255, 160, 60),
            (170, 220, 80), (60, 170, 170), (220, 100, 200), (120, 120, 220),
        };

        internal static VisualizerResult CreateVisualization(
            Document doc,
            FormworkResult result,
            FormworkSettings settings)
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

            // 接触検出込みで面を再計算（幾何学的検査なので rayView 不要）
            var facesByElement = FormworkCalcEngine.RecomputeFaces(doc, result, settings);

            // セクションボックスを有効化（要素全体を含むよう自動フィット）
            EnableSectionBox(doc, view, result);

            // 元躯体要素を 20% 透過 + RGB(94,94,94) で表示
            ApplySourceElementAppearance(doc, view, result);

            // 分類キーに基づく色割当
            var keyAssignment = AssignColors(result, settings);

            // 控除面表示時は「控除面」キーを追加（単色グレー）
            if (settings.ShowDeductedFaces)
                keyAssignment["控除面"] = (180, 180, 180);

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

                    // 部分接触面の控除後面積 (DirectShape ひとつに集約する場合用)
                    double effectiveFeetSq = fi.Area;
                    if (faceHasPartialContact)
                    {
                        double partial = 0;
                        foreach (var pc in fi.PartialContacts) partial += pc.ContactArea;
                        effectiveFeetSq = Math.Max(0, fi.Area - partial);
                    }

                    // 複数 Solid に分割する場合は、全体の面積を Solid 数で均等に配分しない。
                    // 代わりに最初の DirectShape に控除後面積を全て乗せる (集計側では要素毎に合算される)。
                    bool firstShape = true;
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
                                double areaM2 = firstShape
                                    ? UnitUtils.ConvertFromInternalUnits(effectiveFeetSq, UnitTypeId.SquareMeters)
                                    : 0.0;
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
                            // フォールバック時のみ半透明で「一部控除されている」ことを視覚化。
                            if (faceHasPartialContact && !clipperUsed)
                            {
                                try { ApplyPartialContactOverride(view, dsId); } catch { }
                            }
                        }
                        firstShape = false;
                    }
                }
            }

            // View Filter による色分けを適用（個別要素オーバーライドは使わない）
            try
            {
                FormworkFilterManager.ApplyColorFilters(doc, view, keyAssignment);
            }
            catch { }

            return vr;
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
        /// 解析対象要素の BoundingBox を全て囲む形でセクションボックスを設定し有効化。
        /// </summary>
        private static void EnableSectionBox(Document doc, View3D view, FormworkResult result)
        {
            try
            {
                XYZ minP = null, maxP = null;
                foreach (var er in result.ElementResults)
                {
                    var elem = doc.GetElement(new ElementId(er.ElementId));
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
        /// 28Tools_FormworkMarker パラメータに "28Tools_Formwork" 値を持つ
        /// 既存 DirectShape を全て削除する（再実行時の累積を防ぐ）。
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
                    if (p != null && p.StorageType == StorageType.String &&
                        p.AsString() == FormworkParameterManager.MarkerValue)
                    {
                        toDelete.Add(e.Id);
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
                ogs.SetSurfaceTransparency(20);
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

                double thickness = 0.03;  // 約 10mm
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

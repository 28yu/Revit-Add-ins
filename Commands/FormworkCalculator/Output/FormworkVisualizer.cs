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
    /// - 元の躯体要素の表示はそのまま（半透明等のオーバーライドは入れない）
    /// </summary>
    internal static class FormworkVisualizer
    {
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

            var view3DType = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(v => v.ViewFamily == ViewFamily.ThreeDimensional);
            if (view3DType == null) return vr;

            var view = View3D.CreateIsometric(doc, view3DType.Id);
            string viewName = $"型枠分析_{DateTime.Now:yyyyMMdd_HHmmss}";
            try { view.Name = viewName; } catch { }
            vr.AnalysisView = view;

            // 接触検出込みで面を再計算
            var facesByElement = FormworkCalcEngine.RecomputeFaces(doc, result, settings);

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

                foreach (var fi in faces)
                {
                    bool visible =
                        fi.FaceType == FaceType.FormworkRequired ||
                        (settings.ShowDeductedFaces && IsDeducted(fi.FaceType));
                    if (!visible) continue;

                    Solid thin = CreateThinSolidFromFace(fi);
                    if (thin == null) continue;

                    ElementId dsId = null;
                    try
                    {
                        var catOst = new ElementId(BuiltInCategory.OST_GenericModel);
                        var ds = DirectShape.CreateElement(doc, catOst);
                        ds.ApplicationId = "Tools28";
                        ds.ApplicationDataId = "Formwork";
                        ds.SetShape(new GeometryObject[] { thin });

                        if (typeId != null && typeId != ElementId.InvalidElementId)
                        {
                            try { ds.SetTypeId(typeId); } catch { }
                        }

                        try
                        {
                            double areaM2 = UnitUtils.ConvertFromInternalUnits(fi.Area, UnitTypeId.SquareMeters);
                            string filterKey = IsDeducted(fi.FaceType) ? "控除面" : key;
                            FormworkParameterManager.SetInstanceValues(ds, catLabel, filterKey, areaM2);
                        }
                        catch { }

                        dsId = ds.Id;
                    }
                    catch { continue; }

                    if (dsId != null) vr.CreatedShapeIds.Add(dsId);
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

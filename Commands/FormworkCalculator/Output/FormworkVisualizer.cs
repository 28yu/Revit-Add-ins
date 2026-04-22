using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Engine;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Output
{
    /// <summary>
    /// 型枠面を DirectShape 化して色分け3Dビューを作成する。
    /// - 分類キー毎に DirectShapeType を作成し、名前で識別可能にする
    /// - OverrideGraphicSettings で前景サーフェスパターンの可視化＋色付け
    /// - 既存ビューでは自動的に非表示にする（解析専用ビューのみ表示）
    /// </summary>
    internal static class FormworkVisualizer
    {
        internal class VisualizerResult
        {
            public View3D AnalysisView;
            public List<ElementId> CreatedShapeIds = new List<ElementId>();
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

            // GL 高さ
            double? glFeet = null;
            if (settings.UseGLDeduction)
            {
                glFeet = UnitUtils.ConvertToInternalUnits(settings.GLElevationMeters, UnitTypeId.Meters);
            }

            // 分類キーに基づく色割当
            var keyAssignment = AssignColors(result, settings);

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

            // 塗り潰しパターン（Drafting Solid fill）
            ElementId solidFillId = GetDraftingSolidFillPatternId(doc);

            foreach (var er in result.ElementResults)
            {
                var revitElem = doc.GetElement(new ElementId(er.ElementId));
                if (revitElem == null) continue;

                var solids = SolidUnionProcessor.GetSolids(revitElem);
                if (solids.Count == 0) continue;
                var unioned = SolidUnionProcessor.Union(solids);
                var finalSolids = unioned != null ? new List<Solid> { unioned } : solids;
                double minZ = FaceClassifier.GetMinZ(finalSolids);
                var faces = FaceClassifier.ClassifyAll(finalSolids, glFeet, minZ);

                string key = GetKey(er, settings);
                if (!keyAssignment.TryGetValue(key, out var color))
                    color = (180, 180, 180);

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

                        // 共有パラメータに識別用値を書き込み
                        try
                        {
                            double areaM2 = UnitUtils.ConvertFromInternalUnits(fi.Area, UnitTypeId.SquareMeters);
                            FormworkParameterManager.SetInstanceValues(ds, catLabel, key, areaM2);
                        }
                        catch { }

                        dsId = ds.Id;
                    }
                    catch { continue; }

                    if (dsId == null) continue;

                    // 色オーバーライド
                    ApplyColorOverride(view, dsId, color, fi.FaceType, solidFillId);

                    vr.CreatedShapeIds.Add(dsId);
                }
            }

            // 元要素を半透明に
            try
            {
                var baseOgs = new OverrideGraphicSettings();
                baseOgs.SetSurfaceTransparency(70);
                foreach (var er in result.ElementResults)
                {
                    try
                    {
                        var id = new ElementId(er.ElementId);
                        if (doc.GetElement(id) != null)
                            view.SetElementOverrides(id, baseOgs);
                    }
                    catch { }
                }
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
                // Schedule / Legend は HideElements 不可
                if (v is ViewSchedule) continue;
                if (v.ViewType == ViewType.Legend) continue;

                try
                {
                    v.HideElements(idsList);
                }
                catch
                {
                    // ビューで隠せない要素がある場合などは無視
                }
            }
        }

        private static void ApplyColorOverride(
            View view, ElementId targetId,
            (byte R, byte G, byte B) color, FaceType ftype, ElementId solidFillId)
        {
            byte r, g, b;
            if (ftype == FaceType.Error)
            {
                r = 220; g = 30; b = 30;
            }
            else if (IsDeducted(ftype))
            {
                var d = DeductedColor(ftype);
                r = d.R; g = d.G; b = d.B;
            }
            else
            {
                r = color.R; g = color.G; b = color.B;
            }

            var ogs = new OverrideGraphicSettings();
            var revitColor = new Color(r, g, b);

            ogs.SetProjectionLineColor(revitColor);
            ogs.SetSurfaceForegroundPatternColor(revitColor);
            if (solidFillId != null && solidFillId != ElementId.InvalidElementId)
                ogs.SetSurfaceForegroundPatternId(solidFillId);
            ogs.SetSurfaceForegroundPatternVisible(true);
            ogs.SetSurfaceBackgroundPatternVisible(false);

            try { view.SetElementOverrides(targetId, ogs); } catch { }
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

        private static (byte R, byte G, byte B) DeductedColor(FaceType t)
        {
            switch (t)
            {
                case FaceType.DeductedTop: return (180, 230, 180);
                case FaceType.DeductedBottom: return (200, 170, 130);
                case FaceType.DeductedContact: return (200, 200, 200);
                case FaceType.DeductedBelowGL: return (160, 140, 120);
                default: return (220, 220, 220);
            }
        }

        private static string GetKey(ElementResult er, FormworkSettings s)
        {
            switch (s.ColorScheme)
            {
                case ColorSchemeType.ByZone: return string.IsNullOrEmpty(er.Zone) ? "未設定" : er.Zone;
                case ColorSchemeType.ByFormworkType: return string.IsNullOrEmpty(er.FormworkType) ? "未設定" : er.FormworkType;
                default: return "C:" + er.Category;
            }
        }

        private static Dictionary<string, (byte R, byte G, byte B)> AssignColors(
            FormworkResult result, FormworkSettings s)
        {
            var map = new Dictionary<string, (byte, byte, byte)>();
            if (s.ColorScheme == ColorSchemeType.ByCategory)
            {
                foreach (var kv in _categoryColors)
                    map["C:" + kv.Key] = kv.Value;
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
        /// Drafting（製図）ソリッド塗り潰しパターンの Id を取得。サーフェスオーバーライド用。
        /// </summary>
        private static ElementId GetDraftingSolidFillPatternId(Document doc)
        {
            var fps = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>();
            foreach (var fp in fps)
            {
                var pattern = fp.GetFillPattern();
                if (pattern == null) continue;
                if (pattern.IsSolidFill && pattern.Target == FillPatternTarget.Drafting)
                    return fp.Id;
            }
            // フォールバック: Model の solid fill を探す
            foreach (var fp in fps)
            {
                var pattern = fp.GetFillPattern();
                if (pattern == null) continue;
                if (pattern.IsSolidFill)
                    return fp.Id;
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

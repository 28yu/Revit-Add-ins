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
    /// </summary>
    internal static class FormworkVisualizer
    {
        private static readonly Dictionary<CategoryGroup, (byte R, byte G, byte B)> _categoryColors
            = new Dictionary<CategoryGroup, (byte, byte, byte)>
            {
                { CategoryGroup.Column,     (60, 120, 240) },   // 青
                { CategoryGroup.Beam,       (60, 200, 100) },   // 緑
                { CategoryGroup.Wall,       (255, 160, 60) },   // オレンジ
                { CategoryGroup.Slab,       (250, 220, 60) },   // 黄
                { CategoryGroup.Foundation, (160, 100, 220) },  // 紫
                { CategoryGroup.Stairs,     (60, 200, 220) },   // シアン
                { CategoryGroup.Other,      (180, 180, 180) },  // グレー
            };

        private static readonly (byte R, byte G, byte B)[] _autoPalette = new (byte, byte, byte)[]
        {
            (230, 80, 80),   (80, 160, 230), (60, 200, 100), (240, 200, 60),
            (180, 100, 220), (255, 130, 170), (80, 210, 210), (255, 160, 60),
            (170, 220, 80), (60, 170, 170), (220, 100, 200), (120, 120, 220),
        };

        internal static View3D CreateVisualization(
            Document doc,
            FormworkResult result,
            FormworkSettings settings)
        {
            var view3DType = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(v => v.ViewFamily == ViewFamily.ThreeDimensional);
            if (view3DType == null) return null;

            var view = View3D.CreateIsometric(doc, view3DType.Id);
            string viewName = $"型枠分析_{DateTime.Now:yyyyMMdd_HHmmss}";
            try { view.Name = viewName; } catch { }

            // 色 → Material Id キャッシュ
            var colorKeyToMat = new Dictionary<string, ElementId>();
            // カラースキームごとのキー割当
            var keyAssignment = AssignColors(result, settings);

            var catOst = new ElementId(BuiltInCategory.OST_GenericModel);

            // GL 高さ（m → feet）
            double? glFeet = null;
            if (settings.UseGLDeduction)
            {
                glFeet = UnitUtils.ConvertToInternalUnits(settings.GLElevationMeters, UnitTypeId.Meters);
            }

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

                ElementId matId = GetOrCreateMaterial(doc, colorKeyToMat, key, color);

                foreach (var fi in faces)
                {
                    if (fi.FaceType != FaceType.FormworkRequired &&
                        !(settings.ShowDeductedFaces && IsDeducted(fi.FaceType)))
                        continue;

                    Solid thin = CreateThinSolidFromFace(fi);
                    if (thin == null) continue;

                    try
                    {
                        var ds = DirectShape.CreateElement(doc, catOst);
                        ds.ApplicationId = "Tools28";
                        ds.ApplicationDataId = "Formwork";
                        ds.SetShape(new GeometryObject[] { thin });

                        if (matId != null && matId != ElementId.InvalidElementId)
                        {
                            var ogs = new OverrideGraphicSettings();
                            byte r, g, b;
                            if (fi.FaceType == FaceType.Error)
                            {
                                r = 220; g = 30; b = 30;
                            }
                            else if (IsDeducted(fi.FaceType))
                            {
                                var d = DeductedColor(fi.FaceType);
                                r = d.R; g = d.G; b = d.B;
                            }
                            else
                            {
                                r = color.R; g = color.G; b = color.B;
                            }
                            ogs.SetSurfaceForegroundPatternColor(new Color(r, g, b));
                            ogs.SetProjectionLineColor(new Color(r, g, b));
                            try
                            {
                                var solidFillId = GetSolidFillPatternId(doc);
                                if (solidFillId != null)
                                    ogs.SetSurfaceForegroundPatternId(solidFillId);
                            }
                            catch { }
                            try { view.SetElementOverrides(ds.Id, ogs); } catch { }
                        }
                    }
                    catch { }
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

            return view;
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
                default: return CategoryKey(er.Category);
            }
        }

        private static string CategoryKey(CategoryGroup cg) => "C:" + cg;

        private static Dictionary<string, (byte R, byte G, byte B)> AssignColors(
            FormworkResult result, FormworkSettings s)
        {
            var map = new Dictionary<string, (byte, byte, byte)>();
            if (s.ColorScheme == ColorSchemeType.ByCategory)
            {
                foreach (var kv in _categoryColors)
                    map[CategoryKey(kv.Key)] = kv.Value;
                return map;
            }

            var keys = new HashSet<string>();
            foreach (var er in result.ElementResults)
            {
                keys.Add(GetKey(er, s));
            }

            int i = 0;
            foreach (var k in keys.OrderBy(x => x))
            {
                map[k] = _autoPalette[i % _autoPalette.Length];
                i++;
            }
            return map;
        }

        private static ElementId GetOrCreateMaterial(
            Document doc,
            Dictionary<string, ElementId> cache,
            string key,
            (byte R, byte G, byte B) color)
        {
            string name = $"Formwork_{key}_{color.R}_{color.G}_{color.B}";
            if (cache.TryGetValue(name, out var id)) return id;

            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(Material))
                .Cast<Material>()
                .FirstOrDefault(m => m.Name == name);
            if (existing != null)
            {
                cache[name] = existing.Id;
                return existing.Id;
            }

            try
            {
                var mid = Material.Create(doc, name);
                var m = doc.GetElement(mid) as Material;
                if (m != null)
                {
                    m.Color = new Color(color.R, color.G, color.B);
                    m.Transparency = 0;
                }
                cache[name] = mid;
                return mid;
            }
            catch
            {
                return ElementId.InvalidElementId;
            }
        }

        private static ElementId GetSolidFillPatternId(Document doc)
        {
            var fps = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>();
            foreach (var fp in fps)
            {
                if (fp.GetFillPattern().IsSolidFill &&
                    fp.GetFillPattern().Target == FillPatternTarget.Drafting)
                    return fp.Id;
            }
            foreach (var fp in fps)
            {
                if (fp.GetFillPattern().IsSolidFill)
                    return fp.Id;
            }
            return null;
        }

        /// <summary>
        /// Face から法線方向に薄い Solid を生成（可視化用）
        /// </summary>
        private static Solid CreateThinSolidFromFace(FaceClassifier.FaceInfo fi)
        {
            try
            {
                var edges = fi.Face.GetEdgesAsCurveLoops();
                if (edges == null || edges.Count == 0) return null;

                XYZ n = fi.Normal?.Normalize();
                if (n == null) return null;

                double thickness = 0.03;  // 約 10mm
                var extr = GeometryCreationUtils.CreateExtrusion(edges, n, thickness);
                return extr;
            }
            catch
            {
                return null;
            }
        }
    }

    internal static class GeometryCreationUtils
    {
        internal static Solid CreateExtrusion(IList<CurveLoop> loops, XYZ direction, double distance)
        {
            try
            {
                return GeometryCreationUtilities.CreateExtrusionGeometry(loops, direction, distance);
            }
            catch
            {
                try
                {
                    return GeometryCreationUtilities.CreateExtrusionGeometry(loops, direction.Negate(), distance);
                }
                catch
                {
                    return null;
                }
            }
        }
    }
}

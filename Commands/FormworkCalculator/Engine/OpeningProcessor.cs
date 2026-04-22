using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 開口部（ドア/窓/壁開口/床開口）の見込み面を加算する。
    /// 【計算式】開口部による型枠面積調整 = -(開口面積 × 2) + (開口周長 × 躯体厚)
    ///   壁は両面分控除のため × 2。床は片面（スラブ下面）× 1。
    /// </summary>
    internal static class OpeningProcessor
    {
        internal class OpeningDelta
        {
            public int HostElementId { get; set; }
            public double DeductedArea { get; set; }  // 控除する開口面積（両面合計）
            public double AddedEdgeArea { get; set; } // 加算する見込み面
        }

        internal static List<OpeningDelta> Compute(Document doc, IEnumerable<Element> hostElements)
        {
            var list = new List<OpeningDelta>();
            foreach (var host in hostElements)
            {
                if (host == null) continue;
                try
                {
                    if (host is Wall w)
                        list.Add(ComputeForWall(doc, w));
                    else if (host is Floor f)
                        list.Add(ComputeForFloor(doc, f));
                }
                catch
                {
                    // 個別要素で失敗しても全体処理は継続
                }
            }
            return list;
        }

        private static OpeningDelta ComputeForWall(Document doc, Wall wall)
        {
            double thickness = 0;
            try { thickness = wall.Width; } catch { }

            double deducted = 0;
            double added = 0;

            var insertIds = wall.FindInserts(true, true, true, true);
            foreach (ElementId id in insertIds)
            {
                Element insert = doc.GetElement(id);
                if (insert == null) continue;

                double w, h;
                if (!TryGetRectSize(insert, out w, out h)) continue;

                double area = w * h;
                double perim = 2 * (w + h);

                // 壁は両面分控除
                deducted += area * 2.0;
                added += perim * thickness;
            }

            return new OpeningDelta
            {
                HostElementId = wall.Id.IntegerValue,
                DeductedArea = deducted,
                AddedEdgeArea = added,
            };
        }

        private static OpeningDelta ComputeForFloor(Document doc, Floor floor)
        {
            double thickness = 0;
            try
            {
                var ft = doc.GetElement(floor.GetTypeId()) as FloorType;
                if (ft != null)
                {
                    var p = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                    if (p != null && p.HasValue) thickness = p.AsDouble();
                }
            }
            catch { }

            double deducted = 0;
            double added = 0;

            // 床ホスト内の開口
            var openings = new FilteredElementCollector(doc)
                .OfClass(typeof(Opening))
                .Cast<Opening>()
                .Where(o => o.Host != null && o.Host.Id == floor.Id);

            foreach (var op in openings)
            {
                double area, perim;
                if (!TryGetOpeningAreaPerimeter(op, out area, out perim)) continue;

                // 床は下面片側分控除 + スラブ厚見込み面
                deducted += area;
                added += perim * thickness;
            }

            return new OpeningDelta
            {
                HostElementId = floor.Id.IntegerValue,
                DeductedArea = deducted,
                AddedEdgeArea = added,
            };
        }

        private static bool TryGetRectSize(Element insert, out double width, out double height)
        {
            width = 0;
            height = 0;
            try
            {
                Parameter wp = insert.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM);
                Parameter hp = insert.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM);

                if (wp != null && wp.HasValue) width = wp.AsDouble();
                if (hp != null && hp.HasValue) height = hp.AsDouble();

                if (width > 0 && height > 0) return true;

                if (insert is FamilyInstance fi)
                {
                    var sym = fi.Symbol;
                    if (sym != null)
                    {
                        var tw = sym.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM);
                        var th = sym.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM);
                        if (tw != null && tw.HasValue) width = tw.AsDouble();
                        if (th != null && th.HasValue) height = th.AsDouble();
                    }
                }

                // 最終フォールバック: BoundingBox
                if (width <= 0 || height <= 0)
                {
                    var bb = insert.get_BoundingBox(null);
                    if (bb != null)
                    {
                        double dx = bb.Max.X - bb.Min.X;
                        double dy = bb.Max.Y - bb.Min.Y;
                        double dz = bb.Max.Z - bb.Min.Z;
                        if (width <= 0) width = Math.Max(dx, dy);
                        if (height <= 0) height = dz;
                    }
                }

                return width > 0 && height > 0;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetOpeningAreaPerimeter(Opening op, out double area, out double perimeter)
        {
            area = 0;
            perimeter = 0;

            CurveArray boundary = null;
            try { boundary = op.BoundaryCurves; } catch { }
            if (boundary == null || boundary.Size == 0) return false;

            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;

            foreach (Curve c in boundary)
            {
                perimeter += c.Length;
                var p0 = c.GetEndPoint(0);
                var p1 = c.GetEndPoint(1);
                minX = Math.Min(minX, Math.Min(p0.X, p1.X));
                minY = Math.Min(minY, Math.Min(p0.Y, p1.Y));
                maxX = Math.Max(maxX, Math.Max(p0.X, p1.X));
                maxY = Math.Max(maxY, Math.Max(p0.Y, p1.Y));
            }

            // 近似: BBox 面積で代用（多角形でも安全側）
            area = (maxX - minX) * (maxY - minY);
            return area > 0 && perimeter > 0;
        }
    }
}

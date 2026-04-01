using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FireProtection
{
    /// <summary>
    /// 梁の平面輪郭取得とオフセット・統合処理
    /// </summary>
    public static class BeamGeometryHelper
    {
        /// <summary>
        /// 梁の平面投影輪郭をオフセットした矩形CurveLoopを取得
        /// </summary>
        public static CurveLoop GetBeamOffsetOutline(FamilyInstance beam, double offsetFeet)
        {
            LocationCurve locCurve = beam.Location as LocationCurve;
            if (locCurve == null) return null;

            Curve curve = locCurve.Curve;
            Line line = curve as Line;
            if (line == null) return null;

            XYZ start = new XYZ(line.GetEndPoint(0).X, line.GetEndPoint(0).Y, 0);
            XYZ end = new XYZ(line.GetEndPoint(1).X, line.GetEndPoint(1).Y, 0);

            if (start.DistanceTo(end) < 0.001)
                return null;

            XYZ direction = (end - start).Normalize();
            XYZ perp = new XYZ(-direction.Y, direction.X, 0);

            double beamWidth = GetBeamWidth(beam);
            double halfExtent = beamWidth / 2.0 + offsetFeet;

            XYZ extStart = start - direction * offsetFeet;
            XYZ extEnd = end + direction * offsetFeet;

            XYZ p0 = extStart - perp * halfExtent;
            XYZ p1 = extEnd - perp * halfExtent;
            XYZ p2 = extEnd + perp * halfExtent;
            XYZ p3 = extStart + perp * halfExtent;

            CurveLoop loop = new CurveLoop();
            loop.Append(Line.CreateBound(p0, p1));
            loop.Append(Line.CreateBound(p1, p2));
            loop.Append(Line.CreateBound(p2, p3));
            loop.Append(Line.CreateBound(p3, p0));

            return loop;
        }

        /// <summary>
        /// 複数の輪郭をソリッドブーリアン演算で統合
        /// </summary>
        public static List<CurveLoop> MergeOutlines(List<CurveLoop> outlines)
        {
            if (outlines.Count <= 1)
                return outlines;

            Solid result = null;
            var failedLoops = new List<CurveLoop>();

            foreach (var outline in outlines)
            {
                try
                {
                    Solid extrusion = GeometryCreationUtilities.CreateExtrusionGeometry(
                        new List<CurveLoop> { outline }, XYZ.BasisZ, 1.0);

                    if (result == null)
                    {
                        result = extrusion;
                    }
                    else
                    {
                        try
                        {
                            result = BooleanOperationsUtils.ExecuteBooleanOperation(
                                result, extrusion, BooleanOperationsType.Union);
                        }
                        catch
                        {
                            failedLoops.Add(outline);
                        }
                    }
                }
                catch
                {
                    failedLoops.Add(outline);
                }
            }

            var mergedLoops = new List<CurveLoop>();

            if (result != null)
            {
                foreach (Face face in result.Faces)
                {
                    PlanarFace pf = face as PlanarFace;
                    if (pf != null && Math.Abs(pf.FaceNormal.Z - 1.0) < 0.01)
                    {
                        foreach (var edgeLoop in pf.GetEdgesAsCurveLoops())
                        {
                            var flatLoop = new CurveLoop();
                            bool valid = true;

                            foreach (Curve c in edgeLoop)
                            {
                                XYZ s = c.GetEndPoint(0);
                                XYZ e = c.GetEndPoint(1);
                                XYZ fs = new XYZ(s.X, s.Y, 0);
                                XYZ fe = new XYZ(e.X, e.Y, 0);

                                if (fs.DistanceTo(fe) < 0.0001)
                                    continue;

                                try
                                {
                                    flatLoop.Append(Line.CreateBound(fs, fe));
                                }
                                catch
                                {
                                    valid = false;
                                    break;
                                }
                            }

                            if (valid && flatLoop.Count() > 0)
                                mergedLoops.Add(flatLoop);
                        }
                    }
                }
            }

            mergedLoops.AddRange(failedLoops);
            return mergedLoops.Count > 0 ? mergedLoops : outlines;
        }

        /// <summary>
        /// 梁幅を取得（パラメータ探索→BoundingBox→デフォルト値）
        /// </summary>
        public static double GetBeamWidth(FamilyInstance beam)
        {
            var paramNames = new[] { "b", "B", "幅", "梁幅", "W", "w", "Width", "width" };

            foreach (string name in paramNames)
            {
                Parameter p = beam.LookupParameter(name);
                if (p != null && p.StorageType == StorageType.Double)
                {
                    double val = p.AsDouble();
                    if (val > 0) return val;
                }
            }

            ElementType beamType = beam.Document.GetElement(beam.GetTypeId()) as ElementType;
            if (beamType != null)
            {
                foreach (string name in paramNames)
                {
                    Parameter p = beamType.LookupParameter(name);
                    if (p != null && p.StorageType == StorageType.Double)
                    {
                        double val = p.AsDouble();
                        if (val > 0) return val;
                    }
                }
            }

            BoundingBoxXYZ bb = beam.get_BoundingBox(null);
            if (bb != null)
            {
                LocationCurve lc = beam.Location as LocationCurve;
                if (lc != null && lc.Curve is Line line)
                {
                    XYZ dir = line.Direction.Normalize();
                    XYZ perp = new XYZ(-dir.Y, dir.X, 0).Normalize();
                    double width = Math.Abs((bb.Max - bb.Min).DotProduct(perp));
                    if (width > 0.01) return width;
                }
            }

            return 200.0 / 304.8;
        }
    }
}

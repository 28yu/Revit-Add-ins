using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// Solid の Face を走査して型枠要否を判定する。
    /// </summary>
    internal static class FaceClassifier
    {
        private const double VerticalZTol = 0.99;      // 水平面判定: |Nz| > 0.99
        private const double HorizontalZTol = 0.01;    // 垂直面判定: |Nz| < 0.01

        internal class FaceInfo
        {
            public Face Face;
            public XYZ Normal;       // 代表法線（面中央）
            public double Area;
            public FaceType FaceType;
        }

        /// <summary>
        /// 結合 Solid 群から全 Face を取り出して分類する。
        /// </summary>
        internal static List<FaceInfo> ClassifyAll(
            IEnumerable<Solid> unionedSolids,
            double? glElevationFeet,
            double minBottomZ)
        {
            var list = new List<FaceInfo>();
            foreach (var s in unionedSolids)
            {
                if (s == null) continue;
                foreach (Face f in s.Faces)
                {
                    var info = Classify(f, glElevationFeet, minBottomZ);
                    if (info != null) list.Add(info);
                }
            }
            return list;
        }

        internal static FaceInfo Classify(Face f, double? glElevationFeet, double minBottomZ)
        {
            if (f == null) return null;
            XYZ n = GetFaceNormal(f);
            if (n == null) return null;

            double area = 0;
            try { area = f.Area; } catch { }
            if (area <= 1e-6) return null;

            FaceType type;
            double nz = n.Z;

            if (nz > VerticalZTol)
            {
                type = FaceType.DeductedTop;
            }
            else if (nz < -VerticalZTol)
            {
                // 下向き水平面: 最下部（基礎底面）か判定
                BoundingBoxUV bb = f.GetBoundingBox();
                UV mid = (bb.Min + bb.Max) * 0.5;
                XYZ pt = null;
                try { pt = f.Evaluate(mid); } catch { }

                if (pt != null && Math.Abs(pt.Z - minBottomZ) < 0.01)
                {
                    type = FaceType.DeductedBottom;
                }
                else if (glElevationFeet.HasValue && pt != null && pt.Z < glElevationFeet.Value - 1e-3)
                {
                    type = FaceType.DeductedBelowGL;
                }
                else
                {
                    type = FaceType.FormworkRequired;
                }
            }
            else if (Math.Abs(nz) < HorizontalZTol)
            {
                // 垂直面
                if (glElevationFeet.HasValue)
                {
                    BoundingBoxUV bb = f.GetBoundingBox();
                    UV mid = (bb.Min + bb.Max) * 0.5;
                    XYZ pt = null;
                    try { pt = f.Evaluate(mid); } catch { }
                    if (pt != null && pt.Z < glElevationFeet.Value - 1e-3)
                    {
                        type = FaceType.DeductedBelowGL;
                    }
                    else
                    {
                        type = FaceType.FormworkRequired;
                    }
                }
                else
                {
                    type = FaceType.FormworkRequired;
                }
            }
            else
            {
                // 傾斜面
                type = FaceType.Inclined;
            }

            return new FaceInfo
            {
                Face = f,
                Normal = n,
                Area = area,
                FaceType = type,
            };
        }

        internal static XYZ GetFaceNormal(Face f)
        {
            if (f == null) return null;
            try
            {
                BoundingBoxUV bb = f.GetBoundingBox();
                UV mid = (bb.Min + bb.Max) * 0.5;
                XYZ n = f.ComputeNormal(mid);
                return n?.Normalize();
            }
            catch
            {
                return null;
            }
        }

        internal static double GetMinZ(IEnumerable<Solid> solids)
        {
            double minZ = double.MaxValue;
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
                        if (p.Z < minZ) minZ = p.Z;
                    }
                }
            }
            return minZ == double.MaxValue ? 0 : minZ;
        }
    }
}

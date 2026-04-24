using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// Solid の Face を走査して型枠要否を判定する。
    /// 【ルール】
    ///   - 上向き水平面 (nz > 0.99) かつ最上面 → DeductedTop
    ///   - 下向き水平面 (nz < -0.99) かつ最下面 → DeductedBottom
    ///   - その他の上向き面（リビールの底など）→ FormworkRequired
    ///   - その他の下向き面（リビールの天井など）→ FormworkRequired
    ///   - 垂直面 → FormworkRequired
    ///   - 傾斜面 → FormworkRequired (リビール/スイープの斜面など)
    ///   - GL より下の垂直/下向き面 → DeductedBelowGL (オプション)
    /// </summary>
    internal static class FaceClassifier
    {
        private const double VerticalZTol = 0.99;      // |Nz| > 0.99 → 水平
        private const double TopBottomTol = 0.01;      // 面の高さがmin/max±この距離以内なら最上/最下面

        internal class FaceInfo
        {
            public Face Face;
            public XYZ Normal;
            public double Area;
            public FaceType FaceType;

            /// <summary>
            /// 部分接触の記録。FaceType が FormworkRequired のまま残る面について、
            /// 他要素の面と部分的に接触している領域を追跡する。
            /// BuildElementResult で面積を控除するのに使う。
            /// Phase 2 で視覚化を厳密化する際にも使う。
            /// </summary>
            public List<PartialContact> PartialContacts = new List<PartialContact>();
        }

        /// <summary>
        /// 部分接触情報。面A (FormworkRequired) の一部が他要素の面B (小面積) に接触している。
        /// </summary>
        internal class PartialContact
        {
            public int OtherElementId;      // 接触相手の要素ID
            public int OtherFaceIndex;      // 相手の面インデックス (debug用)
            public double ContactArea;      // 控除対象の面積 (feet²)
            public BoundingBoxUV UvBounds;  // B面におけるB自身のUV範囲 (debug用)

            /// <summary>
            /// 面A の UV 空間上で、相手面B の投影が占める矩形。
            /// Phase 2 の DirectShape 厳密化で使用する。
            /// null の場合は A 上への投影に失敗したことを示す → Phase 2 はフォールバック。
            /// </summary>
            public BoundingBoxUV UvBoundsOnA;
        }

        internal static List<FaceInfo> ClassifyAll(
            IEnumerable<Solid> unionedSolids,
            double? glElevationFeet,
            double minBottomZ,
            double maxTopZ)
        {
            var list = new List<FaceInfo>();
            foreach (var s in unionedSolids)
            {
                if (s == null) continue;
                foreach (Face f in s.Faces)
                {
                    var info = Classify(f, glElevationFeet, minBottomZ, maxTopZ);
                    if (info != null) list.Add(info);
                }
            }
            return list;
        }

        internal static FaceInfo Classify(
            Face f, double? glElevationFeet, double minBottomZ, double maxTopZ)
        {
            if (f == null) return null;
            XYZ n = GetFaceNormal(f);
            if (n == null) return null;

            double area = 0;
            try { area = f.Area; } catch { }
            if (area <= 1e-6) return null;

            FaceType type;
            double nz = n.Z;

            BoundingBoxUV bb = null;
            try { bb = f.GetBoundingBox(); } catch { }
            UV mid = bb != null ? (bb.Min + bb.Max) * 0.5 : new UV(0, 0);
            XYZ pt = null;
            try { pt = f.Evaluate(mid); } catch { }

            if (nz > VerticalZTol)
            {
                // 上向き水平面: 最上面のみ DeductedTop。リビール底などは FormworkRequired
                if (pt != null && Math.Abs(pt.Z - maxTopZ) < TopBottomTol)
                    type = FaceType.DeductedTop;
                else
                    type = FaceType.FormworkRequired;
            }
            else if (nz < -VerticalZTol)
            {
                // 下向き水平面: 最下面のみ DeductedBottom。リビール天井などは FormworkRequired
                if (pt != null && Math.Abs(pt.Z - minBottomZ) < TopBottomTol)
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
            else
            {
                // 垂直面 or 傾斜面 → 全て FormworkRequired (リビール/スイープの斜面も含む)
                if (glElevationFeet.HasValue && pt != null && pt.Z < glElevationFeet.Value - 1e-3)
                {
                    type = FaceType.DeductedBelowGL;
                }
                else
                {
                    type = FaceType.FormworkRequired;
                }
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

        internal static (double minZ, double maxZ) GetZRange(IEnumerable<Solid> solids)
        {
            double minZ = double.MaxValue;
            double maxZ = double.MinValue;
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
                        if (p.Z > maxZ) maxZ = p.Z;
                    }
                }
            }
            if (minZ == double.MaxValue) minZ = 0;
            if (maxZ == double.MinValue) maxZ = 0;
            return (minZ, maxZ);
        }

        internal static double GetMinZ(IEnumerable<Solid> solids)
        {
            return GetZRange(solids).minZ;
        }
    }
}

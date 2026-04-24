using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 部分接触がある面について、接触領域を差し引いた残りの矩形から
    /// 薄板 Solid を構築する (Phase 2 視覚化厳密化)。
    ///
    /// 【前提】面 A が軸平行な矩形 (UV空間で 4 角形) である。
    /// 【方式】矩形ベースの 2D 差分: A の矩形から各 PartialContact の投影矩形を
    ///   順次減算し、最大 4 つのサブ矩形に分解する。各サブ矩形から薄板 Solid を作る。
    ///
    /// 【フォールバック】
    ///   - A が矩形でない (開口付き壁、非矩形の面) → Success=false、Visualizer が従来処理
    ///   - UvBoundsOnA が null (投影失敗) → 失敗
    ///   - Solid 構築に失敗 → 失敗
    /// </summary>
    internal static class PartialContactClipper
    {
        private const double UvEps = 1e-6;
        private const double ThicknessFeet = 0.03;  // 約 10mm

        internal class ClipResult
        {
            public List<Solid> Solids = new List<Solid>();
            public bool Success;
            public string FailReason;
        }

        private struct RectUV
        {
            public double UMin, VMin, UMax, VMax;
            public bool IsDegenerate(double tol)
            {
                return (UMax - UMin) <= tol || (VMax - VMin) <= tol;
            }
        }

        internal static ClipResult TryClip(FaceClassifier.FaceInfo fi)
        {
            var result = new ClipResult();
            if (fi?.Face == null) { result.FailReason = "face-null"; return result; }
            if (fi.PartialContacts == null || fi.PartialContacts.Count == 0)
            { result.FailReason = "no-partial-contacts"; return result; }

            // 1. 面の CurveLoop 取得と矩形性判定
            IList<CurveLoop> loops;
            try { loops = fi.Face.GetEdgesAsCurveLoops(); }
            catch { result.FailReason = "get-loops-failed"; return result; }
            if (loops == null || loops.Count != 1)
            { result.FailReason = "multiple-or-no-loops"; return result; }

            var rectMain = TryExtractRectUV(fi.Face, loops[0]);
            if (rectMain == null)
            { result.FailReason = "not-rectangular"; return result; }

            // 全 PartialContact の UvBoundsOnA が揃っているか確認
            foreach (var pc in fi.PartialContacts)
            {
                if (pc.UvBoundsOnA == null)
                { result.FailReason = "no-uv-on-a"; return result; }
            }

            double uDim = rectMain.Value.UMax - rectMain.Value.UMin;
            double vDim = rectMain.Value.VMax - rectMain.Value.VMin;
            double degenTol = Math.Max(uDim, vDim) * 1e-4;

            // 2. 順次差分
            var pieces = new List<RectUV> { rectMain.Value };
            foreach (var pc in fi.PartialContacts)
            {
                var bRect = new RectUV
                {
                    UMin = pc.UvBoundsOnA.Min.U,
                    VMin = pc.UvBoundsOnA.Min.V,
                    UMax = pc.UvBoundsOnA.Max.U,
                    VMax = pc.UvBoundsOnA.Max.V,
                };

                var next = new List<RectUV>();
                foreach (var p in pieces)
                {
                    foreach (var sub in Subtract(p, bRect))
                    {
                        if (!sub.IsDegenerate(degenTol)) next.Add(sub);
                    }
                }
                pieces = next;
                if (pieces.Count == 0) break;
            }

            if (pieces.Count == 0)
            { result.FailReason = "fully-clipped"; return result; }

            // 3. 各サブ矩形から薄板 Solid を作る
            XYZ normal = fi.Normal?.Normalize();
            if (normal == null) { result.FailReason = "no-normal"; return result; }

            foreach (var piece in pieces)
            {
                CurveLoop loop = BuildCurveLoopFromRect(fi.Face, piece);
                if (loop == null) continue;

                Solid solid = null;
                try
                {
                    solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                        new[] { loop }, normal, ThicknessFeet);
                }
                catch
                {
                    try
                    {
                        solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                            new[] { loop }, normal.Negate(), ThicknessFeet);
                    }
                    catch { }
                }
                if (solid != null) result.Solids.Add(solid);
            }

            if (result.Solids.Count == 0)
            { result.FailReason = "solid-build-failed"; return result; }

            result.Success = true;
            return result;
        }

        /// <summary>
        /// CurveLoop が軸平行矩形なら UV 範囲を返す。非矩形なら null。
        /// </summary>
        private static RectUV? TryExtractRectUV(Face face, CurveLoop loop)
        {
            if (loop == null) return null;

            var pts = new List<UV>();
            foreach (Curve c in loop)
            {
                if (!(c is Line)) return null;
                var p0 = c.GetEndPoint(0);
                IntersectionResult proj = null;
                try { proj = face.Project(p0); } catch { }
                if (proj?.UVPoint == null) return null;
                pts.Add(proj.UVPoint);
            }
            if (pts.Count != 4) return null;

            double uMin = double.MaxValue, uMax = double.MinValue;
            double vMin = double.MaxValue, vMax = double.MinValue;
            foreach (var uv in pts)
            {
                if (uv.U < uMin) uMin = uv.U;
                if (uv.U > uMax) uMax = uv.U;
                if (uv.V < vMin) vMin = uv.V;
                if (uv.V > vMax) vMax = uv.V;
            }
            double uDim = uMax - uMin;
            double vDim = vMax - vMin;
            if (uDim <= UvEps || vDim <= UvEps) return null;

            // 4 頂点が矩形の 4 隅にあるか確認
            double tolU = uDim * 1e-3;
            double tolV = vDim * 1e-3;
            int cornerHits = 0;
            foreach (var uv in pts)
            {
                bool atUMin = Math.Abs(uv.U - uMin) < tolU;
                bool atUMax = Math.Abs(uv.U - uMax) < tolU;
                bool atVMin = Math.Abs(uv.V - vMin) < tolV;
                bool atVMax = Math.Abs(uv.V - vMax) < tolV;
                if ((atUMin || atUMax) && (atVMin || atVMax)) cornerHits++;
            }
            if (cornerHits != 4) return null;

            return new RectUV { UMin = uMin, VMin = vMin, UMax = uMax, VMax = vMax };
        }

        /// <summary>
        /// A - B を最大 4 つのサブ矩形に分解。重なりなしなら [A]、完全包含なら []。
        /// </summary>
        private static List<RectUV> Subtract(RectUV a, RectUV b)
        {
            var result = new List<RectUV>();

            double uMinX = Math.Max(a.UMin, b.UMin);
            double uMaxX = Math.Min(a.UMax, b.UMax);
            double vMinX = Math.Max(a.VMin, b.VMin);
            double vMaxX = Math.Min(a.VMax, b.VMax);

            if (uMaxX <= uMinX + UvEps || vMaxX <= vMinX + UvEps)
            {
                result.Add(a);
                return result;
            }

            bool uCoversAll = uMinX <= a.UMin + UvEps && uMaxX >= a.UMax - UvEps;
            bool vCoversAll = vMinX <= a.VMin + UvEps && vMaxX >= a.VMax - UvEps;
            if (uCoversAll && vCoversAll) return result;

            // 左側
            if (a.UMin < uMinX - UvEps)
                result.Add(new RectUV { UMin = a.UMin, VMin = a.VMin, UMax = uMinX, VMax = a.VMax });
            // 右側
            if (a.UMax > uMaxX + UvEps)
                result.Add(new RectUV { UMin = uMaxX, VMin = a.VMin, UMax = a.UMax, VMax = a.VMax });
            // 下側 (u 方向は交差部のみ)
            if (a.VMin < vMinX - UvEps)
                result.Add(new RectUV { UMin = uMinX, VMin = a.VMin, UMax = uMaxX, VMax = vMinX });
            // 上側
            if (a.VMax > vMaxX + UvEps)
                result.Add(new RectUV { UMin = uMinX, VMin = vMaxX, UMax = uMaxX, VMax = a.VMax });

            return result;
        }

        /// <summary>
        /// UV 矩形の 4 隅を Face.Evaluate で 3D 点に変換し、CurveLoop を構築する。
        /// </summary>
        private static CurveLoop BuildCurveLoopFromRect(Face face, RectUV r)
        {
            XYZ p0, p1, p2, p3;
            try
            {
                p0 = face.Evaluate(new UV(r.UMin, r.VMin));
                p1 = face.Evaluate(new UV(r.UMax, r.VMin));
                p2 = face.Evaluate(new UV(r.UMax, r.VMax));
                p3 = face.Evaluate(new UV(r.UMin, r.VMax));
            }
            catch { return null; }
            if (p0 == null || p1 == null || p2 == null || p3 == null) return null;

            const double MinEdge = 1e-4;
            if (p0.DistanceTo(p1) < MinEdge) return null;
            if (p1.DistanceTo(p2) < MinEdge) return null;
            if (p2.DistanceTo(p3) < MinEdge) return null;
            if (p3.DistanceTo(p0) < MinEdge) return null;

            try
            {
                var loop = new CurveLoop();
                loop.Append(Line.CreateBound(p0, p1));
                loop.Append(Line.CreateBound(p1, p2));
                loop.Append(Line.CreateBound(p2, p3));
                loop.Append(Line.CreateBound(p3, p0));
                return loop;
            }
            catch { return null; }
        }
    }
}

using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 部分接触がある面について、接触領域を差し引いた残りの矩形から
    /// 薄板 Solid を構築する (Phase 2 視覚化厳密化)。
    ///
    /// 【前提】面 A の外形が「軸平行矩形」である。
    ///   Solid Union/Join の結果、同軸線上でエッジが分割されていてもOK。
    ///   L字型・切り欠きあり・穴あり等は面積チェックで除外してフォールバックする。
    /// 【方式】矩形ベースの 2D 差分: A の矩形から各 PartialContact の投影矩形を
    ///   順次減算し、最大 4 つのサブ矩形に分解する。各サブ矩形から薄板 Solid を作る。
    /// </summary>
    internal static class PartialContactClipper
    {
        private const double UvEps = 1e-6;
        private const double ThicknessFeet = 0.03;  // 約 10mm
        private const double AreaMatchTolerance = 0.02;  // Face.Area と矩形面積の許容差 ±2%

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
            public override string ToString()
            {
                return $"[U:{UMin:F3}..{UMax:F3}, V:{VMin:F3}..{VMax:F3}]";
            }
        }

        internal static ClipResult TryClip(FaceClassifier.FaceInfo fi)
        {
            var result = new ClipResult();
            if (fi?.Face == null) { result.FailReason = "face-null"; return result; }
            if (fi.PartialContacts == null || fi.PartialContacts.Count == 0)
            { result.FailReason = "no-partial-contacts"; return result; }

            bool log = FormworkDebugLog.Enabled;

            // 1. CurveLoop 取得
            IList<CurveLoop> loops;
            try { loops = fi.Face.GetEdgesAsCurveLoops(); }
            catch { result.FailReason = "get-loops-failed"; return result; }
            if (loops == null || loops.Count == 0)
            { result.FailReason = "no-loops"; return result; }

            // 単一外周想定。内周 (穴) があると複雑なのでフォールバック。
            if (loops.Count > 1)
            { result.FailReason = $"has-holes(count={loops.Count})"; return result; }

            // 2. 矩形判定
            var rectMain = TryExtractRectUV(fi.Face, loops[0], log, out string extractDetail);
            if (rectMain == null)
            { result.FailReason = "not-rectangular:" + extractDetail; return result; }

            if (log)
            {
                FormworkDebugLog.Log($"  [Clipper] face area={fi.Face.Area:F4} rect={rectMain.Value} partialContacts={fi.PartialContacts.Count}");
            }

            // 3. PartialContact の UvBoundsOnA をクリップ (A の境界内にクランプ)
            var bRects = new List<RectUV>();
            foreach (var pc in fi.PartialContacts)
            {
                if (pc.UvBoundsOnA == null)
                { result.FailReason = "no-uv-on-a"; return result; }

                // 投影が A の境界外にはみ出している場合はクランプ
                var br = new RectUV
                {
                    UMin = Math.Max(pc.UvBoundsOnA.Min.U, rectMain.Value.UMin),
                    VMin = Math.Max(pc.UvBoundsOnA.Min.V, rectMain.Value.VMin),
                    UMax = Math.Min(pc.UvBoundsOnA.Max.U, rectMain.Value.UMax),
                    VMax = Math.Min(pc.UvBoundsOnA.Max.V, rectMain.Value.VMax),
                };
                if (br.UMax > br.UMin && br.VMax > br.VMin)
                {
                    bRects.Add(br);
                    if (log) FormworkDebugLog.Log($"  [Clipper] subtract E{pc.OtherElementId}/f{pc.OtherFaceIndex} raw={FmtUvBounds(pc.UvBoundsOnA)} clipped={br}");
                }
                else if (log)
                {
                    FormworkDebugLog.Log($"  [Clipper] skip E{pc.OtherElementId}/f{pc.OtherFaceIndex} (out of bounds) raw={FmtUvBounds(pc.UvBoundsOnA)}");
                }
            }
            if (bRects.Count == 0)
            { result.FailReason = "no-valid-b-rects"; return result; }

            double uDim = rectMain.Value.UMax - rectMain.Value.UMin;
            double vDim = rectMain.Value.VMax - rectMain.Value.VMin;
            double degenTol = Math.Max(uDim, vDim) * 1e-4;

            // 4. 順次差分
            var pieces = new List<RectUV> { rectMain.Value };
            foreach (var br in bRects)
            {
                var next = new List<RectUV>();
                foreach (var p in pieces)
                {
                    foreach (var sub in Subtract(p, br))
                    {
                        if (!sub.IsDegenerate(degenTol)) next.Add(sub);
                    }
                }
                pieces = next;
                if (pieces.Count == 0) break;
            }

            if (log)
            {
                FormworkDebugLog.Log($"  [Clipper] pieces after subtraction: {pieces.Count}");
                foreach (var p in pieces)
                    FormworkDebugLog.Log($"    piece={p}");
            }

            if (pieces.Count == 0)
            { result.FailReason = "fully-clipped"; return result; }

            // 5. 各サブ矩形から薄板 Solid を作る
            XYZ normal = fi.Normal?.Normalize();
            if (normal == null) { result.FailReason = "no-normal"; return result; }

            int solidBuildFails = 0;
            foreach (var piece in pieces)
            {
                CurveLoop loop = BuildCurveLoopFromRect(fi.Face, piece);
                if (loop == null) { solidBuildFails++; continue; }

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
                else solidBuildFails++;
            }

            if (result.Solids.Count == 0)
            { result.FailReason = $"solid-build-failed(attempts={pieces.Count},fails={solidBuildFails})"; return result; }

            result.Success = true;
            if (log)
                FormworkDebugLog.Log($"  [Clipper] SUCCESS: created {result.Solids.Count} solids (failed {solidBuildFails})");
            return result;
        }

        /// <summary>
        /// Face の外形が軸平行矩形ならその UV 範囲を返す。
        /// エッジが同軸線上で複数に分割されていてもOK。
        /// L字型・切り欠き・凹型は Face.Area vs 矩形面積で除外する。
        /// </summary>
        private static RectUV? TryExtractRectUV(Face face, CurveLoop loop, bool log, out string detail)
        {
            detail = null;
            if (loop == null) { detail = "loop-null"; return null; }

            // (1) 全エッジが直線か確認
            var endpoints = new List<UV>();
            var endpoints3D = new List<XYZ>();
            int lineCount = 0;
            foreach (Curve c in loop)
            {
                if (!(c is Line))
                {
                    detail = "non-line-edge(" + c.GetType().Name + ")";
                    return null;
                }
                lineCount++;
                var p0 = c.GetEndPoint(0);
                var p1 = c.GetEndPoint(1);
                endpoints3D.Add(p0);
                endpoints3D.Add(p1);
                IntersectionResult proj0 = null, proj1 = null;
                try { proj0 = face.Project(p0); } catch { }
                try { proj1 = face.Project(p1); } catch { }
                if (proj0?.UVPoint == null || proj1?.UVPoint == null)
                {
                    detail = "project-null";
                    return null;
                }
                endpoints.Add(proj0.UVPoint);
                endpoints.Add(proj1.UVPoint);
            }

            if (lineCount < 4) { detail = $"too-few-edges({lineCount})"; return null; }

            // (2) UV AABB を計算
            double uMin = double.MaxValue, uMax = double.MinValue;
            double vMin = double.MaxValue, vMax = double.MinValue;
            foreach (var uv in endpoints)
            {
                if (uv.U < uMin) uMin = uv.U;
                if (uv.U > uMax) uMax = uv.U;
                if (uv.V < vMin) vMin = uv.V;
                if (uv.V > vMax) vMax = uv.V;
            }
            double uDim = uMax - uMin;
            double vDim = vMax - vMin;
            if (uDim <= UvEps || vDim <= UvEps)
            {
                detail = "degenerate-uv";
                return null;
            }

            // (3) 全頂点が AABB の境界上にあるか (内部点があれば L字やノッチ)
            double tolU = uDim * 5e-3;   // 0.5% 許容
            double tolV = vDim * 5e-3;
            int interiorPoints = 0;
            foreach (var uv in endpoints)
            {
                bool onUEdge = Math.Abs(uv.U - uMin) < tolU || Math.Abs(uv.U - uMax) < tolU;
                bool onVEdge = Math.Abs(uv.V - vMin) < tolV || Math.Abs(uv.V - vMax) < tolV;
                if (!onUEdge && !onVEdge) interiorPoints++;
            }
            if (interiorPoints > 0)
            {
                detail = $"interior-points({interiorPoints})";
                return null;
            }

            // (4) 矩形面積 vs Face.Area で最終確認 (凹型を除外)
            XYZ p00 = null, p10 = null, p01 = null;
            try
            {
                p00 = face.Evaluate(new UV(uMin, vMin));
                p10 = face.Evaluate(new UV(uMax, vMin));
                p01 = face.Evaluate(new UV(uMin, vMax));
            }
            catch { }
            if (p00 == null || p10 == null || p01 == null)
            {
                detail = "corner-evaluate-failed";
                return null;
            }
            double lenU = p00.DistanceTo(p10);
            double lenV = p00.DistanceTo(p01);
            double rectArea = lenU * lenV;

            double faceArea = 0;
            try { faceArea = face.Area; } catch { }
            if (faceArea <= 1e-6)
            {
                detail = "face-area-zero";
                return null;
            }

            double ratio = rectArea / faceArea;
            if (ratio < 1.0 - AreaMatchTolerance || ratio > 1.0 + AreaMatchTolerance)
            {
                detail = $"area-mismatch(rect={rectArea:F3},face={faceArea:F3},ratio={ratio:F3})";
                return null;
            }

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

            // 左
            if (a.UMin < uMinX - UvEps)
                result.Add(new RectUV { UMin = a.UMin, VMin = a.VMin, UMax = uMinX, VMax = a.VMax });
            // 右
            if (a.UMax > uMaxX + UvEps)
                result.Add(new RectUV { UMin = uMaxX, VMin = a.VMin, UMax = a.UMax, VMax = a.VMax });
            // 下
            if (a.VMin < vMinX - UvEps)
                result.Add(new RectUV { UMin = uMinX, VMin = a.VMin, UMax = uMaxX, VMax = vMinX });
            // 上
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

        private static string FmtUvBounds(BoundingBoxUV bb)
        {
            if (bb == null) return "null";
            return $"[U:{bb.Min.U:F3}..{bb.Max.U:F3},V:{bb.Min.V:F3}..{bb.Max.V:F3}]";
        }
    }
}

using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 隣接要素のソリッド内部に「埋もれている」FormworkRequired 面を検出して
    /// DeductedContact に降格する。ContactFaceDetector では拾えない 2 つの盲点を埋める:
    ///
    ///   1) フラッシュ接触で BB ペアフィルタを通過できないケース
    ///      例: 梁の下段側面が隣接スラブ側面と同一平面で接しているが、ctx.BB の
    ///          精度不足 (床BB が実体より小さく算出される) で SpatialGrid が
    ///          両者をペアに含めず、接触検出が走らない。
    ///
    ///   2) 接合されずに体積が重なっている要素の埋没面
    ///      例: スラブと布基礎が Revit の「結合」されておらず体積が干渉している。
    ///          スラブ側面が基礎ソリッド内部にあり、対応する基礎側面が無いため
    ///          ContactFaceDetector の「対向平行面の接触」判定が成立しない。
    ///
    /// 【判定方法】
    ///   各 FormworkRequired 面 f について:
    ///     p_out = (face center) + (outward normal) × SampleOffsetFeet
    ///   この p_out が他要素のソリッド内部にあれば、その面は埋もれている。
    ///
    /// 【ソリッド内包判定】
    ///   凸ソリッド前提の平面署名距離テスト: ソリッドの全平面状の面について
    ///   p が「面の内向き側」(signedDist ≤ tol) にあれば内部とみなす。
    ///   RC 躯体 (柱/梁/壁/スラブ/基礎) は概ね凸または準凸なので十分な精度。
    ///   非平面 (CylindricalFace 等) はスキップする。
    /// </summary>
    internal static class BuriedFaceDetector
    {
        // 面中心から外向きに少し動かすオフセット (≈ 5mm)
        // フラッシュ接触面でも確実に隣接要素のソリッド内部に入る量。
        private const double SampleOffsetFeet = 0.0164;

        // ソリッド内包判定の許容差 (≈ 1.5mm)。フラッシュ接触で
        // 境界に乗ったときも内部とみなすために少し緩める。
        private const double InsideTolFeet = 0.005;

        // 候補要素 BB の事前フィルタ余裕 (≈ 15mm)。
        // ctx.BB が実体より小さく報告されるケースに対する保険。
        // 内部判定の前段の高速フィルタなので大きめでも害は少ない。
        private const double BBoxMarginFeet = 0.05;

        internal static void Run(List<ContactFaceDetector.ElementFacesContext> contexts,
            Document doc = null)
        {
            if (contexts == null || contexts.Count == 0) return;

            FormworkDebugLog.Section($"Pass 2b: Buried Face Detection (elements={contexts.Count})");

            // ctx.BB は ComputeWorldBoundingBox の精度問題で実体より小さく算出される
            // ケースがあるため、ソリッドのエッジ端点 + 面 UV 中心点から再構築する。
            var robustBBs = BuildRobustBBs(contexts);

            // GenericModel カテゴリの「障害物」を別途収集する。
            // 一般モデル要素 (設備・什器・埋め込み物等) に接している RC 面は型枠不要として
            // 控除する目的。formwork DS は除外する (旧版の OST_GenericModel DS や、現行版の
            // OST_NurseCallDevices DS は FormworkMarker パラメータで識別)。
            var genericModelObstacles = doc != null ? CollectGenericModelObstacles(doc) : new List<GenericModelObstacle>();
            if (genericModelObstacles.Count > 0)
            {
                FormworkDebugLog.Log(
                    $"  [Buried] GenericModel obstacles collected: {genericModelObstacles.Count}");
            }

            int demoted = 0;
            int demotedByGm = 0;
            int pointInSolidChecks = 0;

            foreach (var ctxA in contexts)
            {
                if (ctxA?.Faces == null) continue;

                for (int fi = 0; fi < ctxA.Faces.Count; fi++)
                {
                    var face = ctxA.Faces[fi];
                    if (face == null || face.Face == null || face.Normal == null) continue;
                    if (face.FaceType != FaceType.FormworkRequired) continue;

                    XYZ pCenter = TryGetFaceCenter(face.Face);
                    if (pCenter == null) continue;

                    XYZ pOut = pCenter + face.Normal.Multiply(SampleOffsetFeet);

                    bool faceDemoted = false;
                    foreach (var ctxB in contexts)
                    {
                        if (ctxB == null) continue;
                        if (ctxB.ElementId == ctxA.ElementId) continue;
                        if (ctxB.Solids == null || ctxB.Solids.Count == 0) continue;

                        // 既存の ContactFaceDetector が「小面 ctxB が大面 face に部分接触する」
                        // パターンを PartialContact として記録済みのケースをスキップ。
                        // 例: 梁端 (小面) が柱側面 (大面) に当たる場合、柱側面の中心は
                        // 梁体積内に入るが、梁からはみ出た部分は型枠が必要なので
                        // 面全体を DeductedContact に降格してはならない。
                        if (HasPartialContactWith(face, ctxB.ElementId)) continue;

                        if (!robustBBs.TryGetValue(ctxB.ElementId, out var bbB) || bbB == null) continue;
                        if (!IsPointInsideBBox(pOut, bbB, BBoxMarginFeet)) continue;

                        pointInSolidChecks++;
                        if (IsPointInsideAnySolid(pOut, ctxB.Solids, InsideTolFeet))
                        {
                            face.FaceType = FaceType.DeductedContact;
                            demoted++;
                            faceDemoted = true;
                            if (FormworkDebugLog.Enabled)
                            {
                                FormworkDebugLog.Log(
                                    $"  [Buried] E{ctxA.ElementId} face[{fi}] " +
                                    $"n=({face.Normal.X:F2},{face.Normal.Y:F2},{face.Normal.Z:F2}) " +
                                    $"area={face.Area:F3}ft² → DeductedContact " +
                                    $"(inside E{ctxB.ElementId}/{ctxB.CategoryName})");
                            }
                            break;
                        }
                    }

                    // GenericModel 障害物との照合 (既に降格された面はスキップ)
                    if (!faceDemoted && genericModelObstacles.Count > 0)
                    {
                        foreach (var gm in genericModelObstacles)
                        {
                            if (!IsPointInsideBBox(pOut, gm.BB, BBoxMarginFeet)) continue;
                            pointInSolidChecks++;
                            if (IsPointInsideAnySolid(pOut, gm.Solids, InsideTolFeet))
                            {
                                face.FaceType = FaceType.DeductedContact;
                                demoted++;
                                demotedByGm++;
                                if (FormworkDebugLog.Enabled)
                                {
                                    FormworkDebugLog.Log(
                                        $"  [Buried] E{ctxA.ElementId} face[{fi}] " +
                                        $"n=({face.Normal.X:F2},{face.Normal.Y:F2},{face.Normal.Z:F2}) " +
                                        $"area={face.Area:F3}ft² → DeductedContact " +
                                        $"(inside GenericModel E{gm.ElementId})");
                                }
                                break;
                            }
                        }
                    }
                }
            }

            FormworkDebugLog.Section("Pass 2b Summary");
            FormworkDebugLog.Log($"point-in-solid checks: {pointInSolidChecks}");
            FormworkDebugLog.Log($"faces demoted to contact: {demoted} (by GenericModel: {demotedByGm})");
            FormworkDebugLog.Flush();
        }

        /// <summary>
        /// GenericModel カテゴリの「障害物」要素を収集する。
        /// formwork DS (現行 OST_NurseCallDevices および旧 OST_GenericModel の DS) は
        /// FormworkMarker パラメータで識別して除外する。
        /// </summary>
        private class GenericModelObstacle
        {
            public int ElementId;
            public BoundingBoxXYZ BB;
            public List<Solid> Solids;
        }

        private static List<GenericModelObstacle> CollectGenericModelObstacles(Document doc)
        {
            var result = new List<GenericModelObstacle>();
            try
            {
                var collector = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .WhereElementIsNotElementType();
                foreach (Element e in collector)
                {
                    if (e == null) continue;
                    // formwork DS は除外 (FormworkMarker で識別)
                    try
                    {
                        var p = e.LookupParameter(FormworkParameterManager.ParamMarker);
                        if (p != null && p.StorageType == StorageType.String)
                        {
                            string val = p.AsString();
                            if (!string.IsNullOrEmpty(val) &&
                                val.StartsWith(FormworkParameterManager.MarkerValue)) continue;
                        }
                    }
                    catch { }

                    var solids = SolidUnionProcessor.GetSolids(e, null,
                        skipWallSweepRetry: true);
                    if (solids.Count == 0) continue;

                    var bb = ComputeBBFromSolids(solids);
                    if (bb == null) continue;

                    result.Add(new GenericModelObstacle
                    {
                        ElementId = e.Id.IntValue(),
                        BB = bb,
                        Solids = solids,
                    });
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Buried] CollectGenericModelObstacles EX: {ex.Message}");
            }
            return result;
        }

        private static BoundingBoxXYZ ComputeBBFromSolids(List<Solid> solids)
        {
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            bool any = false;
            foreach (var s in solids)
            {
                if (s == null) continue;
                foreach (Edge e in s.Edges)
                {
                    Curve c = null;
                    try { c = e.AsCurve(); } catch { }
                    if (c == null) continue;
                    for (int i = 0; i <= 1; i++)
                    {
                        XYZ p;
                        try { p = c.GetEndPoint(i); } catch { continue; }
                        if (p == null) continue;
                        if (p.X < minX) minX = p.X;
                        if (p.Y < minY) minY = p.Y;
                        if (p.Z < minZ) minZ = p.Z;
                        if (p.X > maxX) maxX = p.X;
                        if (p.Y > maxY) maxY = p.Y;
                        if (p.Z > maxZ) maxZ = p.Z;
                        any = true;
                    }
                }
            }
            return any ? new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ),
            } : null;
        }

        /// <summary>
        /// 各要素について、ソリッドのエッジ端点と全 face の UV 中心点を統合した
        /// 信頼できる BoundingBox を構築する。ctx.BB が実体より小さい場合の保険。
        /// </summary>
        private static Dictionary<int, BoundingBoxXYZ> BuildRobustBBs(
            List<ContactFaceDetector.ElementFacesContext> contexts)
        {
            var result = new Dictionary<int, BoundingBoxXYZ>(contexts.Count);
            foreach (var ctx in contexts)
            {
                if (ctx == null) continue;

                double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
                double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
                bool any = false;

                if (ctx.Solids != null)
                {
                    foreach (var s in ctx.Solids)
                    {
                        if (s == null) continue;
                        foreach (Edge e in s.Edges)
                        {
                            Curve c = null;
                            try { c = e.AsCurve(); } catch { }
                            if (c == null) continue;
                            for (int i = 0; i <= 1; i++)
                            {
                                XYZ p;
                                try { p = c.GetEndPoint(i); } catch { continue; }
                                if (p == null) continue;
                                if (p.X < minX) minX = p.X;
                                if (p.Y < minY) minY = p.Y;
                                if (p.Z < minZ) minZ = p.Z;
                                if (p.X > maxX) maxX = p.X;
                                if (p.Y > maxY) maxY = p.Y;
                                if (p.Z > maxZ) maxZ = p.Z;
                                any = true;
                            }
                        }
                    }
                }

                // 念のため面 UV 中心点も BB に含める。エッジ端点が
                // 何らかの理由で取れなかった面の位置をカバーする。
                if (ctx.Faces != null)
                {
                    foreach (var fi in ctx.Faces)
                    {
                        if (fi?.Face == null) continue;
                        XYZ p = TryGetFaceCenter(fi.Face);
                        if (p == null) continue;
                        if (p.X < minX) minX = p.X;
                        if (p.Y < minY) minY = p.Y;
                        if (p.Z < minZ) minZ = p.Z;
                        if (p.X > maxX) maxX = p.X;
                        if (p.Y > maxY) maxY = p.Y;
                        if (p.Z > maxZ) maxZ = p.Z;
                        any = true;
                    }
                }

                if (any)
                {
                    result[ctx.ElementId] = new BoundingBoxXYZ
                    {
                        Min = new XYZ(minX, minY, minZ),
                        Max = new XYZ(maxX, maxY, maxZ),
                    };
                }
                else if (ctx.BB != null)
                {
                    result[ctx.ElementId] = ctx.BB;
                }
            }
            return result;
        }

        private static bool HasPartialContactWith(FaceClassifier.FaceInfo face, int otherElementId)
        {
            if (face.PartialContacts == null) return false;
            foreach (var pc in face.PartialContacts)
            {
                if (pc.OtherElementId == otherElementId) return true;
            }
            return false;
        }

        private static XYZ TryGetFaceCenter(Face f)
        {
            try
            {
                BoundingBoxUV bb = f.GetBoundingBox();
                if (bb == null) return null;
                UV mid = (bb.Min + bb.Max) * 0.5;
                return f.Evaluate(mid);
            }
            catch { return null; }
        }

        private static bool IsPointInsideBBox(XYZ p, BoundingBoxXYZ bb, double margin)
        {
            return p.X >= bb.Min.X - margin && p.X <= bb.Max.X + margin
                && p.Y >= bb.Min.Y - margin && p.Y <= bb.Max.Y + margin
                && p.Z >= bb.Min.Z - margin && p.Z <= bb.Max.Z + margin;
        }

        private static bool IsPointInsideAnySolid(XYZ p, List<Solid> solids, double tol)
        {
            foreach (var s in solids)
            {
                if (s == null) continue;
                // 高速パス: 凸前提の平面署名距離テスト。全平面が「面の内向き側」を
                // 許容すれば凸ソリッドの場合は内部確定。
                if (IsPointInsideConvexSolid(p, s, tol)) return true;
                // 凸テストが否定的だった場合、ソリッドが非凸 (例: 段差を持つ梁) の
                // 可能性があるため、レイキャストで再判定する。
                if (IsPointInsideSolidRayCast(p, s)) return true;
            }
            return false;
        }

        /// <summary>
        /// 凸ソリッド前提の内包判定。ソリッドの全平面状の面について
        /// p が「面の外向き側」に出ていないことを確認する。
        /// 平面状の面が 1 枚も無いソリッド (純粋曲面ソリッド) は判定不能として false。
        /// </summary>
        private static bool IsPointInsideConvexSolid(XYZ p, Solid s, double tol)
        {
            bool anyPlanar = false;
            foreach (Face f in s.Faces)
            {
                if (!(f is PlanarFace pf)) continue;
                anyPlanar = true;
                XYZ outwardN = pf.FaceNormal;
                if (outwardN == null) continue;
                XYZ origin = pf.Origin;
                if (origin == null) continue;
                double signedDist = (p - origin).DotProduct(outwardN);
                if (signedDist > tol) return false;
            }
            return anyPlanar;
        }

        // レイキャスト用の方向ベクトル: 軸非整列にすることで面のエッジに
        // ヒットして交差カウントが揺らぐリスクを軽減する。
        private static readonly XYZ RayCastDir = new XYZ(1.0, 0.123, 0.456).Normalize();
        private const double RayCastLengthFeet = 10000.0;  // 約 3km。要素モデル全域を貫通する長さ

        /// <summary>
        /// レイキャストによる内包判定。p から非整列方向に長い線を伸ばし、
        /// ソリッドの全面との交差数を数える。奇数なら内部。
        /// 非凸ソリッド (段差付き梁等) でも正しく判定できる。
        /// </summary>
        private static bool IsPointInsideSolidRayCast(XYZ p, Solid s)
        {
            XYZ far = p + RayCastDir.Multiply(RayCastLengthFeet);
            Line ray;
            try { ray = Line.CreateBound(p, far); }
            catch { return false; }

            int crossings = 0;
            foreach (Face f in s.Faces)
            {
                if (f == null) continue;
                IntersectionResultArray ira = null;
                SetComparisonResult result;
                try { result = f.Intersect(ray, out ira); }
                catch { continue; }
                if (result == SetComparisonResult.Overlap && ira != null)
                    crossings += ira.Size;
            }
            return (crossings & 1) == 1;
        }
    }
}

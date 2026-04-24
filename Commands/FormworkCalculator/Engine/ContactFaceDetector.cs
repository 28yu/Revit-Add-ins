using System;
using System.Collections.Generic;
using System.Globalization;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 要素同士の接触面を検出する。2段階判定方式。
    ///
    /// 【Stage 1: Full Contact】(従来の動作)
    ///   面 a の中心が面 b の内部に投影でき、a の面積 ≤ b の面積 × 1.5
    ///   → 面 a 全体を DeductedContact に変更
    ///   (小面が大面に当たる典型: 梁端面→柱側面、壁端面→主壁側面)
    ///
    /// 【Stage 2: Partial Contact】(Phase 1 追加)
    ///   面 a の一部に小さな面 b が接している (a > b × 1.5 だが中心投影は成立)
    ///   → 面 a は FormworkRequired のまま、PartialContact リストに b の情報を追加
    ///   (T字壁結合: 主壁側面の一部に端壁端面が当たるケース)
    ///
    /// 【候補の絞り込み】SpatialGrid で隣接セル内の要素だけ走査 → 大規模対応。
    /// </summary>
    internal class ContactFaceDetector
    {
        internal class ElementFacesContext
        {
            public int ElementId;
            public CategoryGroup Category;
            public string CategoryName;
            public BoundingBoxXYZ BB;
            public List<FaceClassifier.FaceInfo> Faces = new List<FaceClassifier.FaceInfo>();
        }

        private const double CoincidenceTolFeet = 0.05;   // ≈ 15mm
        private const double AreaRatioLimit = 1.5;
        private const double AntiParallelThreshold = -0.90;  // cos ≤ -0.9 (≈ 155°)

        internal static void RefineContactFaces(List<ElementFacesContext> contexts)
        {
            int n = contexts.Count;
            int contactChanges = 0;
            int partialContactAdded = 0;
            int pairsChecked = 0;
            int pairsBBoxOverlap = 0;

            FormworkDebugLog.Section($"Pass 2: Contact Face Detection (elements={n})");

            // 空間索引で候補絞り込み
            var grid = SpatialGrid.Build(contexts, CoincidenceTolFeet);
            FormworkDebugLog.Log($"SpatialGrid cell size: {grid.CellSize:F3} ft");

            var ctxById = new Dictionary<int, ElementFacesContext>();
            foreach (var c in contexts) ctxById[c.ElementId] = c;

            foreach (var ci in contexts)
            {
                if (ci.BB == null) continue;

                foreach (var otherId in grid.GetCandidates(ci.ElementId))
                {
                    if (!ctxById.TryGetValue(otherId, out var cj)) continue;
                    pairsChecked++;
                    pairsBBoxOverlap++;

                    bool pairLogged = false;

                    for (int fiIdx = 0; fiIdx < ci.Faces.Count; fiIdx++)
                    {
                        var fi = ci.Faces[fiIdx];

                        // FormworkRequired 面のみが対象
                        if (fi.FaceType != FaceType.FormworkRequired) continue;

                        // 既に Full Contact 化された面は以降スキップ
                        bool fullyContacted = false;

                        for (int fjIdx = 0; fjIdx < cj.Faces.Count; fjIdx++)
                        {
                            if (fullyContacted) break;
                            var fj = cj.Faces[fjIdx];
                            var result = EvaluateContact(
                                ci, fi, fiIdx,
                                cj, fj, fjIdx,
                                ref pairLogged);

                            if (result.Kind == ContactKind.Full)
                            {
                                fi.FaceType = FaceType.DeductedContact;
                                contactChanges++;
                                fullyContacted = true;
                            }
                            else if (result.Kind == ContactKind.Partial)
                            {
                                // 重複を防ぐ (同じ相手要素・同じ面を2回加算しない)
                                bool dup = false;
                                foreach (var pc in fi.PartialContacts)
                                {
                                    if (pc.OtherElementId == cj.ElementId &&
                                        pc.OtherFaceIndex == fjIdx) { dup = true; break; }
                                }
                                if (!dup)
                                {
                                    fi.PartialContacts.Add(new FaceClassifier.PartialContact
                                    {
                                        OtherElementId = cj.ElementId,
                                        OtherFaceIndex = fjIdx,
                                        ContactArea = result.ContactArea,
                                        UvBounds = result.UvBounds,
                                        UvBoundsOnA = result.UvBoundsOnA,
                                    });
                                    partialContactAdded++;
                                }
                            }
                        }
                    }
                }
            }

            FormworkDebugLog.Section("Pass 2 Summary");
            FormworkDebugLog.Log($"pairs checked (grid):        {pairsChecked}");
            FormworkDebugLog.Log($"pairs with BB overlap:       {pairsBBoxOverlap}");
            FormworkDebugLog.Log($"full contact changes:        {contactChanges}");
            FormworkDebugLog.Log($"partial contacts added:      {partialContactAdded}");
            FormworkDebugLog.Flush();
        }

        private enum ContactKind { None, Full, Partial }

        private struct ContactResult
        {
            public ContactKind Kind;
            public double ContactArea;
            public BoundingBoxUV UvBounds;       // B 自身の UV (debug)
            public BoundingBoxUV UvBoundsOnA;    // A 面上での B の投影 UV (Phase 2 用)
        }

        /// <summary>
        /// 面 a に対する面 b の接触を評価する。None / Full / Partial のいずれかを返す。
        /// </summary>
        private static ContactResult EvaluateContact(
            ElementFacesContext ca, FaceClassifier.FaceInfo a, int aIdx,
            ElementFacesContext cb, FaceClassifier.FaceInfo b, int bIdx,
            ref bool pairHeaderLogged)
        {
            var none = new ContactResult { Kind = ContactKind.None };
            if (a?.Face == null || b?.Face == null) return none;
            if (a.Normal == null || b.Normal == null) return none;

            bool log = FormworkDebugLog.Enabled;

            // --- 条件 1: 反平行 ---
            double dot = a.Normal.DotProduct(b.Normal);
            bool cond1 = dot <= AntiParallelThreshold;
            if (!cond1) return none;

            // --- 面積取得 ---
            double aArea = 0, bArea = 0;
            try { aArea = a.Face.Area; bArea = b.Face.Area; } catch { }
            bool cond2Area = aArea > 1e-6 && bArea > 1e-6;

            // --- 投影と距離 ---
            BoundingBoxUV bbA = null;
            try { bbA = a.Face.GetBoundingBox(); } catch { }
            XYZ pA = null;
            if (bbA != null) { try { pA = a.Face.Evaluate((bbA.Min + bbA.Max) * 0.5); } catch { } }

            IntersectionResult proj = null;
            if (pA != null) { try { proj = b.Face.Project(pA); } catch { } }
            bool cond3Dist = proj != null && proj.Distance <= CoincidenceTolFeet;

            UV uv = proj?.UVPoint;
            BoundingBoxUV bbB = null;
            try { bbB = b.Face.GetBoundingBox(); } catch { }

            bool cond4UVInside = false;
            bool cond4NotNearBoundary = false;
            if (uv != null && bbB != null)
            {
                double marginU = (bbB.Max.U - bbB.Min.U) * 0.01;
                double marginV = (bbB.Max.V - bbB.Min.V) * 0.01;

                cond4UVInside =
                    uv.U >= bbB.Min.U - 1e-6 && uv.U <= bbB.Max.U + 1e-6 &&
                    uv.V >= bbB.Min.V - 1e-6 && uv.V <= bbB.Max.V + 1e-6;

                bool nearBoundary =
                    uv.U < bbB.Min.U + marginU || uv.U > bbB.Max.U - marginU ||
                    uv.V < bbB.Min.V + marginV || uv.V > bbB.Max.V - marginV;

                cond4NotNearBoundary = !nearBoundary ||
                    (proj != null && proj.Distance <= CoincidenceTolFeet * 0.5);
            }

            // --- Stage 1: Full Contact 判定 (従来ロジック a→b の向き) ---
            bool stage1Accepted = false;
            string stage1Reason = null;
            if (!cond2Area) stage1Reason = "cond2-area-zero";
            else if (aArea > bArea * AreaRatioLimit) stage1Reason = "cond2-area-ratio";
            else if (bbA == null) stage1Reason = "cond3-bbA-null";
            else if (pA == null) stage1Reason = "cond3-pA-null";
            else if (proj == null) stage1Reason = "cond3-project-null";
            else if (!cond3Dist) stage1Reason = "cond3-distance";
            else if (uv == null) stage1Reason = "cond4-uv-null";
            else if (bbB == null) stage1Reason = "cond4-bbB-null";
            else if (!cond4UVInside) stage1Reason = "cond4-uv-out-of-bounds";
            else if (!cond4NotNearBoundary) stage1Reason = "cond4-near-boundary";
            else stage1Accepted = true;

            // --- Stage 2: Partial Contact 判定 (a >> b の場合) ---
            //   大面 a に小面 b が部分的に接触しているケース
            //   判定: b の中心を a に投影して、a 内部にあり距離≈0ならば部分接触
            ContactResult stage2 = none;
            string stage2Reason = null;
            if (!stage1Accepted)
            {
                // 条件: aがbより大きい (stage1がarea-ratioで弾かれた) 場合のみ stage2 を評価
                if (stage1Reason == "cond2-area-ratio")
                {
                    stage2 = EvaluatePartialBToA(a, b, out stage2Reason);
                }
                else
                {
                    stage2Reason = "not-evaluated";
                }
            }

            // --- ログ出力 ---
            if (log)
            {
                if (!pairHeaderLogged)
                {
                    FormworkDebugLog.Log(
                        $"[Pair E{ca.ElementId}({ca.Category}) x E{cb.ElementId}({cb.Category})]");
                    pairHeaderLogged = true;
                }

                var sb = new System.Text.StringBuilder();
                sb.Append("  f").Append(aIdx).Append("->f").Append(bIdx);
                sb.Append(" dot=").Append(Fmt(dot, 3));
                sb.Append(" aArea=").Append(Fmt(aArea, 4));
                sb.Append(" bArea=").Append(Fmt(bArea, 4));
                sb.Append(" d=").Append(proj != null ? Fmt(proj.Distance, 4) : "-");
                sb.Append(" uv=").Append(FmtUV(uv));
                if (stage1Accepted)
                {
                    sb.Append(" FULL_CONTACT");
                }
                else if (stage2.Kind == ContactKind.Partial)
                {
                    sb.Append(" PARTIAL_CONTACT area=").Append(Fmt(stage2.ContactArea, 4));
                }
                else
                {
                    sb.Append(" REJECTED(s1=").Append(stage1Reason ?? "-")
                      .Append(",s2=").Append(stage2Reason ?? "-").Append(")");
                }
                FormworkDebugLog.Log(sb.ToString());
            }

            if (stage1Accepted) return new ContactResult { Kind = ContactKind.Full };
            if (stage2.Kind == ContactKind.Partial) return stage2;
            return none;
        }

        /// <summary>
        /// 面 b の中心を面 a に投影して、a 内部に b が「一部として」収まるかを評価する。
        /// 成立するなら面 b の面積を Partial Contact 面積として返す。
        /// (a >> b の前提で、b 全体が a に覆われている想定)
        ///
        /// Phase 2 用に、b の 4 隅を a に投影した UV 矩形 (UvBoundsOnA) も算出する。
        /// </summary>
        private static ContactResult EvaluatePartialBToA(
            FaceClassifier.FaceInfo a, FaceClassifier.FaceInfo b, out string reason)
        {
            reason = null;
            var none = new ContactResult { Kind = ContactKind.None };

            if (a?.Face == null || b?.Face == null) { reason = "s2-face-null"; return none; }

            BoundingBoxUV bbB = null;
            try { bbB = b.Face.GetBoundingBox(); } catch { }
            if (bbB == null) { reason = "s2-bbB-null"; return none; }

            XYZ pB = null;
            try { pB = b.Face.Evaluate((bbB.Min + bbB.Max) * 0.5); } catch { }
            if (pB == null) { reason = "s2-pB-null"; return none; }

            IntersectionResult proj = null;
            try { proj = a.Face.Project(pB); } catch { }
            if (proj == null) { reason = "s2-project-null"; return none; }
            if (proj.Distance > CoincidenceTolFeet) { reason = "s2-distance"; return none; }

            UV uv = proj.UVPoint;
            BoundingBoxUV bbA = null;
            try { bbA = a.Face.GetBoundingBox(); } catch { }
            if (uv == null || bbA == null) { reason = "s2-uv-null"; return none; }

            // b の中心投影が a の UV 内部か? (境界付近でもOK、全体が覆われている前提)
            double marginU = (bbA.Max.U - bbA.Min.U) * 0.001;
            double marginV = (bbA.Max.V - bbA.Min.V) * 0.001;
            bool uvInside =
                uv.U >= bbA.Min.U - marginU && uv.U <= bbA.Max.U + marginU &&
                uv.V >= bbA.Min.V - marginV && uv.V <= bbA.Max.V + marginV;
            if (!uvInside) { reason = "s2-uv-out-of-bounds"; return none; }

            double bArea = 0;
            try { bArea = b.Face.Area; } catch { }
            if (bArea <= 1e-6) { reason = "s2-b-area-zero"; return none; }

            // Phase 2: b の 4 隅を a に投影した UV 矩形を算出
            BoundingBoxUV uvOnA = ProjectBCornersToA(a.Face, b.Face, bbB);

            return new ContactResult
            {
                Kind = ContactKind.Partial,
                ContactArea = bArea,
                UvBounds = bbB,
                UvBoundsOnA = uvOnA,
            };
        }

        /// <summary>
        /// 面 b の 4 隅 (BoundingBoxUV の 4 隅) を面 a に投影し、
        /// 投影先 UV を囲む AABB (Axis-Aligned Bounding Box in UV) を返す。
        /// 投影失敗や全て境界外の場合は null を返す (→ Phase 2 フォールバック)。
        /// </summary>
        private static BoundingBoxUV ProjectBCornersToA(Face a, Face b, BoundingBoxUV bbB)
        {
            var corners = new UV[]
            {
                new UV(bbB.Min.U, bbB.Min.V),
                new UV(bbB.Max.U, bbB.Min.V),
                new UV(bbB.Max.U, bbB.Max.V),
                new UV(bbB.Min.U, bbB.Max.V),
            };

            double uMin = double.MaxValue, vMin = double.MaxValue;
            double uMax = double.MinValue, vMax = double.MinValue;
            int okCount = 0;

            foreach (var c in corners)
            {
                XYZ p;
                try { p = b.Evaluate(c); } catch { continue; }
                if (p == null) continue;

                IntersectionResult proj;
                try { proj = a.Project(p); } catch { continue; }
                if (proj == null || proj.UVPoint == null) continue;
                if (proj.Distance > CoincidenceTolFeet * 2) continue; // 許容 2倍

                var uv = proj.UVPoint;
                if (uv.U < uMin) uMin = uv.U;
                if (uv.V < vMin) vMin = uv.V;
                if (uv.U > uMax) uMax = uv.U;
                if (uv.V > vMax) vMax = uv.V;
                okCount++;
            }

            if (okCount < 3) return null; // 4隅中 3 つ以上投影できないと信頼性低い
            if (uMax <= uMin || vMax <= vMin) return null;

            return new BoundingBoxUV(uMin, vMin, uMax, vMax);
        }

        private static string Fmt(double v, int decimals)
        {
            return v.ToString("F" + decimals, CultureInfo.InvariantCulture);
        }

        private static string FmtUV(UV uv)
        {
            if (uv == null) return "-";
            return "(" + Fmt(uv.U, 3) + "," + Fmt(uv.V, 3) + ")";
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 要素同士の接触面（結合面）を検出して DeductedContact に変更する。
    ///
    /// 【アプローチ】直接幾何検査（ReferenceIntersector に依存しない）
    ///   BoundingBox が重なる要素ペアで、以下の条件を全て満たす場合に接触:
    ///     1. 法線が反平行 (a.N · b.N < -0.9)
    ///     2. 面積比 a ≤ b × 1.5
    ///        (a が b より大きすぎる場合は「大面積主壁の側面」なので除外)
    ///     3. 面 a の中心点 pt が面 b に perpendicular distance ≤ tol で投影できる
    ///     4. 投影点の UV が面 b の UV bounds 内部にある (境界上ではない)
    ///
    /// 【なぜ ReferenceIntersector を使わないか】
    ///   ReferenceIntersector は View3D の可視性制約やセクション・テンプレートに
    ///   依存し、密着面（Proximity ≈ 0）の検出が不安定。直接幾何検査は Revit の
    ///   ビュー状態に依存せず決定論的に動作する。
    ///
    /// 【判定が期待通りに動くケース】
    ///   - 壁 T 字結合: 端面（小）center は相手側面（大）上 → Contact
    ///     相手側面（大）center は端面（小）の外 → 非 Contact (主壁側面を残す)
    ///   - 梁柱取り合い: 梁端面（小）center は柱側面（大）上 → Contact
    ///   - 基礎上のフレーム: 梁底（小）center は基礎上面（大）上 → Contact
    ///   - スラブ結合: 端面同士 center が相互の面上 → 両方 Contact
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
            int pairsChecked = 0;
            int pairsBBoxOverlap = 0;

            FormworkDebugLog.Section($"Pass 2: Contact Face Detection (elements={n})");

            for (int i = 0; i < n; i++)
            {
                var ci = contexts[i];
                if (ci.BB == null) continue;
                for (int j = 0; j < n; j++)
                {
                    if (i == j) continue;
                    var cj = contexts[j];
                    if (cj.BB == null) continue;
                    pairsChecked++;
                    if (!BBoxesOverlap(ci.BB, cj.BB, CoincidenceTolFeet)) continue;
                    pairsBBoxOverlap++;

                    bool pairLogged = false;

                    for (int fiIdx = 0; fiIdx < ci.Faces.Count; fiIdx++)
                    {
                        var fi = ci.Faces[fiIdx];
                        if (fi.FaceType != FaceType.FormworkRequired) continue;

                        for (int fjIdx = 0; fjIdx < cj.Faces.Count; fjIdx++)
                        {
                            var fj = cj.Faces[fjIdx];
                            bool covered = IsFaceCovered(
                                ci, fi, fiIdx,
                                cj, fj, fjIdx,
                                ref pairLogged);
                            if (covered)
                            {
                                fi.FaceType = FaceType.DeductedContact;
                                contactChanges++;
                                break;
                            }
                        }
                    }
                }
            }

            FormworkDebugLog.Section("Pass 2 Summary");
            FormworkDebugLog.Log($"pairs checked:         {pairsChecked}");
            FormworkDebugLog.Log($"pairs with BB overlap: {pairsBBoxOverlap}");
            FormworkDebugLog.Log($"contact changes:       {contactChanges}");
            FormworkDebugLog.Flush();
        }

        private static bool BBoxesOverlap(BoundingBoxXYZ a, BoundingBoxXYZ b, double tol)
        {
            return a.Min.X - tol <= b.Max.X && b.Min.X - tol <= a.Max.X
                && a.Min.Y - tol <= b.Max.Y && b.Min.Y - tol <= a.Max.Y
                && a.Min.Z - tol <= b.Max.Z && b.Min.Z - tol <= a.Max.Z;
        }

        /// <summary>
        /// 面 a が面 b によって「覆われている」（中心が面 b 内部にある）かを判定。
        /// 全ての条件を評価し、最終的な採用/棄却とその理由を決定する。
        /// ログが有効ならば詳細を Formwork_debug.txt に出力する。
        /// </summary>
        private static bool IsFaceCovered(
            ElementFacesContext ca, FaceClassifier.FaceInfo a, int aIdx,
            ElementFacesContext cb, FaceClassifier.FaceInfo b, int bIdx,
            ref bool pairHeaderLogged)
        {
            if (a?.Face == null || b?.Face == null) return false;
            if (a.Normal == null || b.Normal == null) return false;

            bool log = FormworkDebugLog.Enabled;

            // --- 条件 1: 反平行 ---
            double dot = a.Normal.DotProduct(b.Normal);
            bool cond1 = dot <= AntiParallelThreshold;

            // 条件 1 で落ちる (反平行ではない) ペアは膨大なのでログを出さない
            if (!cond1) return false;

            // --- 条件 2: 面積比 ---
            double aArea = 0, bArea = 0;
            try { aArea = a.Face.Area; bArea = b.Face.Area; } catch { }
            bool cond2Area = aArea > 1e-6 && bArea > 1e-6;
            bool cond2Ratio = cond2Area && aArea <= bArea * AreaRatioLimit;

            // --- 面 a の中心点 ---
            BoundingBoxUV bbA = null;
            try { bbA = a.Face.GetBoundingBox(); } catch { }
            XYZ pA = null;
            if (bbA != null)
            {
                try { pA = a.Face.Evaluate((bbA.Min + bbA.Max) * 0.5); } catch { }
            }

            // --- 条件 3: 面 b への投影と距離 ---
            IntersectionResult proj = null;
            if (pA != null)
            {
                try { proj = b.Face.Project(pA); } catch { }
            }
            bool cond3Dist = proj != null && proj.Distance <= CoincidenceTolFeet;

            // --- 条件 4: 投影先 UV が面 b 内部 ---
            UV uv = proj?.UVPoint;
            BoundingBoxUV bbB = null;
            try { bbB = b.Face.GetBoundingBox(); } catch { }

            bool cond4UVInside = false;
            bool cond4NotNearBoundary = false;
            double marginU = 0, marginV = 0;
            if (uv != null && bbB != null)
            {
                marginU = (bbB.Max.U - bbB.Min.U) * 0.01;
                marginV = (bbB.Max.V - bbB.Min.V) * 0.01;

                cond4UVInside =
                    uv.U >= bbB.Min.U - 1e-6 && uv.U <= bbB.Max.U + 1e-6 &&
                    uv.V >= bbB.Min.V - 1e-6 && uv.V <= bbB.Max.V + 1e-6;

                bool nearBoundary =
                    uv.U < bbB.Min.U + marginU ||
                    uv.U > bbB.Max.U - marginU ||
                    uv.V < bbB.Min.V + marginV ||
                    uv.V > bbB.Max.V - marginV;

                // 境界上でも、perpendicular distance が十分小さければ接触と見なす
                cond4NotNearBoundary = !nearBoundary ||
                    (proj != null && proj.Distance <= CoincidenceTolFeet * 0.5);
            }

            // --- 最終判定と棄却理由 ---
            string rejectReason = null;
            bool accepted = false;
            if (!cond2Area) rejectReason = "cond2-area-zero";
            else if (!cond2Ratio) rejectReason = "cond2-area-ratio";
            else if (bbA == null) rejectReason = "cond3-bbA-null";
            else if (pA == null) rejectReason = "cond3-pA-null";
            else if (proj == null) rejectReason = "cond3-project-null";
            else if (!cond3Dist) rejectReason = "cond3-distance";
            else if (uv == null) rejectReason = "cond4-uv-null";
            else if (bbB == null) rejectReason = "cond4-bbB-null";
            else if (!cond4UVInside) rejectReason = "cond4-uv-out-of-bounds";
            else if (!cond4NotNearBoundary) rejectReason = "cond4-near-boundary";
            else accepted = true;

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
                sb.Append(" nA=").Append(FmtXYZ(a.Normal));
                sb.Append(" nB=").Append(FmtXYZ(b.Normal));
                sb.Append(" aArea=").Append(Fmt(aArea, 4));
                sb.Append(" bArea=").Append(Fmt(bArea, 4));
                sb.Append(" pA=").Append(FmtXYZ(pA));
                sb.Append(" d=").Append(proj != null ? Fmt(proj.Distance, 4) : "-");
                sb.Append(" uv=").Append(FmtUV(uv));
                sb.Append(" bbB=").Append(FmtBBoxUV(bbB));
                sb.Append(" ").Append(accepted ? "ACCEPTED" : "REJECTED(" + rejectReason + ")");
                FormworkDebugLog.Log(sb.ToString());
            }

            return accepted;
        }

        private static string Fmt(double v, int decimals)
        {
            return v.ToString("F" + decimals, CultureInfo.InvariantCulture);
        }

        private static string FmtXYZ(XYZ p)
        {
            if (p == null) return "-";
            return "(" + Fmt(p.X, 3) + "," + Fmt(p.Y, 3) + "," + Fmt(p.Z, 3) + ")";
        }

        private static string FmtUV(UV uv)
        {
            if (uv == null) return "-";
            return "(" + Fmt(uv.U, 3) + "," + Fmt(uv.V, 3) + ")";
        }

        private static string FmtBBoxUV(BoundingBoxUV bb)
        {
            if (bb == null) return "-";
            return "[" + Fmt(bb.Min.U, 3) + "," + Fmt(bb.Min.V, 3) + ")-("
                 + Fmt(bb.Max.U, 3) + "," + Fmt(bb.Max.V, 3) + "]";
        }
    }
}

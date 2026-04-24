using System;
using System.Collections.Generic;
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
            public BoundingBoxXYZ BB;
            public List<FaceClassifier.FaceInfo> Faces = new List<FaceClassifier.FaceInfo>();
        }

        private const double CoincidenceTolFeet = 0.05;   // ≈ 15mm
        private const double AreaRatioLimit = 1.5;
        private const double AntiParallelThreshold = -0.90;  // cos ≤ -0.9 (≈ 155°)

        internal static void RefineContactFaces(List<ElementFacesContext> contexts)
        {
            int n = contexts.Count;
            for (int i = 0; i < n; i++)
            {
                var ci = contexts[i];
                if (ci.BB == null) continue;
                for (int j = 0; j < n; j++)
                {
                    if (i == j) continue;
                    var cj = contexts[j];
                    if (cj.BB == null) continue;
                    if (!BBoxesOverlap(ci.BB, cj.BB, CoincidenceTolFeet)) continue;

                    foreach (var fi in ci.Faces)
                    {
                        if (fi.FaceType != FaceType.FormworkRequired) continue;

                        foreach (var fj in cj.Faces)
                        {
                            if (IsFaceCovered(fi, fj))
                            {
                                fi.FaceType = FaceType.DeductedContact;
                                break;
                            }
                        }
                    }
                }
            }
        }

        private static bool BBoxesOverlap(BoundingBoxXYZ a, BoundingBoxXYZ b, double tol)
        {
            return a.Min.X - tol <= b.Max.X && b.Min.X - tol <= a.Max.X
                && a.Min.Y - tol <= b.Max.Y && b.Min.Y - tol <= a.Max.Y
                && a.Min.Z - tol <= b.Max.Z && b.Min.Z - tol <= a.Max.Z;
        }

        /// <summary>
        /// 面 a が面 b によって「覆われている」（中心が面 b 内部にある）かを判定。
        /// </summary>
        private static bool IsFaceCovered(FaceClassifier.FaceInfo a, FaceClassifier.FaceInfo b)
        {
            if (a?.Face == null || b?.Face == null) return false;
            if (a.Normal == null || b.Normal == null) return false;

            // 条件 1: 反平行
            double dot = a.Normal.DotProduct(b.Normal);
            if (dot > AntiParallelThreshold) return false;

            // 条件 2: 面積比
            double aArea = 0, bArea = 0;
            try { aArea = a.Face.Area; bArea = b.Face.Area; } catch { }
            if (aArea <= 1e-6 || bArea <= 1e-6) return false;
            if (aArea > bArea * AreaRatioLimit) return false;

            // 面 a の中心点
            BoundingBoxUV bbA;
            try { bbA = a.Face.GetBoundingBox(); }
            catch { return false; }

            XYZ pA;
            try { pA = a.Face.Evaluate((bbA.Min + bbA.Max) * 0.5); }
            catch { return false; }
            if (pA == null) return false;

            // 条件 3: 面 a 中心点を面 b に投影し、距離が tol 以内
            IntersectionResult proj;
            try { proj = b.Face.Project(pA); }
            catch { return false; }
            if (proj == null) return false;
            if (proj.Distance > CoincidenceTolFeet) return false;

            // 条件 4: 投影先 UV が面 b の UV 範囲の内部にある（境界のすぐ外ではない）
            // Face.Project は「面上の最近点」を返すため、面 a 中心が面 b の外にある場合
            // UV が面 b の境界上に落ちる。境界から十分内側にあることを確認する。
            try
            {
                var uv = proj.UVPoint;
                if (uv == null) return false;
                BoundingBoxUV bbB;
                try { bbB = b.Face.GetBoundingBox(); }
                catch { return false; }

                double marginU = (bbB.Max.U - bbB.Min.U) * 0.01;  // 1% margin
                double marginV = (bbB.Max.V - bbB.Min.V) * 0.01;

                if (uv.U < bbB.Min.U - 1e-6) return false;
                if (uv.U > bbB.Max.U + 1e-6) return false;
                if (uv.V < bbB.Min.V - 1e-6) return false;
                if (uv.V > bbB.Max.V + 1e-6) return false;

                // 境界から少し内側か? (margin を使って boundary-only のプロジェクションを弾く)
                bool nearBoundary =
                    uv.U < bbB.Min.U + marginU ||
                    uv.U > bbB.Max.U - marginU ||
                    uv.V < bbB.Min.V + marginV ||
                    uv.V > bbB.Max.V - marginV;

                // 境界上でも、perpendicular distance が十分小さければ接触と見なす
                // (faces が同一境界上で密着しているケース)
                if (nearBoundary && proj.Distance > CoincidenceTolFeet * 0.5)
                    return false;
            }
            catch { return false; }

            return true;
        }
    }
}

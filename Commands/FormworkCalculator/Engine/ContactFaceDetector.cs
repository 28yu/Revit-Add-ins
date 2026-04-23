using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 要素同士の接触面（結合面）を検出して DeductedContact に変更する。
    ///
    /// 【判定ロジック】
    ///   1. BoundingBox が重なる要素ペアのみ対象（効率化）
    ///   2. 面の法線が反平行（互いに向き合う）
    ///   3. 面 a の「中心点」が面 b 上に (tolerance 以内) 投影される
    ///   4. 面 a の面積 <= 面 b の面積 × 1.5 (a が b より大きすぎない)
    ///
    /// この緩めた判定により以下のケースを拾う:
    ///   - 壁 T 字結合の端面（壁 B の端面 ⊂ 壁 A の側面）
    ///   - 梁と柱の取り合い（梁端面 ⊂ 柱側面）
    ///   - スラブ同士の結合（端面同士）
    ///
    /// 一方、壁 A の大きな側面は面積条件で弾かれ FormworkRequired のまま残る。
    /// </summary>
    internal class ContactFaceDetector
    {
        internal class ElementFacesContext
        {
            public int ElementId;
            public BoundingBoxXYZ BB;
            public List<FaceClassifier.FaceInfo> Faces = new List<FaceClassifier.FaceInfo>();
        }

        private const double CoincidenceTolFeet = 0.05;   // ~15mm 許容（モデリング誤差吸収）
        private const double AreaRatioLimit = 1.5;        // 面 a は面 b の 1.5 倍まで

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
                            if (IsContactFace(fi, fj))
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
        /// 面 a が面 b と接触しているかを判定。
        /// 面 a の中心点が面 b に投影できて、面積比も妥当な場合のみ true。
        /// </summary>
        private static bool IsContactFace(FaceClassifier.FaceInfo a, FaceClassifier.FaceInfo b)
        {
            if (a.Normal == null || b.Normal == null) return false;
            // 反平行（向き合っている）を要求。0.90 に緩めてコーナーでの角度ずれも許容
            if (a.Normal.DotProduct(b.Normal) > -0.90) return false;

            BoundingBoxUV bbA;
            try { bbA = a.Face.GetBoundingBox(); }
            catch { return false; }
            UV midA = (bbA.Min + bbA.Max) * 0.5;

            XYZ ptA;
            try { ptA = a.Face.Evaluate(midA); }
            catch { return false; }
            if (ptA == null) return false;

            // 面 a の中心点 ptA を面 b に投影し、距離が tolerance 以内なら接触
            try
            {
                var proj = b.Face.Project(ptA);
                if (proj == null) return false;
                if (proj.Distance > CoincidenceTolFeet) return false;
            }
            catch { return false; }

            // 面積比チェック: a が b より大きすぎる場合は「大きな面の中央が偶然小面に当たった」
            // ケースなので除外する（主壁の側面が小端面に誤って Contact 判定されるのを防ぐ）
            double aArea = 0, bArea = 0;
            try { aArea = a.Face.Area; bArea = b.Face.Area; } catch { }
            if (aArea > 0 && bArea > 0 && aArea > bArea * AreaRatioLimit) return false;

            return true;
        }
    }
}

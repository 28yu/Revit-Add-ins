using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 要素同士の接触面（結合面）を検出して DeductedContact に変更する。
    /// 壁が直交結合している箇所・スラブ同士が結合している箇所などを
    /// 型枠対象外とするために使用。
    ///
    /// アルゴリズム:
    ///   1. BoundingBox が重なる要素ペアのみ対象（効率化）
    ///   2. 面の法線が反平行（互いに向き合う）かチェック
    ///   3. 面a の UV サンプル点 5 個が全て 面b の上にある場合に
    ///      「完全に覆われている」として接触面と判定
    ///   4. 部分的に覆われているだけの大きな面はそのまま型枠必要
    ///      （T字結合の主壁側面など）
    /// </summary>
    internal class ContactFaceDetector
    {
        internal class ElementFacesContext
        {
            public int ElementId;
            public BoundingBoxXYZ BB;
            public List<FaceClassifier.FaceInfo> Faces = new List<FaceClassifier.FaceInfo>();
        }

        internal static void RefineContactFaces(
            List<ElementFacesContext> contexts,
            double coincidenceTolFeet = 0.02)
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
                    if (!BBoxesOverlap(ci.BB, cj.BB, coincidenceTolFeet)) continue;

                    foreach (var fi in ci.Faces)
                    {
                        if (fi.FaceType != FaceType.FormworkRequired) continue;

                        foreach (var fj in cj.Faces)
                        {
                            if (IsFaceFullyCovered(fi, fj, coincidenceTolFeet))
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
        /// 面 a が面 b によって完全に覆われているか（a の全域が b の面上にあるか）を判定。
        /// </summary>
        private static bool IsFaceFullyCovered(
            FaceClassifier.FaceInfo a,
            FaceClassifier.FaceInfo b,
            double tol)
        {
            if (a.Normal == null || b.Normal == null) return false;

            // 反平行（向き合っている）のみ対象
            if (a.Normal.DotProduct(b.Normal) > -0.99) return false;

            BoundingBoxUV bbA;
            try { bbA = a.Face.GetBoundingBox(); }
            catch { return false; }

            // 面 a の UV サンプル点 5 箇所（中央・四隅）
            var samples = new UV[]
            {
                (bbA.Min + bbA.Max) * 0.5,
                new UV(bbA.Min.U + (bbA.Max.U - bbA.Min.U) * 0.1,
                       bbA.Min.V + (bbA.Max.V - bbA.Min.V) * 0.1),
                new UV(bbA.Max.U - (bbA.Max.U - bbA.Min.U) * 0.1,
                       bbA.Min.V + (bbA.Max.V - bbA.Min.V) * 0.1),
                new UV(bbA.Min.U + (bbA.Max.U - bbA.Min.U) * 0.1,
                       bbA.Max.V - (bbA.Max.V - bbA.Min.V) * 0.1),
                new UV(bbA.Max.U - (bbA.Max.U - bbA.Min.U) * 0.1,
                       bbA.Max.V - (bbA.Max.V - bbA.Min.V) * 0.1),
            };

            foreach (var uv in samples)
            {
                XYZ pt;
                try { pt = a.Face.Evaluate(uv); }
                catch { return false; }
                if (pt == null) return false;

                try
                {
                    var proj = b.Face.Project(pt);
                    if (proj == null) return false;
                    if (proj.Distance > tol) return false;
                }
                catch { return false; }
            }
            return true;
        }
    }
}

using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 要素同士の接触面（結合面）を検出して DeductedContact に変更する。
    ///
    /// 【アプローチ】
    ///   1. 各 FormworkRequired 面の中心から "要素の内側" に小さくオフセットした点を origin とする
    ///   2. origin から法線方向（外向き）に ReferenceIntersector でレイを発射
    ///   3. 返されるヒットは Proximity 順なので:
    ///        a. 最初のヒットは自分自身の面 (origin は内側 → 法線方向へ進むと自面に当たる)
    ///        b. ownId フィルタで自分を除外
    ///        c. 次のヒットまでの「面の先の距離 (gap)」が閾値以内なら接触
    ///
    /// 【なぜ origin を内側にするか】
    ///   origin を面の外側に置くと、面同士が完全に密着している場合に
    ///   ReferenceIntersector が自面と隣面の両方を origin より前方と見なせず、
    ///   代わりに隣要素の「反対側の面」を最初にヒットしてしまう。
    ///   隣要素が厚いと Proximity が閾値を超えて接触なしと誤判定される。
    /// </summary>
    internal class ContactFaceDetector
    {
        internal class ElementFacesContext
        {
            public int ElementId;
            public BoundingBoxXYZ BB;
            public List<FaceClassifier.FaceInfo> Faces = new List<FaceClassifier.FaceInfo>();
        }

        // origin を面の内側に置くオフセット (要素内部側)
        private const double OriginOffsetInsideFeet = 0.033;  // ≈ 10mm
        // 自面通過後の gap 許容値 (隙間・モデリング誤差吸収)
        private const double MaxGapFeet = 0.16;                // ≈ 50mm

        internal static void RefineContactFaces(
            Document doc,
            View3D rayView,
            List<ElementFacesContext> contexts)
        {
            if (rayView == null) return;

            var bicList = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_Stairs,
            };
            var multiCatFilter = new ElementMulticategoryFilter(bicList);

            ReferenceIntersector ri;
            try
            {
                ri = new ReferenceIntersector(multiCatFilter, FindReferenceTarget.Face, rayView);
                ri.FindReferencesInRevitLinks = false;
            }
            catch
            {
                return;
            }

            foreach (var ctx in contexts)
            {
                var ownId = new ElementId(ctx.ElementId);

                foreach (var fi in ctx.Faces)
                {
                    if (fi.FaceType != FaceType.FormworkRequired) continue;
                    if (fi.Normal == null) continue;

                    BoundingBoxUV bbA;
                    try { bbA = fi.Face.GetBoundingBox(); }
                    catch { continue; }

                    XYZ pt;
                    try { pt = fi.Face.Evaluate((bbA.Min + bbA.Max) * 0.5); }
                    catch { continue; }
                    if (pt == null) continue;

                    // origin を面の内側（要素の中心側）に設定。法線方向に発射すれば
                    // 先に自面に当たり、さらに先に隣要素があればそれにも当たる。
                    XYZ origin = pt - fi.Normal * OriginOffsetInsideFeet;

                    bool isContact = false;
                    try
                    {
                        var hits = ri.Find(origin, fi.Normal);
                        if (hits != null)
                        {
                            foreach (var h in hits)
                            {
                                var hitId = h.GetReference()?.ElementId;
                                if (hitId == null || hitId == ownId) continue;

                                // 自面（Proximity ≈ OriginOffsetInsideFeet）を通過してから
                                // 隣要素までの gap を計算
                                double gap = h.Proximity - OriginOffsetInsideFeet;
                                if (gap > MaxGapFeet) break;

                                isContact = true;
                                break;
                            }
                        }
                    }
                    catch { }

                    if (isContact)
                        fi.FaceType = FaceType.DeductedContact;
                }
            }
        }
    }
}

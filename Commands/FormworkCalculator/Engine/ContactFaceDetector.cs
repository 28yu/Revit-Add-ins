using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 要素同士の接触面（結合面）を検出して DeductedContact に変更する。
    ///
    /// 【アプローチ】両方向レイキャスティング
    ///   - 面の中心点から、法線方向の両側 (+normal / -normal) にそれぞれレイを発射
    ///   - ReferenceIntersector で「自面の先に別要素の面が近接」しているかを確認
    ///   - どちらかの方向で接触が検出されたら DeductedContact に変更
    ///
    /// 【なぜ両方向か】
    ///   Boolean Union 後の Solid で、Face.ComputeNormal() が返す法線が
    ///   必ずしも "外向き" とは限らないため、normal 方向の仮定を外す。
    ///   接触は片側だけで発生するため、正しい側でのみヒットが発生する。
    /// </summary>
    internal class ContactFaceDetector
    {
        internal class ElementFacesContext
        {
            public int ElementId;
            public BoundingBoxXYZ BB;
            public List<FaceClassifier.FaceInfo> Faces = new List<FaceClassifier.FaceInfo>();
        }

        // 面の反対側に origin を置くオフセット（要素の内部側）
        private const double OriginOffsetInsideFeet = 0.033;  // ≈ 10mm
        // 自面通過後の gap 許容値
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

                    // 両方向にレイを発射（法線の向きの不確かさを吸収）
                    bool isContact =
                        CheckRayDirection(ri, pt, fi.Normal, ownId) ||
                        CheckRayDirection(ri, pt, -fi.Normal, ownId);

                    if (isContact)
                        fi.FaceType = FaceType.DeductedContact;
                }
            }
        }

        /// <summary>
        /// 指定方向にレイを発射して、自面通過後の近接ヒットがあれば接触と判定。
        /// origin は方向の反対側にオフセットを取る（自面が最初のヒットになるよう）。
        /// </summary>
        private static bool CheckRayDirection(
            ReferenceIntersector ri, XYZ pt, XYZ direction, ElementId ownId)
        {
            XYZ origin = pt - direction * OriginOffsetInsideFeet;

            try
            {
                var hits = ri.Find(origin, direction);
                if (hits == null) return false;

                foreach (var h in hits)
                {
                    var hitId = h.GetReference()?.ElementId;
                    if (hitId == null) continue;
                    if (hitId == ownId) continue;  // 自要素をスキップ

                    // 自面（Proximity ≈ OriginOffsetInsideFeet）を通過してから
                    // 隣要素までの gap を計算
                    double gap = h.Proximity - OriginOffsetInsideFeet;
                    if (gap > MaxGapFeet) return false;
                    if (gap < -MaxGapFeet) continue;  // 明らかに後方はスキップ

                    return true;
                }
            }
            catch { }

            return false;
        }
    }
}

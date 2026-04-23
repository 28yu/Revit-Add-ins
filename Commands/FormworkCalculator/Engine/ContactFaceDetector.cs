using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 要素同士の接触面（結合面）を検出して DeductedContact に変更する。
    ///
    /// アプローチ: ReferenceIntersector を使用
    ///   - 各 FormworkRequired 面の中央から外向き（法線方向）に極短いレイを発射
    ///   - 即座に他の構造要素にヒットした場合、接触面と判定
    ///   - これにより以下のケースが確実に検出される:
    ///       - 壁の T 字結合 (壁端面 → 別壁側面)
    ///       - 梁と柱の取り合い (梁端面 → 柱側面)
    ///       - 梁底と基礎/柱の接触
    ///       - スラブ同士の結合
    /// </summary>
    internal class ContactFaceDetector
    {
        internal class ElementFacesContext
        {
            public int ElementId;
            public BoundingBoxXYZ BB;
            public List<FaceClassifier.FaceInfo> Faces = new List<FaceClassifier.FaceInfo>();
        }

        // ヒットの近接距離: 50mm 以内なら接触とみなす（モデリング誤差・小隙間を吸収）
        private const double ProximityFeet = 0.16;     // ≒ 50mm
        // 面中央から発射するレイの始点オフセット (面の外側 1mm)
        private const double RayOriginOffsetFeet = 0.0033; // ≒ 1mm

        /// <summary>
        /// View3D を使って ReferenceIntersector で接触面検出を実行。
        /// View3D は ReferenceIntersector の必須引数。
        /// </summary>
        internal static void RefineContactFaces(
            Document doc,
            View3D rayView,
            List<ElementFacesContext> contexts)
        {
            if (rayView == null) return;

            // 検出対象は対象スコープの構造カテゴリ全般
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

                    // 面の中心点を取得
                    BoundingBoxUV bbA;
                    try { bbA = fi.Face.GetBoundingBox(); }
                    catch { continue; }

                    XYZ pt;
                    try { pt = fi.Face.Evaluate((bbA.Min + bbA.Max) * 0.5); }
                    catch { continue; }
                    if (pt == null) continue;

                    // 面の外側にわずかに出した点から、法線方向に短いレイを発射
                    XYZ origin = pt + fi.Normal * RayOriginOffsetFeet;

                    bool isContact = false;
                    try
                    {
                        var hits = ri.Find(origin, fi.Normal);
                        if (hits != null)
                        {
                            foreach (var h in hits)
                            {
                                // 自分自身のヒットは無視
                                var hitId = h.GetReference()?.ElementId;
                                if (hitId == null || hitId == ownId) continue;

                                // 近接距離チェック
                                if (h.Proximity > ProximityFeet) break;

                                // 他要素に近接ヒット = 接触
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

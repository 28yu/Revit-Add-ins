using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 下面の型枠要否判定（レイキャスト）。
    /// 下向き水平面の中央点から下方向へレイを飛ばし、他要素にヒットした場合は接触面として控除。
    /// </summary>
    internal static class BottomFaceAnalyzer
    {
        private const double RayTolerance = 0.01;

        internal static void RefineDownwardFaces(
            Document doc,
            View3D rayView,
            List<FaceClassifier.FaceInfo> faces,
            ICollection<ElementId> ownElementIds)
        {
            if (rayView == null) return;

            ReferenceIntersector ri;
            try
            {
                ri = new ReferenceIntersector(
                    new ElementCategoryFilter(BuiltInCategory.OST_Walls, false),
                    FindReferenceTarget.Face,
                    rayView);
                ri.FindReferencesInRevitLinks = false;
            }
            catch
            {
                return;
            }

            foreach (var info in faces)
            {
                if (info.FaceType != FaceType.FormworkRequired) continue;
                if (info.Normal == null || info.Normal.Z > -0.99) continue;

                BoundingBoxUV bb = info.Face.GetBoundingBox();
                UV mid = (bb.Min + bb.Max) * 0.5;
                XYZ pt;
                try { pt = info.Face.Evaluate(mid); } catch { continue; }
                if (pt == null) continue;

                // 少し下からレイを飛ばす
                XYZ origin = new XYZ(pt.X, pt.Y, pt.Z - RayTolerance);
                XYZ dir = new XYZ(0, 0, -1);

                try
                {
                    var hit = ri.FindNearest(origin, dir);
                    if (hit != null)
                    {
                        info.FaceType = FaceType.DeductedContact;
                    }
                }
                catch { }
            }
        }
    }
}

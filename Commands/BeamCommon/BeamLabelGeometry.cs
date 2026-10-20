using System;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamCommon
{
    /// <summary>
    /// 梁ラベル（文字・梁タグ）を梁の脇に置くための寸法計算
    /// </summary>
    public static class BeamLabelGeometry
    {
        // 梁幅が取れない・細すぎる梁でもラベルが梁に重ならないようにする最小幅（200mm）
        private const double MinBeamWidthFt = 200.0 / 304.8;

        private static readonly string[] WidthParamNames =
            { "b", "B", "幅", "梁幅", "W", "w", "Width", "width" };

        /// <summary>
        /// 梁幅を取得する。インスタンス／タイプのパラメータを優先し、無ければビューでの外形から求める
        /// </summary>
        /// <param name="direction">梁の向き（XY 平面、正規化済み）</param>
        public static double GetBeamWidth(FamilyInstance beam, View view, XYZ direction)
        {
            double width = GetBeamWidthFromParameters(beam);

            if (width <= 0)
            {
                BoundingBoxXYZ bb = beam.get_BoundingBox(view);
                if (bb != null)
                {
                    XYZ normal = new XYZ(-direction.Y, direction.X, 0);
                    XYZ[] corners =
                    {
                        new XYZ(bb.Min.X, bb.Min.Y, 0),
                        new XYZ(bb.Max.X, bb.Min.Y, 0),
                        new XYZ(bb.Min.X, bb.Max.Y, 0),
                        new XYZ(bb.Max.X, bb.Max.Y, 0)
                    };
                    double minProj = double.MaxValue;
                    double maxProj = double.MinValue;
                    foreach (var corner in corners)
                    {
                        double proj = corner.X * normal.X + corner.Y * normal.Y;
                        minProj = Math.Min(minProj, proj);
                        maxProj = Math.Max(maxProj, proj);
                    }
                    width = maxProj - minProj;
                }
            }

            return Math.Max(width, MinBeamWidthFt);
        }

        private static double GetBeamWidthFromParameters(FamilyInstance beam)
        {
            foreach (string name in WidthParamNames)
            {
                Parameter param = beam.LookupParameter(name);
                if (param != null && param.StorageType == StorageType.Double)
                {
                    double val = param.AsDouble();
                    if (val > 0) return val;
                }
            }

            ElementType beamType = beam.Symbol;
            if (beamType != null)
            {
                foreach (string name in WidthParamNames)
                {
                    Parameter param = beamType.LookupParameter(name);
                    if (param != null && param.StorageType == StorageType.Double)
                    {
                        double val = param.AsDouble();
                        if (val > 0) return val;
                    }
                }
            }

            return 0;
        }
    }
}

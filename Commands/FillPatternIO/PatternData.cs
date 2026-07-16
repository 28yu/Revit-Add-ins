using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FillPatternIO
{
    /// <summary>
    /// FillPattern から描画に必要な数値だけを取り出したプレーンデータ。
    /// Revit API はスレッド安全ではないため、抽出は必ず Revit（UI）スレッドで行い、
    /// 実際の描画はこのデータを使ってバックグラウンドスレッドで行う。
    /// 長さの単位は Revit 内部単位（フィート）。
    /// </summary>
    internal sealed class PatternGridData
    {
        public double Angle;
        public double OriginU;
        public double OriginV;
        public double Offset;
        public double Shift;
        public double[] Segments;
    }

    internal sealed class PatternData
    {
        public bool IsSolid;
        public readonly List<PatternGridData> Grids = new List<PatternGridData>();

        /// <summary>Revit（UI）スレッドで呼び出すこと。</summary>
        public static PatternData From(FillPattern fp)
        {
            var data = new PatternData { IsSolid = fp == null || fp.IsSolidFill };
            if (fp == null || fp.IsSolidFill) return data;

            foreach (FillGrid g in fp.GetFillGrids())
            {
                var segs = g.GetSegments();
                data.Grids.Add(new PatternGridData
                {
                    Angle = g.Angle,
                    OriginU = g.Origin.U,
                    OriginV = g.Origin.V,
                    Offset = g.Offset,
                    Shift = g.Shift,
                    Segments = segs != null ? segs.ToArray() : new double[0]
                });
            }
            return data;
        }
    }
}

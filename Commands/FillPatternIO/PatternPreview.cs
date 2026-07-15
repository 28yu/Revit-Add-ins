using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
// System.Windows.Point と Autodesk.Revit.DB.Point の衝突を回避
using Point = System.Windows.Point;

namespace Tools28.Commands.FillPatternIO
{
    /// <summary>
    /// FillPattern を一覧表示用のプレビュー画像（ハッチの見た目）に描画するユーティリティ。
    /// Revit の「塗り潰しパターン」ダイアログのように、一目で判別できるようにする。
    /// </summary>
    internal static class PatternPreview
    {
        private const double Eps = 1e-9;
        private const double MmToFeet = 1.0 / 304.8;

        public static ImageSource Render(FillPattern fp, int pxW, int pxH)
        {
            var dv = new DrawingVisual();
            using (DrawingContext dc = dv.RenderOpen())
            {
                var full = new Rect(0, 0, pxW, pxH);
                dc.DrawRectangle(Brushes.White, null, full);
                dc.PushClip(new RectangleGeometry(full));

                try
                {
                    if (fp == null || fp.IsSolidFill)
                    {
                        dc.DrawRectangle(Brushes.Black, null, full);
                    }
                    else
                    {
                        var grids = fp.GetFillGrids();
                        if (grids != null && grids.Count > 0)
                            DrawGrids(dc, grids, pxW, pxH);
                    }
                }
                catch
                {
                    // 描画に失敗しても空のプレビューを返す
                }

                dc.Pop();
            }

            var rtb = new RenderTargetBitmap(pxW, pxH, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(dv);
            rtb.Freeze();
            return rtb;
        }

        private static void DrawGrids(DrawingContext dc, IList<FillGrid> grids, int pxW, int pxH)
        {
            // 代表間隔（最小の正のオフセット）を基準に、高さ方向へ約6本収まるスケールにする
            double minOffset = grids
                .Select(g => Math.Abs(g.Offset))
                .Where(o => o > Eps)
                .DefaultIfEmpty(MmToFeet)
                .Min();

            double modelH = minOffset * 6.0;
            if (modelH < Eps) modelH = MmToFeet;

            double scale = pxH / modelH;
            double modelW = pxW / scale;

            // 表示窓の中心をグリッド原点の平均に合わせる
            double cU = grids.Average(g => g.Origin.U);
            double cV = grids.Average(g => g.Origin.V);
            double winMinX = cU - modelW / 2.0;
            double winMinY = cV - modelH / 2.0;

            double diag = Math.Sqrt(modelW * modelW + modelH * modelH);

            var pen = new Pen(Brushes.Black, 1.0);
            pen.Freeze();

            Point ToPx(double mx, double my)
                => new Point((mx - winMinX) * scale, pxH - (my - winMinY) * scale);

            foreach (var g in grids)
            {
                double offset = Math.Abs(g.Offset);
                if (offset < Eps) continue;

                double dx = Math.Cos(g.Angle), dy = Math.Sin(g.Angle);
                double perpX = -dy, perpY = dx;

                var segs = g.GetSegments();
                double repeat = (segs != null && segs.Count > 0)
                    ? segs.Sum(s => Math.Abs(s))
                    : 0.0;

                int nRange = (int)Math.Ceiling(diag / offset) + 1;
                if (nRange > 4000) nRange = 4000;

                // ピクセル間隔が細かすぎる場合は間引く
                double pxStep = offset * scale;
                int nStep = pxStep < 1.0 ? (int)Math.Ceiling(1.0 / pxStep) : 1;

                for (int n = -nRange; n <= nRange; n += nStep)
                {
                    double bx = g.Origin.U + n * offset * perpX;
                    double by = g.Origin.V + n * offset * perpY;

                    if (repeat < Eps)
                    {
                        dc.DrawLine(pen,
                            ToPx(bx - diag * dx, by - diag * dy),
                            ToPx(bx + diag * dx, by + diag * dy));
                        continue;
                    }

                    // 破線（ダッシュ）パターン
                    double phase = (n * g.Shift) % repeat;
                    if (phase < 0) phase += repeat;

                    double u = phase - repeat * (Math.Floor((phase + diag) / repeat) + 1);
                    int guard = 0;
                    while (u < diag && guard++ < 100000)
                    {
                        foreach (var s in segs)
                        {
                            double len = Math.Abs(s);
                            if (len < Eps) { continue; }

                            if (s > 0)
                            {
                                double aa = Math.Max(u, -diag);
                                double bb = Math.Min(u + len, diag);
                                if (bb > aa)
                                {
                                    dc.DrawLine(pen,
                                        ToPx(bx + aa * dx, by + aa * dy),
                                        ToPx(bx + bb * dx, by + bb * dy));
                                }
                            }

                            u += len;
                            if (u >= diag) break;
                        }
                    }
                }
            }
        }
    }
}

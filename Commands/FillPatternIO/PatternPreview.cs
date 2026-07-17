using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Tools28.Commands.FillPatternIO
{
    /// <summary>
    /// PatternData（Revit 非依存のプレーンデータ）をハッチのプレビュー画像へ描画する。
    /// Revit API を参照しないため、バックグラウンドスレッドから呼び出しても安全。
    ///
    /// 表示は「固定の実寸ウィンドウ」で行う（Revit の塗り潰しパターンダイアログと同様）。
    /// これにより、間隔の細かいパターンは密に、粗いパターンは疎に、忠実に表示される。
    /// </summary>
    internal static class PatternPreview
    {
        private const double Eps = 1e-9;
        private const double MmPerFoot = 304.8;

        // プレビューの高さ(px)が表す実寸(mm)。小さいほど拡大表示（本数が減る）。
        private const double PreviewHeightMm = 12.0;

        // 1グリッドあたりの最大描画本数（これを超える＝サブピクセル間隔なので間引いても見た目は不変）
        private const int MaxLinesPerGrid = 400;

        public static ImageSource Render(PatternData data, int pxW, int pxH)
        {
            var dv = new DrawingVisual();
            using (DrawingContext dc = dv.RenderOpen())
            {
                var full = new Rect(0, 0, pxW, pxH);
                dc.DrawRectangle(Brushes.White, null, full);
                dc.PushClip(new RectangleGeometry(full));

                try
                {
                    if (data == null || data.IsSolid)
                        dc.DrawRectangle(Brushes.Black, null, full);
                    else if (data.Grids.Count > 0)
                        DrawGrids(dc, data.Grids, pxW, pxH, data.IsModel);
                }
                catch
                {
                    // 描画失敗時は空のプレビュー
                }

                dc.Pop();
            }

            var rtb = new RenderTargetBitmap(pxW, pxH, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(dv);
            rtb.Freeze();
            return rtb;
        }

        private static void DrawGrids(DrawingContext dc, List<PatternGridData> grids, int pxW, int pxH, bool isModel)
        {
            // 製図パターンは紙面mm基準の固定ウィンドウ（12mm）で忠実に表示する。
            // モデルパターンは実寸（数十〜数百mm）のため固定窓に収まらない。
            // その場合はパターン自身の間隔に合わせてズームアウトする（尺度がおかしくならないように）。
            double fixedH = PreviewHeightMm / MmPerFoot;       // フィート
            double modelH = fixedH;
            if (isModel)
            {
                double minOffset = grids
                    .Select(g => Math.Abs(g.Offset))
                    .Where(o => o > Eps)
                    .DefaultIfEmpty(fixedH)
                    .Min();
                // 最も細かいグリッドが縦に約6本収まる高さ。固定窓より拡大はしない。
                modelH = Math.Max(fixedH, minOffset * 6.0);
            }

            double scale = pxH / modelH;                       // px / フィート
            double modelW = pxW / scale;
            double halfW = modelW / 2.0;
            double halfH = modelH / 2.0;

            var pen = new Pen(Brushes.Black, 0.5)
            {
                StartLineCap = PenLineCap.Round,
                EndLineCap = PenLineCap.Round
            };
            pen.Freeze();

            const double dotRadius = 0.6;

            // モデル座標 → ピクセル座標（ウィンドウ中心を原点(0,0)とする）
            Point ToPx(double mx, double my)
                => new Point((mx + halfW) * scale, pxH - (my + halfH) * scale);

            // ウィンドウ4隅
            var corners = new[]
            {
                new[] { -halfW, -halfH }, new[] { halfW, -halfH },
                new[] { -halfW,  halfH }, new[] { halfW,  halfH }
            };

            foreach (var g in grids)
            {
                double offset = Math.Abs(g.Offset);
                if (offset < Eps) continue;

                double dx = Math.Cos(g.Angle), dy = Math.Sin(g.Angle);
                double px = -dy, py = dx; // 線に垂直な方向

                // 各線は O + n*offset*perp を通る。ウィンドウを覆う n の範囲を求める
                double pMin = double.MaxValue, pMax = double.MinValue;
                foreach (var c in corners)
                {
                    double proj = (c[0] - g.OriginU) * px + (c[1] - g.OriginV) * py;
                    if (proj < pMin) pMin = proj;
                    if (proj > pMax) pMax = proj;
                }
                int nMin = (int)Math.Floor(pMin / offset) - 1;
                int nMax = (int)Math.Ceiling(pMax / offset) + 1;

                int count = nMax - nMin + 1;
                if (count < 1) continue;
                int step = count > MaxLinesPerGrid ? (int)Math.Ceiling((double)count / MaxLinesPerGrid) : 1;

                double[] segs = g.Segments ?? new double[0];
                double repeat = segs.Length > 0 ? segs.Sum(v => Math.Abs(v)) : 0.0;

                // ダッシュ周期がサブピクセル(0.5px未満)なら実線と見分けが付かないので実線化
                bool asSolid = repeat < Eps || (repeat * scale) < 0.5;

                for (int n = nMin; n <= nMax; n += step)
                {
                    double bx = g.OriginU + n * offset * px;
                    double by = g.OriginV + n * offset * py;

                    // この線がウィンドウを横切る区間 [uStart, uEnd] を厳密に求める
                    // （4隅を線方向 d に射影）。無駄な描画を避けられる。
                    double uStart = double.MaxValue, uEnd = double.MinValue;
                    foreach (var c in corners)
                    {
                        double tproj = (c[0] - bx) * dx + (c[1] - by) * dy;
                        if (tproj < uStart) uStart = tproj;
                        if (tproj > uEnd) uEnd = tproj;
                    }

                    if (asSolid)
                    {
                        dc.DrawLine(pen,
                            ToPx(bx + uStart * dx, by + uStart * dy),
                            ToPx(bx + uEnd * dx, by + uEnd * dy));
                        continue;
                    }

                    double phase = (n * g.Shift) % repeat;
                    if (phase < 0) phase += repeat;
                    double u = phase + repeat * Math.Floor((uStart - phase) / repeat);

                    int guard = 0;
                    while (u < uEnd && guard++ < 20000)
                    {
                        // Revit のセグメントは「偶数index=ダッシュ、奇数index=ギャップ」で交互
                        for (int i = 0; i < segs.Length; i++)
                        {
                            double len = Math.Abs(segs[i]);
                            bool penDown = (i % 2 == 0);
                            if (penDown)
                            {
                                if (len < Eps)
                                {
                                    // 長さ0のダッシュ = 点（ドット）
                                    if (u >= uStart && u <= uEnd)
                                        dc.DrawEllipse(Brushes.Black, null,
                                            ToPx(bx + u * dx, by + u * dy), dotRadius, dotRadius);
                                }
                                else
                                {
                                    double aa = Math.Max(u, uStart);
                                    double bb = Math.Min(u + len, uEnd);
                                    if (bb > aa)
                                        dc.DrawLine(pen,
                                            ToPx(bx + aa * dx, by + aa * dy),
                                            ToPx(bx + bb * dx, by + bb * dy));
                                }
                            }
                            u += len;
                            if (u >= uEnd) break;
                        }
                    }
                }
            }
        }
    }
}

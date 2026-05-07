using System;
using System.IO;
using ClosedXML.Excel;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Output
{
    internal static class ExcelExporter
    {
        private const string HeaderColor = "9BBB59";

        internal static void Export(string path, FormworkSettings settings, FormworkResult result)
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var wb = new XLWorkbook())
            {
                WriteSummary(wb, settings, result);
                WriteByCategory(wb, result);

                if (settings.GroupByZone)
                    WriteByZone(wb, result);

                if (settings.GroupByFormworkType)
                    WriteByType(wb, result);

                WriteElementDetail(wb, result);

                if (result.Errors.Count > 0)
                    WriteErrors(wb, result);

                wb.SaveAs(path);
            }
        }

        private static void FormatHeader(IXLRow row)
        {
            row.Style.Fill.BackgroundColor = XLColor.FromHtml("#" + HeaderColor);
            row.Style.Font.FontColor = XLColor.White;
            row.Style.Font.Bold = true;
            row.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        private static void WriteSummary(XLWorkbook wb, FormworkSettings settings, FormworkResult result)
        {
            var ws = wb.Worksheets.Add("総括表");

            ws.Cell(1, 1).Value = "型枠数量集計 総括表";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Range(1, 1, 1, 4).Merge();

            ws.Cell(3, 1).Value = "計算範囲";
            ws.Cell(3, 2).Value = settings.Scope == CalculationScope.EntireProject
                ? "プロジェクト全体" : "現在のビュー";
            ws.Cell(4, 1).Value = "対象要素数";
            ws.Cell(4, 2).Value = result.ProcessedElementCount;
            ws.Cell(5, 1).Value = "実行日時";
            ws.Cell(5, 2).Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

            int r = 7;
            ws.Cell(r, 1).Value = "項目";
            ws.Cell(r, 2).Value = "面積 (㎡)";
            FormatHeader(ws.Row(r));

            r++;
            ws.Cell(r++, 1).Value = "型枠面積（合計）"; ws.Cell(r - 1, 2).Value = Round(result.TotalFormworkArea);
            ws.Cell(r++, 1).Value = "控除面積（合計）"; ws.Cell(r - 1, 2).Value = Round(result.TotalDeductedArea);
            ws.Cell(r++, 1).Value = "傾斜面（計算対象外）"; ws.Cell(r - 1, 2).Value = Round(result.InclinedFaceArea);

            AutoFitColumns(ws);
        }

        private static void WriteByCategory(XLWorkbook wb, FormworkResult result)
        {
            var ws = wb.Worksheets.Add("部位別");
            ws.Cell(1, 1).Value = "部位";
            ws.Cell(1, 2).Value = "要素数";
            ws.Cell(1, 3).Value = "型枠面積 (㎡)";
            ws.Cell(1, 4).Value = "控除面積 (㎡)";
            FormatHeader(ws.Row(1));

            int r = 2;
            foreach (var c in result.CategoryResults)
            {
                ws.Cell(r, 1).Value = CategoryLabel(c.Category);
                ws.Cell(r, 2).Value = c.ElementCount;
                ws.Cell(r, 3).Value = Round(c.FormworkArea);
                ws.Cell(r, 4).Value = Round(c.DeductedArea);
                r++;
            }
            AutoFitColumns(ws);
        }

        private static void WriteByZone(XLWorkbook wb, FormworkResult result)
        {
            var ws = wb.Worksheets.Add("工区別");
            ws.Cell(1, 1).Value = "工区";
            ws.Cell(1, 2).Value = "要素数";
            ws.Cell(1, 3).Value = "型枠面積 (㎡)";
            FormatHeader(ws.Row(1));

            int r = 2;
            foreach (var z in result.ZoneResults)
            {
                ws.Cell(r, 1).Value = z.Zone;
                ws.Cell(r, 2).Value = z.ElementCount;
                ws.Cell(r, 3).Value = Round(z.FormworkArea);
                r++;
            }
            AutoFitColumns(ws);
        }

        private static void WriteByType(XLWorkbook wb, FormworkResult result)
        {
            var ws = wb.Worksheets.Add("型枠種別");
            ws.Cell(1, 1).Value = "型枠種別";
            ws.Cell(1, 2).Value = "要素数";
            ws.Cell(1, 3).Value = "型枠面積 (㎡)";
            FormatHeader(ws.Row(1));

            int r = 2;
            foreach (var t in result.TypeResults)
            {
                ws.Cell(r, 1).Value = t.FormworkType;
                ws.Cell(r, 2).Value = t.ElementCount;
                ws.Cell(r, 3).Value = Round(t.FormworkArea);
                r++;
            }
            AutoFitColumns(ws);
        }

        private static void WriteElementDetail(XLWorkbook wb, FormworkResult result)
        {
            var ws = wb.Worksheets.Add("要素明細");
            ws.Cell(1, 1).Value = "要素ID";
            ws.Cell(1, 2).Value = "部位";
            ws.Cell(1, 3).Value = "要素名";
            ws.Cell(1, 4).Value = "工区";
            ws.Cell(1, 5).Value = "型枠種別";
            ws.Cell(1, 6).Value = "型枠面積 (㎡)";
            ws.Cell(1, 7).Value = "天端控除";
            ws.Cell(1, 8).Value = "底面控除";
            ws.Cell(1, 9).Value = "接触控除";
            ws.Cell(1, 10).Value = "傾斜面";
            ws.Cell(1, 11).Value = "開口控除";
            ws.Cell(1, 12).Value = "開口見込み加算";
            FormatHeader(ws.Row(1));

            int r = 2;
            foreach (var er in result.ElementResults)
            {
                ws.Cell(r, 1).Value = er.ElementId;
                ws.Cell(r, 2).Value = CategoryLabel(er.Category);
                ws.Cell(r, 3).Value = er.ElementName;
                ws.Cell(r, 4).Value = er.Zone;
                ws.Cell(r, 5).Value = er.FormworkType;
                ws.Cell(r, 6).Value = Round(er.FormworkArea);
                ws.Cell(r, 7).Value = Round(er.DeductedTopArea);
                ws.Cell(r, 8).Value = Round(er.DeductedBottomArea);
                ws.Cell(r, 9).Value = Round(er.DeductedContactArea);
                ws.Cell(r, 10).Value = Round(er.InclinedArea);
                ws.Cell(r, 11).Value = Round(er.OpeningAreaDeducted);
                ws.Cell(r, 12).Value = Round(er.OpeningEdgeAreaAdded);
                r++;
            }
            ws.SheetView.FreezeRows(1);
            ws.RangeUsed()?.SetAutoFilter();
            AutoFitColumns(ws);
        }

        private static void WriteErrors(XLWorkbook wb, FormworkResult result)
        {
            var ws = wb.Worksheets.Add("エラー・注記");
            ws.Cell(1, 1).Value = "要素ID";
            ws.Cell(1, 2).Value = "カテゴリ";
            ws.Cell(1, 3).Value = "要素名";
            ws.Cell(1, 4).Value = "種別";
            ws.Cell(1, 5).Value = "メッセージ";
            FormatHeader(ws.Row(1));

            int r = 2;
            foreach (var e in result.Errors)
            {
                ws.Cell(r, 1).Value = e.ElementId;
                ws.Cell(r, 2).Value = e.CategoryName;
                ws.Cell(r, 3).Value = e.ElementName;
                ws.Cell(r, 4).Value = e.ErrorKind;
                ws.Cell(r, 5).Value = e.Message;
                r++;
            }
            AutoFitColumns(ws);
        }

        private static string CategoryLabel(CategoryGroup cg)
        {
            switch (cg)
            {
                case CategoryGroup.Column: return "柱";
                case CategoryGroup.Beam: return "梁";
                case CategoryGroup.Wall: return "壁";
                case CategoryGroup.Slab: return "スラブ";
                case CategoryGroup.Foundation: return "基礎";
                case CategoryGroup.Stairs: return "階段";
                default: return "その他";
            }
        }

        private static double Round(double v) => Math.Round(v, 2);

        /// <summary>
        /// セル幅を内容に合わせて自動調整する。CJK 文字 (漢字・かな) は半角の 2 倍幅で
        /// 計算するため ClosedXML の AdjustToContents() より日本語が見切れにくい。
        /// </summary>
        private static void AutoFitColumns(IXLWorksheet ws, double padding = 2.0, double minWidth = 8.0, double maxWidth = 60.0)
        {
            var range = ws.RangeUsed();
            if (range == null) return;
            int firstCol = range.FirstColumn().ColumnNumber();
            int lastCol = range.LastColumn().ColumnNumber();
            int firstRow = range.FirstRow().RowNumber();
            int lastRow = range.LastRow().RowNumber();

            for (int col = firstCol; col <= lastCol; col++)
            {
                double maxLen = 0;
                for (int row = firstRow; row <= lastRow; row++)
                {
                    var cell = ws.Cell(row, col);
                    string s;
                    try { s = cell.GetFormattedString(); }
                    catch { s = cell.Value?.ToString() ?? string.Empty; }
                    if (string.IsNullOrEmpty(s)) continue;

                    // セル内改行は最長行で計測
                    foreach (var line in s.Split('\n'))
                    {
                        double w = MeasureWidth(line);
                        if (w > maxLen) maxLen = w;
                    }
                }
                double width = Math.Max(minWidth, Math.Min(maxWidth, maxLen + padding));
                try { ws.Column(col).Width = width; } catch { }
            }
        }

        /// <summary>
        /// 文字列の表示幅を「半角=1 / 全角=2」基準で計算する。
        /// </summary>
        private static double MeasureWidth(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            double w = 0;
            foreach (char c in s)
            {
                w += IsFullWidth(c) ? 2.0 : 1.1;
            }
            return w;
        }

        private static bool IsFullWidth(char c)
        {
            // CJK・かな・ハングル・全角記号などを「全角」として扱う
            if (c >= 0x1100 && c <= 0x115F) return true; // Hangul Jamo
            if (c >= 0x2E80 && c <= 0x303E) return true; // CJK Radicals / 句読点
            if (c >= 0x3041 && c <= 0x33FF) return true; // ひらがな・カタカナ・CJK 記号
            if (c >= 0x3400 && c <= 0x4DBF) return true; // CJK 拡張 A
            if (c >= 0x4E00 && c <= 0x9FFF) return true; // CJK 統合漢字
            if (c >= 0xA000 && c <= 0xA4CF) return true; // ヤオ文字
            if (c >= 0xAC00 && c <= 0xD7A3) return true; // ハングル音節
            if (c >= 0xF900 && c <= 0xFAFF) return true; // CJK 互換漢字
            if (c >= 0xFE30 && c <= 0xFE4F) return true; // CJK 互換形
            if (c >= 0xFF00 && c <= 0xFF60) return true; // 全角英数・記号
            if (c >= 0xFFE0 && c <= 0xFFE6) return true; // 全角通貨記号
            return false;
        }
    }
}

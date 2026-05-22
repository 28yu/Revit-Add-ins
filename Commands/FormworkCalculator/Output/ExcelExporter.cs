using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Output
{
    internal static class ExcelExporter
    {
        private const string HeaderColor = "9BBB59";

        /// <summary>
        /// 単一ソースビュー結果を Excel に出力する (後方互換)。
        /// </summary>
        internal static void Export(string path, FormworkSettings settings, FormworkResult result)
        {
            ExportMulti(path, settings,
                new List<FormworkResult> { result },
                new List<string> { string.Empty });
        }

        /// <summary>
        /// 複数ソースビューの結果を 1 つの Excel に集約出力する。
        /// 集計表（合計シート）と同じ「全ビュー作成済み DirectShape 合算」値を表示する。
        /// </summary>
        internal static void ExportMulti(
            string path, FormworkSettings settings,
            IList<FormworkResult> results, IList<string> sourceViewNames)
        {
            if (results == null || results.Count == 0) return;
            if (sourceViewNames == null || sourceViewNames.Count != results.Count)
            {
                sourceViewNames = Enumerable.Range(0, results.Count)
                    .Select(_ => string.Empty).ToList();
            }

            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var wb = new XLWorkbook())
            {
                WriteSummary(wb, settings, results, sourceViewNames);
                WriteByCategoryAggregated(wb, results);

                if (settings.GroupByZone)
                    WriteByZoneAggregated(wb, results);

                if (settings.GroupByFormworkType)
                    WriteByTypeAggregated(wb, results);

                WriteElementDetailMulti(wb, results, sourceViewNames);

                if (results.Any(r => r.Errors != null && r.Errors.Count > 0))
                    WriteErrorsMulti(wb, results, sourceViewNames);

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

        private static void WriteSummary(XLWorkbook wb, FormworkSettings settings,
            IList<FormworkResult> results, IList<string> sourceViewNames)
        {
            var ws = wb.Worksheets.Add("総括表");

            ws.Cell(1, 1).Value = "型枠数量集計 総括表";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Range(1, 1, 1, 4).Merge();

            ws.Cell(3, 1).Value = "計算範囲";
            ws.Cell(3, 2).Value = settings.Scope == CalculationScope.EntireProject
                ? "プロジェクト全体"
                : (settings.Scope == CalculationScope.SelectedViews
                    ? $"選択ビュー ({results.Count}件)"
                    : "現在のビュー");
            ws.Cell(4, 1).Value = "対象要素数";
            ws.Cell(4, 2).Value = results.Sum(r => r.ProcessedElementCount);
            ws.Cell(5, 1).Value = "実行日時";
            ws.Cell(5, 2).Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

            int r = 7;
            ws.Cell(r, 1).Value = "項目";
            ws.Cell(r, 2).Value = "面積 (㎡)";
            FormatHeader(ws.Row(r));

            double totalFormwork = results.Sum(x => x.TotalFormworkArea);
            double totalDeducted = results.Sum(x => x.TotalDeductedArea);
            double totalInclined = results.Sum(x => x.InclinedFaceArea);

            r++;
            ws.Cell(r++, 1).Value = "型枠面積（合計）"; ws.Cell(r - 1, 2).Value = Round(totalFormwork);
            ws.Cell(r++, 1).Value = "控除面積（合計）"; ws.Cell(r - 1, 2).Value = Round(totalDeducted);
            ws.Cell(r++, 1).Value = "傾斜面（計算対象外）"; ws.Cell(r - 1, 2).Value = Round(totalInclined);

            // 複数ビュー時はソースビュー別の内訳も追加
            if (results.Count > 1)
            {
                r += 2;
                ws.Cell(r, 1).Value = "ソースビュー別 内訳";
                ws.Cell(r, 1).Style.Font.Bold = true;
                r++;
                ws.Cell(r, 1).Value = "ソースビュー";
                ws.Cell(r, 2).Value = "型枠面積 (㎡)";
                ws.Cell(r, 3).Value = "控除面積 (㎡)";
                ws.Cell(r, 4).Value = "要素数";
                FormatHeader(ws.Row(r));
                r++;
                for (int i = 0; i < results.Count; i++)
                {
                    ws.Cell(r, 1).Value = string.IsNullOrEmpty(sourceViewNames[i])
                        ? $"(view {i + 1})" : sourceViewNames[i];
                    ws.Cell(r, 2).Value = Round(results[i].TotalFormworkArea);
                    ws.Cell(r, 3).Value = Round(results[i].TotalDeductedArea);
                    ws.Cell(r, 4).Value = results[i].ProcessedElementCount;
                    r++;
                }
                ws.Cell(r, 1).Value = "合計";
                ws.Cell(r, 1).Style.Font.Bold = true;
                ws.Cell(r, 2).Value = Round(totalFormwork);
                ws.Cell(r, 2).Style.Font.Bold = true;
                ws.Cell(r, 3).Value = Round(totalDeducted);
                ws.Cell(r, 3).Style.Font.Bold = true;
                ws.Cell(r, 4).Value = results.Sum(x => x.ProcessedElementCount);
                ws.Cell(r, 4).Style.Font.Bold = true;
            }

            AutoFitColumns(ws);
        }

        private static void WriteByCategoryAggregated(XLWorkbook wb, IList<FormworkResult> results)
        {
            var ws = wb.Worksheets.Add("部位別");
            ws.Cell(1, 1).Value = "部位";
            ws.Cell(1, 2).Value = "要素数";
            ws.Cell(1, 3).Value = "型枠面積 (㎡)";
            ws.Cell(1, 4).Value = "控除面積 (㎡)";
            FormatHeader(ws.Row(1));

            var agg = new Dictionary<CategoryGroup, CategoryResult>();
            foreach (var res in results)
            {
                if (res.CategoryResults == null) continue;
                foreach (var c in res.CategoryResults)
                {
                    if (!agg.TryGetValue(c.Category, out var cur))
                    {
                        cur = new CategoryResult { Category = c.Category, CategoryName = c.CategoryName };
                        agg[c.Category] = cur;
                    }
                    cur.ElementCount += c.ElementCount;
                    cur.FormworkArea += c.FormworkArea;
                    cur.DeductedArea += c.DeductedArea;
                }
            }

            int r = 2;
            foreach (var c in agg.Values.OrderBy(x => x.Category))
            {
                ws.Cell(r, 1).Value = CategoryLabel(c.Category);
                ws.Cell(r, 2).Value = c.ElementCount;
                ws.Cell(r, 3).Value = Round(c.FormworkArea);
                ws.Cell(r, 4).Value = Round(c.DeductedArea);
                r++;
            }
            // 合計行
            ws.Cell(r, 1).Value = "合計";
            ws.Cell(r, 1).Style.Font.Bold = true;
            ws.Cell(r, 2).Value = agg.Values.Sum(x => x.ElementCount);
            ws.Cell(r, 2).Style.Font.Bold = true;
            ws.Cell(r, 3).Value = Round(agg.Values.Sum(x => x.FormworkArea));
            ws.Cell(r, 3).Style.Font.Bold = true;
            ws.Cell(r, 4).Value = Round(agg.Values.Sum(x => x.DeductedArea));
            ws.Cell(r, 4).Style.Font.Bold = true;
            AutoFitColumns(ws);
        }

        private static void WriteByZoneAggregated(XLWorkbook wb, IList<FormworkResult> results)
        {
            var ws = wb.Worksheets.Add("工区別");
            ws.Cell(1, 1).Value = "工区";
            ws.Cell(1, 2).Value = "要素数";
            ws.Cell(1, 3).Value = "型枠面積 (㎡)";
            FormatHeader(ws.Row(1));

            var agg = new Dictionary<string, ZoneResult>();
            foreach (var res in results)
            {
                if (res.ZoneResults == null) continue;
                foreach (var z in res.ZoneResults)
                {
                    string key = z.Zone ?? string.Empty;
                    if (!agg.TryGetValue(key, out var cur))
                    {
                        cur = new ZoneResult { Zone = key };
                        agg[key] = cur;
                    }
                    cur.ElementCount += z.ElementCount;
                    cur.FormworkArea += z.FormworkArea;
                }
            }

            int r = 2;
            foreach (var z in agg.Values.OrderBy(x => x.Zone))
            {
                ws.Cell(r, 1).Value = z.Zone;
                ws.Cell(r, 2).Value = z.ElementCount;
                ws.Cell(r, 3).Value = Round(z.FormworkArea);
                r++;
            }
            ws.Cell(r, 1).Value = "合計";
            ws.Cell(r, 1).Style.Font.Bold = true;
            ws.Cell(r, 2).Value = agg.Values.Sum(x => x.ElementCount);
            ws.Cell(r, 2).Style.Font.Bold = true;
            ws.Cell(r, 3).Value = Round(agg.Values.Sum(x => x.FormworkArea));
            ws.Cell(r, 3).Style.Font.Bold = true;
            AutoFitColumns(ws);
        }

        private static void WriteByTypeAggregated(XLWorkbook wb, IList<FormworkResult> results)
        {
            var ws = wb.Worksheets.Add("型枠種別");
            ws.Cell(1, 1).Value = "型枠種別";
            ws.Cell(1, 2).Value = "要素数";
            ws.Cell(1, 3).Value = "型枠面積 (㎡)";
            FormatHeader(ws.Row(1));

            var agg = new Dictionary<string, FormworkTypeResult>();
            foreach (var res in results)
            {
                if (res.TypeResults == null) continue;
                foreach (var t in res.TypeResults)
                {
                    string key = t.FormworkType ?? string.Empty;
                    if (!agg.TryGetValue(key, out var cur))
                    {
                        cur = new FormworkTypeResult { FormworkType = key };
                        agg[key] = cur;
                    }
                    cur.ElementCount += t.ElementCount;
                    cur.FormworkArea += t.FormworkArea;
                }
            }

            int r = 2;
            foreach (var t in agg.Values.OrderBy(x => x.FormworkType))
            {
                ws.Cell(r, 1).Value = t.FormworkType;
                ws.Cell(r, 2).Value = t.ElementCount;
                ws.Cell(r, 3).Value = Round(t.FormworkArea);
                r++;
            }
            ws.Cell(r, 1).Value = "合計";
            ws.Cell(r, 1).Style.Font.Bold = true;
            ws.Cell(r, 2).Value = agg.Values.Sum(x => x.ElementCount);
            ws.Cell(r, 2).Style.Font.Bold = true;
            ws.Cell(r, 3).Value = Round(agg.Values.Sum(x => x.FormworkArea));
            ws.Cell(r, 3).Style.Font.Bold = true;
            AutoFitColumns(ws);
        }

        private static void WriteElementDetailMulti(XLWorkbook wb,
            IList<FormworkResult> results, IList<string> sourceViewNames)
        {
            var ws = wb.Worksheets.Add("要素明細");
            bool multiView = results.Count > 1;
            int col = 1;
            ws.Cell(1, col++).Value = "要素ID";
            ws.Cell(1, col++).Value = "部位";
            ws.Cell(1, col++).Value = "要素名";
            if (multiView) ws.Cell(1, col++).Value = "ソースビュー";
            ws.Cell(1, col++).Value = "工区";
            ws.Cell(1, col++).Value = "型枠種別";
            ws.Cell(1, col++).Value = "型枠面積 (㎡)";
            ws.Cell(1, col++).Value = "天端控除";
            ws.Cell(1, col++).Value = "底面控除";
            ws.Cell(1, col++).Value = "接触控除";
            ws.Cell(1, col++).Value = "傾斜面";
            ws.Cell(1, col++).Value = "開口控除";
            ws.Cell(1, col++).Value = "開口見込み加算";
            FormatHeader(ws.Row(1));

            int r = 2;
            for (int i = 0; i < results.Count; i++)
            {
                string svName = sourceViewNames[i];
                foreach (var er in results[i].ElementResults)
                {
                    int c = 1;
                    ws.Cell(r, c++).Value = er.ElementId;
                    ws.Cell(r, c++).Value = CategoryLabel(er.Category);
                    ws.Cell(r, c++).Value = er.ElementName;
                    if (multiView) ws.Cell(r, c++).Value = svName;
                    ws.Cell(r, c++).Value = er.Zone;
                    ws.Cell(r, c++).Value = er.FormworkType;
                    ws.Cell(r, c++).Value = Round(er.FormworkArea);
                    ws.Cell(r, c++).Value = Round(er.DeductedTopArea);
                    ws.Cell(r, c++).Value = Round(er.DeductedBottomArea);
                    ws.Cell(r, c++).Value = Round(er.DeductedContactArea);
                    ws.Cell(r, c++).Value = Round(er.InclinedArea);
                    ws.Cell(r, c++).Value = Round(er.OpeningAreaDeducted);
                    ws.Cell(r, c++).Value = Round(er.OpeningEdgeAreaAdded);
                    r++;
                }
            }
            ws.SheetView.FreezeRows(1);
            ws.RangeUsed()?.SetAutoFilter();
            AutoFitColumns(ws, padding: 8.0);
        }

        private static void WriteErrorsMulti(XLWorkbook wb,
            IList<FormworkResult> results, IList<string> sourceViewNames)
        {
            var ws = wb.Worksheets.Add("エラー・注記");
            bool multiView = results.Count > 1;
            int col = 1;
            ws.Cell(1, col++).Value = "要素ID";
            ws.Cell(1, col++).Value = "カテゴリ";
            ws.Cell(1, col++).Value = "要素名";
            if (multiView) ws.Cell(1, col++).Value = "ソースビュー";
            ws.Cell(1, col++).Value = "種別";
            ws.Cell(1, col++).Value = "メッセージ";
            FormatHeader(ws.Row(1));

            int r = 2;
            for (int i = 0; i < results.Count; i++)
            {
                string svName = sourceViewNames[i];
                if (results[i].Errors == null) continue;
                foreach (var e in results[i].Errors)
                {
                    int c = 1;
                    ws.Cell(r, c++).Value = e.ElementId;
                    ws.Cell(r, c++).Value = e.CategoryName;
                    ws.Cell(r, c++).Value = e.ElementName;
                    if (multiView) ws.Cell(r, c++).Value = svName;
                    ws.Cell(r, c++).Value = e.ErrorKind;
                    ws.Cell(r, c++).Value = e.Message;
                    r++;
                }
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
                case CategoryGroup.Roof: return "屋根";
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
                    catch { s = cell.Value.ToString() ?? string.Empty; }
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

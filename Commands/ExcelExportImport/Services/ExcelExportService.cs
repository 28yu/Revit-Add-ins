using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using ClosedXML.Excel;
using Tools28.Commands.ExcelExportImport.Models;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// Excel書き出し処理サービス
    /// </summary>
    public static class ExcelExportService
    {
        /// <summary>
        /// パラメータデータをExcelファイルにエクスポート
        /// </summary>
        /// <param name="doc">Revitドキュメント</param>
        /// <param name="filePath">出力ファイルパス</param>
        /// <param name="selectedCategories">選択されたカテゴリ一覧</param>
        /// <param name="outputParameters">出力パラメータ一覧（順序付き）</param>
        /// <param name="splitByCategory">trueならカテゴリ毎にシート分割、falseなら1シート</param>
        /// <returns>エクスポート結果（カテゴリ名, 要素数）</returns>
        public static Dictionary<string, int> Export(
            Document doc,
            string filePath,
            List<CategoryInfo> selectedCategories,
            List<ParameterInfo> outputParameters,
            bool splitByCategory = true)
        {
            var results = new Dictionary<string, int>();

            using (var workbook = new XLWorkbook())
            {
                if (splitByCategory)
                {
                    ExportSplitByCategory(doc, workbook, selectedCategories, outputParameters, results);
                }
                else
                {
                    ExportSingleSheet(doc, workbook, selectedCategories, outputParameters, results);
                }

                workbook.SaveAs(filePath);
            }

            return results;
        }

        /// <summary>
        /// カテゴリ毎にシートを分割してエクスポート
        /// </summary>
        private static void ExportSplitByCategory(
            Document doc,
            XLWorkbook workbook,
            List<CategoryInfo> selectedCategories,
            List<ParameterInfo> outputParameters,
            Dictionary<string, int> results)
        {
            foreach (var category in selectedCategories)
            {
                // このカテゴリに属するパラメータのみ抽出
                var categoryParams = outputParameters
                    .Where(p => p.CategoryName == category.Name)
                    .ToList();

                if (categoryParams.Count == 0)
                    continue;

                // シート名を安全な文字列に変換（Excelの制限: 31文字以内）
                string sheetName = SanitizeSheetName(category.Name);

                var worksheet = workbook.Worksheets.Add(sheetName);

                // シート全体のフォントをＭＳ 明朝に設定
                worksheet.Style.Font.FontName = "ＭＳ 明朝";

                // ヘッダー行を作成
                worksheet.Cell(1, 1).Value = "要素ID";
                worksheet.Cell(1, 2).Value = "カテゴリ";
                for (int i = 0; i < categoryParams.Count; i++)
                {
                    var p = categoryParams[i];
                    string headerName = p.DisplayName;
                    if (p.IsReadOnly && p.RawName != "タイプ")
                    {
                        headerName += "(*変更不可)";
                    }
                    worksheet.Cell(1, i + 3).Value = headerName;
                }

                // ヘッダー行のスタイル設定
                var headerRange = worksheet.Range(1, 1, 1, categoryParams.Count + 2);
                headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(155, 187, 89);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                // 1行目の高さを25に設定
                worksheet.Row(1).Height = 25;

                // データ行を作成
                var elements = RevitCategoryHelper.GetElementsByCategory(doc, category.BuiltInCategory);
                int row = 2;

                foreach (var elem in elements)
                {
#if REVIT2026
                    worksheet.Cell(row, 1).Value = elem.Id.Value;
#else
                    worksheet.Cell(row, 1).Value = elem.Id.IntegerValue;
#endif
                    worksheet.Cell(row, 2).Value = category.Name;

                    for (int i = 0; i < categoryParams.Count; i++)
                    {
                        var paramInfo = categoryParams[i];
                        var param = ParameterService.FindParameter(
                            elem, paramInfo.RawName, paramInfo.IsTypeParameter, doc);
                        string value = ParameterService.GetParameterValueAsString(param);

                        if (double.TryParse(value, out double numValue))
                        {
                            worksheet.Cell(row, i + 3).Value = numValue;
                        }
                        else
                        {
                            worksheet.Cell(row, i + 3).Value = value;
                        }
                    }

                    row++;
                }

                AdjustColumnWidths(worksheet, categoryParams.Count + 2, row - 1);

                // オートフィルタを設定
                if (row > 2)
                {
                    worksheet.RangeUsed().SetAutoFilter();
                }

                results[category.Name] = elements.Count;
            }
        }

        /// <summary>
        /// 全カテゴリを1シートにまとめてエクスポート
        /// </summary>
        private static void ExportSingleSheet(
            Document doc,
            XLWorkbook workbook,
            List<CategoryInfo> selectedCategories,
            List<ParameterInfo> outputParameters,
            Dictionary<string, int> results)
        {
            // 全カテゴリ共通のユニークなパラメータリスト（DisplayName順序を維持）
            var allParams = outputParameters
                .GroupBy(p => p.DisplayName)
                .Select(g => g.First())
                .ToList();

            if (allParams.Count == 0)
                return;

            var worksheet = workbook.Worksheets.Add("データ");
            worksheet.Style.Font.FontName = "ＭＳ 明朝";

            // ヘッダー行を作成
            worksheet.Cell(1, 1).Value = "要素ID";
            worksheet.Cell(1, 2).Value = "カテゴリ";
            for (int i = 0; i < allParams.Count; i++)
            {
                var p = allParams[i];
                string headerName = p.DisplayName;
                if (p.IsReadOnly && p.RawName != "タイプ")
                {
                    headerName += "(*変更不可)";
                }
                worksheet.Cell(1, i + 3).Value = headerName;
            }

            // ヘッダー行のスタイル設定
            var headerRange = worksheet.Range(1, 1, 1, allParams.Count + 2);
            headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(155, 187, 89);
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            worksheet.Row(1).Height = 25;

            int row = 2;

            foreach (var category in selectedCategories)
            {
                var categoryParams = outputParameters
                    .Where(p => p.CategoryName == category.Name)
                    .ToList();

                if (categoryParams.Count == 0)
                    continue;

                var elements = RevitCategoryHelper.GetElementsByCategory(doc, category.BuiltInCategory);

                foreach (var elem in elements)
                {
#if REVIT2026
                    worksheet.Cell(row, 1).Value = elem.Id.Value;
#else
                    worksheet.Cell(row, 1).Value = elem.Id.IntegerValue;
#endif
                    worksheet.Cell(row, 2).Value = category.Name;

                    // 各パラメータの値を対応する列に書き込む
                    for (int i = 0; i < allParams.Count; i++)
                    {
                        var headerParam = allParams[i];
                        // このカテゴリに該当パラメータがあるか確認
                        var matchParam = categoryParams
                            .FirstOrDefault(p => p.DisplayName == headerParam.DisplayName);

                        if (matchParam == null)
                            continue;

                        var param = ParameterService.FindParameter(
                            elem, matchParam.RawName, matchParam.IsTypeParameter, doc);
                        string value = ParameterService.GetParameterValueAsString(param);

                        if (double.TryParse(value, out double numValue))
                        {
                            worksheet.Cell(row, i + 3).Value = numValue;
                        }
                        else
                        {
                            worksheet.Cell(row, i + 3).Value = value;
                        }
                    }

                    row++;
                }

                results[category.Name] = elements.Count;
            }

            AdjustColumnWidths(worksheet, allParams.Count + 2, row - 1);

            // オートフィルタを設定
            if (row > 2)
            {
                worksheet.RangeUsed().SetAutoFilter();
            }
        }

        /// <summary>
        /// 列幅を自動調整
        /// </summary>
        private static void AdjustColumnWidths(IXLWorksheet worksheet, int lastCol, int lastRow)
        {
            for (int col = 1; col <= lastCol; col++)
            {
                double maxDataWidth = 0;
                for (int r = 2; r <= lastRow; r++)
                {
                    string text = worksheet.Cell(r, col).GetString();
                    if (string.IsNullOrEmpty(text)) continue;
                    double w = CalculateTextWidth(text);
                    if (w > maxDataWidth) maxDataWidth = w;
                }
                string headerText = worksheet.Cell(1, col).GetString();
                double headerWidth = CalculateTextWidth(headerText ?? "") + 8;
                double dataWidth = maxDataWidth + 2;
                worksheet.Column(col).Width = Math.Max(headerWidth, dataWidth);
            }

            // 要素ID列の最小幅を確保
            if (worksheet.Column(1).Width < 12)
                worksheet.Column(1).Width = 12;
        }

        /// <summary>
        /// 文字列の表示幅をExcel列幅単位で計算（全角=2, 半角=1）
        /// </summary>
        private static double CalculateTextWidth(string text)
        {
            double width = 0;
            foreach (char c in text)
            {
                // 全角文字（日本語、全角英数記号など）は2、半角は1
                if (c > 0x7F)
                    width += 2.0;
                else
                    width += 1.0;
            }
            return width;
        }

        /// <summary>
        /// Excelシート名として安全な文字列に変換
        /// </summary>
        private static string SanitizeSheetName(string name)
        {
            // Excelのシート名禁止文字を置換
            char[] invalidChars = { '\\', '/', '*', '[', ']', ':', '?' };
            string result = name;
            foreach (char c in invalidChars)
            {
                result = result.Replace(c, '_');
            }

            // 31文字制限
            if (result.Length > 31)
                result = result.Substring(0, 31);

            return result;
        }
    }
}

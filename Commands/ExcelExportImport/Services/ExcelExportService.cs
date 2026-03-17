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
        /// <returns>エクスポート結果（カテゴリ名, 要素数）</returns>
        public static Dictionary<string, int> Export(
            Document doc,
            string filePath,
            List<CategoryInfo> selectedCategories,
            List<ParameterInfo> outputParameters)
        {
            var results = new Dictionary<string, int>();

            using (var workbook = new XLWorkbook())
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

                    // ヘッダー行を作成
                    worksheet.Cell(1, 1).Value = "要素ID";
                    worksheet.Cell(1, 2).Value = "カテゴリ";
                    for (int i = 0; i < categoryParams.Count; i++)
                    {
                        worksheet.Cell(1, i + 3).Value = categoryParams[i].DisplayName;
                    }

                    // ヘッダー行のスタイル設定
                    var headerRange = worksheet.Range(1, 1, 1, categoryParams.Count + 2);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(198, 239, 206);
                    headerRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                    // データ行を作成
                    var elements = RevitCategoryHelper.GetElementsByCategory(doc, category.BuiltInCategory);
                    int row = 2;

                    foreach (var elem in elements)
                    {
                        worksheet.Cell(row, 1).Value = elem.Id.IntegerValue;
                        worksheet.Cell(row, 2).Value = category.Name;

                        for (int i = 0; i < categoryParams.Count; i++)
                        {
                            var paramInfo = categoryParams[i];
                            var param = ParameterService.FindParameter(
                                elem, paramInfo.RawName, paramInfo.IsTypeParameter, doc);
                            string value = ParameterService.GetParameterValueAsString(param);
                            worksheet.Cell(row, i + 3).Value = value;
                        }

                        row++;
                    }

                    // 列幅の自動調整（余白を加えて文字が見切れないように）
                    worksheet.Columns().AdjustToContents(1, row, 2.0);

                    // 要素ID列の最小幅を確保
                    if (worksheet.Column(1).Width < 12)
                        worksheet.Column(1).Width = 12;

                    // オートフィルタを設定
                    if (row > 2)
                    {
                        worksheet.RangeUsed().SetAutoFilter();
                    }

                    results[category.Name] = elements.Count;
                }

                workbook.SaveAs(filePath);
            }

            return results;
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

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using ClosedXML.Excel;
using Tools28.Commands.ExcelExportImport.Models;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// インポート結果を保持するクラス
    /// </summary>
    public class ImportResult
    {
        public int SuccessCount { get; set; }
        public int FailCount { get; set; }
        public int SkipCount { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
    }

    /// <summary>
    /// インポートプレビュー行
    /// </summary>
    public class ImportPreviewRow
    {
        public int ElementId { get; set; }
        public string CategoryName { get; set; }
        public string ParameterName { get; set; }
        public string CurrentValue { get; set; }
        public string NewValue { get; set; }
        public bool HasChange { get; set; }
        public bool IsReadOnly { get; set; }
    }

    /// <summary>
    /// Excel読み込み処理サービス
    /// </summary>
    public static class ExcelImportService
    {
        /// <summary>
        /// Excelファイルのシート一覧を取得
        /// </summary>
        public static List<string> GetSheetNames(string filePath)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                return workbook.Worksheets.Select(ws => ws.Name).ToList();
            }
        }

        /// <summary>
        /// Excelファイルからインポートプレビューを生成
        /// </summary>
        public static List<ImportPreviewRow> GeneratePreview(Document doc, string filePath)
        {
            var preview = new List<ImportPreviewRow>();

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    var lastRow = worksheet.LastRowUsed();
                    var lastCol = worksheet.LastColumnUsed();
                    if (lastRow == null || lastCol == null)
                        continue;

                    int rowCount = lastRow.RowNumber();
                    int colCount = lastCol.ColumnNumber();

                    if (rowCount < 2 || colCount < 3)
                        continue;

                    // ヘッダーからパラメータ名を取得
                    var paramHeaders = new List<string>();
                    for (int col = 3; col <= colCount; col++)
                    {
                        paramHeaders.Add(worksheet.Cell(1, col).GetString());
                    }

                    // データ行を処理
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string elementIdStr = worksheet.Cell(row, 1).GetString();
                        string categoryName = worksheet.Cell(row, 2).GetString();

                        if (!int.TryParse(elementIdStr, out int elementIdInt))
                            continue;

                        var elementId = new ElementId(elementIdInt);
                        var elem = doc.GetElement(elementId);
                        if (elem == null)
                        {
                            // 要素が見つからない場合もプレビューに追加
                            for (int i = 0; i < paramHeaders.Count; i++)
                            {
                                preview.Add(new ImportPreviewRow
                                {
                                    ElementId = elementIdInt,
                                    CategoryName = categoryName,
                                    ParameterName = paramHeaders[i],
                                    CurrentValue = "（要素が見つかりません）",
                                    NewValue = worksheet.Cell(row, i + 3).GetString(),
                                    HasChange = false,
                                    IsReadOnly = true
                                });
                            }
                            continue;
                        }

                        for (int i = 0; i < paramHeaders.Count; i++)
                        {
                            string headerName = paramHeaders[i];
                            string newValue = worksheet.Cell(row, i + 3).GetString();

                            bool isTypeParam = headerName.StartsWith("T-");
                            string rawName = headerName.StartsWith("T-") || headerName.StartsWith("I-")
                                ? headerName.Substring(2)
                                : headerName;

                            var param = ParameterService.FindParameter(elem, rawName, isTypeParam, doc);
                            string currentValue = ParameterService.GetParameterValueAsString(param);
                            bool isReadOnly = param == null || param.IsReadOnly;
                            bool hasChange = currentValue != newValue && !isReadOnly;

                            preview.Add(new ImportPreviewRow
                            {
                                ElementId = elementIdInt,
                                CategoryName = categoryName,
                                ParameterName = headerName,
                                CurrentValue = currentValue,
                                NewValue = newValue,
                                HasChange = hasChange,
                                IsReadOnly = isReadOnly
                            });
                        }
                    }
                }
            }

            return preview;
        }

        /// <summary>
        /// Excelファイルからパラメータ値をインポート
        /// </summary>
        public static ImportResult Import(Document doc, string filePath)
        {
            var result = new ImportResult();

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    var lastRow = worksheet.LastRowUsed();
                    var lastCol = worksheet.LastColumnUsed();
                    if (lastRow == null || lastCol == null)
                        continue;

                    int rowCount = lastRow.RowNumber();
                    int colCount = lastCol.ColumnNumber();

                    if (rowCount < 2 || colCount < 3)
                        continue;

                    // ヘッダーからパラメータ名を取得
                    var paramHeaders = new List<string>();
                    for (int col = 3; col <= colCount; col++)
                    {
                        paramHeaders.Add(worksheet.Cell(1, col).GetString());
                    }

                    // データ行を処理
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string elementIdStr = worksheet.Cell(row, 1).GetString();

                        if (!int.TryParse(elementIdStr, out int elementIdInt))
                        {
                            result.FailCount++;
                            result.Errors.Add($"シート '{worksheet.Name}' 行{row}: ElementIdの解析に失敗");
                            continue;
                        }

                        var elementId = new ElementId(elementIdInt);
                        var elem = doc.GetElement(elementId);
                        if (elem == null)
                        {
                            result.FailCount++;
                            result.Errors.Add($"シート '{worksheet.Name}' 行{row}: ElementId {elementIdInt} が見つかりません");
                            continue;
                        }

                        for (int i = 0; i < paramHeaders.Count; i++)
                        {
                            string headerName = paramHeaders[i];
                            string newValue = worksheet.Cell(row, i + 3).GetString();

                            bool isTypeParam = headerName.StartsWith("T-");
                            string rawName = headerName.StartsWith("T-") || headerName.StartsWith("I-")
                                ? headerName.Substring(2)
                                : headerName;

                            var param = ParameterService.FindParameter(elem, rawName, isTypeParam, doc);

                            if (param == null)
                            {
                                result.Warnings.Add($"シート '{worksheet.Name}' 行{row}: パラメータ '{headerName}' が見つかりません");
                                result.SkipCount++;
                                continue;
                            }

                            if (param.IsReadOnly)
                            {
                                result.SkipCount++;
                                continue;
                            }

                            // 現在値と同じならスキップ
                            string currentValue = ParameterService.GetParameterValueAsString(param);
                            if (currentValue == newValue)
                            {
                                result.SkipCount++;
                                continue;
                            }

                            if (ParameterService.SetParameterValue(param, newValue))
                            {
                                result.SuccessCount++;
                            }
                            else
                            {
                                result.FailCount++;
                                result.Errors.Add($"シート '{worksheet.Name}' 行{row}: パラメータ '{headerName}' の値設定に失敗（値: '{newValue}'）");
                            }
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// インポートで変更されたセルにExcelファイル上で色を付ける
        /// </summary>
        public static void MarkImportedCells(string filePath, List<ImportPreviewRow> previewRows)
        {
            // 変更のあったセルを (ElementId, ParameterName) で検索用セットに
            var changedSet = new HashSet<string>(
                previewRows
                    .Where(r => r.HasChange)
                    .Select(r => r.ElementId + "|" + r.ParameterName));

            if (changedSet.Count == 0)
                return;

            using (var workbook = new XLWorkbook(filePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    var lastRow = worksheet.LastRowUsed();
                    var lastCol = worksheet.LastColumnUsed();
                    if (lastRow == null || lastCol == null)
                        continue;

                    int rowCount = lastRow.RowNumber();
                    int colCount = lastCol.ColumnNumber();

                    if (rowCount < 2 || colCount < 3)
                        continue;

                    // ヘッダーからパラメータ名を取得
                    var paramHeaders = new List<string>();
                    for (int col = 3; col <= colCount; col++)
                    {
                        paramHeaders.Add(worksheet.Cell(1, col).GetString());
                    }

                    // データ行を走査して該当セルに色を付ける
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string elementIdStr = worksheet.Cell(row, 1).GetString();

                        for (int i = 0; i < paramHeaders.Count; i++)
                        {
                            string key = elementIdStr + "|" + paramHeaders[i];
                            if (changedSet.Contains(key))
                            {
                                worksheet.Cell(row, i + 3).Style.Fill.BackgroundColor =
                                    XLColor.FromArgb(252, 213, 180);
                            }
                        }
                    }
                }

                workbook.Save();
            }
        }
    }
}

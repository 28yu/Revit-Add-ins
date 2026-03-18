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
                                    NewValue = GetCellValueAsString(worksheet.Cell(row, i + 3)),
                                    HasChange = false,
                                    IsReadOnly = true
                                });
                            }
                            continue;
                        }

                        for (int i = 0; i < paramHeaders.Count; i++)
                        {
                            string headerName = paramHeaders[i];
                            string newValue = GetCellValueAsString(worksheet.Cell(row, i + 3));

                            bool isTypeParam = headerName.StartsWith("T-");
                            string rawName = headerName.StartsWith("T-") || headerName.StartsWith("I-")
                                ? headerName.Substring(2)
                                : headerName;

                            var param = ParameterService.FindParameter(elem, rawName, isTypeParam, doc);
                            string currentValue = ParameterService.GetParameterValueAsString(param);
                            bool isReadOnly = param == null || param.IsReadOnly;
                            bool hasChange = !ValuesAreEqual(currentValue, newValue) && !isReadOnly;

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
                            string newValue = GetCellValueAsString(worksheet.Cell(row, i + 3));

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
                            if (ValuesAreEqual(currentValue, newValue))
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
        /// Excelが開いている場合はCOM経由で直接色付け、閉じている場合はClosedXMLで上書き
        /// </summary>
        /// <returns>色付けファイルの保存先パス。COM経由成功時は元ファイルパス。色付け不要/失敗の場合はnull</returns>
        public static string MarkImportedCells(string filePath, List<ImportPreviewRow> previewRows)
        {
            // 変更のあったセルを (ElementId, ParameterName) で検索用セットに
            var changedSet = new HashSet<string>(
                previewRows
                    .Where(r => r.HasChange)
                    .Select(r => r.ElementId.ToString() + "|" + r.ParameterName));

            if (changedSet.Count == 0)
                return null;

            // まずCOM経由（開いているExcelに直接色付け）を試行
            if (ExcelProcessHelper.MarkCellsViaCom(filePath, changedSet))
                return filePath;

            // Excelが開いていない場合、ClosedXMLでファイルを直接編集
            return MarkCellsViaClosedXml(filePath, changedSet);
        }

        /// <summary>
        /// Excelセルの値を文字列として取得（数値セルは整数なら小数点なしで返す）
        /// </summary>
        private static string GetCellValueAsString(IXLCell cell)
        {
            if (cell.DataType == XLDataType.Number)
            {
                double numVal = cell.GetDouble();
                // 整数なら小数点なしの文字列にする（Revit の AsValueString() と一致させる）
                if (numVal == Math.Floor(numVal))
                    return ((long)numVal).ToString();
                return numVal.ToString();
            }
            return cell.GetString();
        }

        /// <summary>
        /// 2つの値が等しいか比較（数値の場合は数値比較、テキストは文字列比較）
        /// </summary>
        private static bool ValuesAreEqual(string val1, string val2)
        {
            if (val1 == val2)
                return true;
            if (string.IsNullOrEmpty(val1) && string.IsNullOrEmpty(val2))
                return true;

            // 両方が数値の場合は数値として比較（"4700" vs "4700.0" 等の差異を吸収）
            if (double.TryParse(val1, out double d1) && double.TryParse(val2, out double d2))
                return Math.Abs(d1 - d2) < 0.0001;

            return false;
        }

        /// <summary>
        /// ClosedXMLを使用してExcelファイルのセルに色を付ける（Excelが閉じている場合のフォールバック）
        /// </summary>
        private static string MarkCellsViaClosedXml(string filePath, HashSet<string> changedSet)
        {
            byte[] fileBytes;
            try
            {
                using (var readStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var memRead = new MemoryStream();
                    readStream.CopyTo(memRead);
                    fileBytes = memRead.ToArray();
                }
            }
            catch (IOException)
            {
                return null;
            }

            using (var memStream = new MemoryStream(fileBytes))
            using (var workbook = new XLWorkbook(memStream))
            {
                bool anyMarked = false;

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

                    var paramHeaders = new List<string>();
                    for (int col = 3; col <= colCount; col++)
                    {
                        paramHeaders.Add(worksheet.Cell(1, col).GetString());
                    }

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string elementIdStr = worksheet.Cell(row, 1).GetString().Trim();
                        if (double.TryParse(elementIdStr, out double idDouble))
                        {
                            elementIdStr = ((int)idDouble).ToString();
                        }

                        // この行に変更があるかチェック
                        bool rowHasChange = false;
                        for (int i = 0; i < paramHeaders.Count; i++)
                        {
                            string key = elementIdStr + "|" + paramHeaders[i];
                            if (changedSet.Contains(key))
                            {
                                rowHasChange = true;
                                break;
                            }
                        }

                        // 変更がある行は全列に色を付ける
                        if (rowHasChange)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                worksheet.Cell(row, col).Style.Fill.BackgroundColor =
                                    XLColor.FromArgb(255, 255, 153);
                            }
                            anyMarked = true;
                        }
                    }
                }

                if (!anyMarked)
                    return null;

                using (var saveStream = new MemoryStream())
                {
                    workbook.SaveAs(saveStream);
                    byte[] savedBytes = saveStream.ToArray();

                    try
                    {
                        File.WriteAllBytes(filePath, savedBytes);
                        return filePath;
                    }
                    catch (IOException)
                    {
                        string dir = Path.GetDirectoryName(filePath);
                        string name = Path.GetFileNameWithoutExtension(filePath);
                        string ext = Path.GetExtension(filePath);
                        string altPath = Path.Combine(dir, name + "_imported" + ext);
                        File.WriteAllBytes(altPath, savedBytes);
                        return altPath;
                    }
                }
            }
        }
    }
}

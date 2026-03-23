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
        /// <summary>インポートに失敗したセルのキー（"ElementId|ParameterName"）</summary>
        public HashSet<string> FailedCells { get; set; } = new HashSet<string>();
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

                    // ヘッダーからパラメータ名を取得（(*変更不可)サフィックスは除去）
                    var paramHeaders = new List<string>();
                    for (int col = 3; col <= colCount; col++)
                    {
                        string header = worksheet.Cell(1, col).GetString();
                        paramHeaders.Add(StripReadOnlySuffix(header));
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
                            // タイプ変更パラメータはIsReadOnlyでもChangeTypeIdで変更可能
                            bool isReadOnly = param == null
                                || (param.IsReadOnly && !ParameterService.IsTypeChangeParameter(param));
                            bool hasChange = !ValuesAreEqual(currentValue, newValue);

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

                    // ヘッダーからパラメータ名を取得（(*変更不可)サフィックスは除去）
                    var paramHeaders = new List<string>();
                    for (int col = 3; col <= colCount; col++)
                    {
                        string header = worksheet.Cell(1, col).GetString();
                        paramHeaders.Add(StripReadOnlySuffix(header));
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

                            if (param.IsReadOnly && !ParameterService.IsTypeChangeParameter(param))
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

                            // タイプ変更パラメータ（ELEM_TYPE_PARAM）はChangeTypeIdで処理
                            bool success;
                            if (ParameterService.IsTypeChangeParameter(param))
                            {
                                success = ParameterService.ChangeElementType(elem, newValue, doc);
                            }
                            else
                            {
                                success = ParameterService.SetParameterValue(param, newValue);
                            }

                            if (success)
                            {
                                result.SuccessCount++;
                            }
                            else
                            {
                                result.FailCount++;
                                result.FailedCells.Add(elementIdInt.ToString() + "|" + headerName);
                                string errorDetail = ParameterService.IsTypeChangeParameter(param)
                                    ? $"シート '{worksheet.Name}' 行{row}: タイプ変更に失敗（値: '{newValue}'）— 一致するタイプが見つかりません"
                                    : $"シート '{worksheet.Name}' 行{row}: パラメータ '{headerName}' の値設定に失敗（値: '{newValue}'）";
                                result.Errors.Add(errorDetail);
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
        public static string MarkImportedCells(string filePath, List<ImportPreviewRow> previewRows, out string colorMethod, HashSet<string> failedSet = null)
        {
            colorMethod = null;

            // 実際にインポートされたセル（変更あり＋書き込み可能）を検索用セットに
            var changedSet = new HashSet<string>(
                previewRows
                    .Where(r => r.HasChange && !r.IsReadOnly)
                    .Select(r => r.ElementId.ToString() + "|" + r.ParameterName));

            // 成功セットから失敗セルを除外
            if (failedSet != null && failedSet.Count > 0)
            {
                changedSet.ExceptWith(failedSet);
            }

            if (changedSet.Count == 0 && (failedSet == null || failedSet.Count == 0))
                return null;

            // まずCOM経由（開いているExcelに直接色付け）を試行
            if (ExcelProcessHelper.MarkCellsViaCom(filePath, changedSet, failedSet))
            {
                colorMethod = "COM";
                return filePath;
            }

            // Excelが開いていない or COM失敗の場合、ClosedXMLでファイルを直接編集
            colorMethod = "ClosedXML";
            return MarkCellsViaClosedXml(filePath, changedSet, failedSet);
        }

        /// <summary>
        /// インポートで変更されたセルにExcelファイル上で色を付ける（互換用オーバーロード）
        /// </summary>
        public static string MarkImportedCells(string filePath, List<ImportPreviewRow> previewRows)
        {
            return MarkImportedCells(filePath, previewRows, out _);
        }

        /// <summary>
        /// ヘッダー名から(*変更不可)サフィックスを除去
        /// </summary>
        private static string StripReadOnlySuffix(string headerName)
        {
            if (headerName != null && headerName.EndsWith("(*変更不可)"))
                return headerName.Substring(0, headerName.Length - "(*変更不可)".Length);
            return headerName;
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
        private static string MarkCellsViaClosedXml(string filePath, HashSet<string> changedSet, HashSet<string> failedSet = null)
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

            // 成功セル: 青(R79,G129,B189)/太字、失敗セル: 赤/太字
            var blueColor = XLColor.FromArgb(79, 129, 189);

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
                        paramHeaders.Add(StripReadOnlySuffix(worksheet.Cell(1, col).GetString()));
                    }

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string elementIdStr = worksheet.Cell(row, 1).GetString().Trim();
                        if (double.TryParse(elementIdStr, out double idDouble))
                        {
                            elementIdStr = ((int)idDouble).ToString();
                        }

                        // セル単位で成功/失敗を判定
                        var successCols = new HashSet<int>();
                        var failedCols = new HashSet<int>();
                        for (int i = 0; i < paramHeaders.Count; i++)
                        {
                            string key = elementIdStr + "|" + paramHeaders[i];
                            if (changedSet.Contains(key))
                                successCols.Add(i + 3);
                            else if (failedSet != null && failedSet.Contains(key))
                                failedCols.Add(i + 3);
                        }

                        // 変更・失敗がある行は全列に背景色、セル単位で色分け
                        if (successCols.Count > 0 || failedCols.Count > 0)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                worksheet.Cell(row, col).Style.Fill.BackgroundColor =
                                    XLColor.FromArgb(255, 255, 153);
                            }
                            foreach (int col in successCols)
                            {
                                worksheet.Cell(row, col).Style.Font.FontColor = blueColor;
                                worksheet.Cell(row, col).Style.Font.Bold = true;
                            }
                            foreach (int col in failedCols)
                            {
                                worksheet.Cell(row, col).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(row, col).Style.Font.Bold = true;
                            }
                            anyMarked = true;
                        }
                    }
                }

                if (!anyMarked)
                    return null;

                // 各シートの1行目（最終列の次）に凡例を追加
                foreach (var worksheet in workbook.Worksheets)
                {
                    var lastCol = worksheet.LastColumnUsed();
                    if (lastCol == null) continue;
                    int legendCol = lastCol.ColumnNumber() + 1;

                    var legendCell = worksheet.Cell(1, legendCol);
                    var richText = legendCell.CreateRichText();
                    richText.AddText("(*");
                    var bluePart = richText.AddText("青字");
                    bluePart.SetFontColor(blueColor);
                    bluePart.SetBold(true);
                    richText.AddText("はインポート成功、");
                    var redPart = richText.AddText("赤字");
                    redPart.SetFontColor(XLColor.Red);
                    redPart.SetBold(true);
                    richText.AddText("はインポート失敗)");
                }

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

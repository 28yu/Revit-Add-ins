using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28;
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
            bool splitByCategory = true,
            ExportScope scope = ExportScope.EntireProject,
            View activeView = null,
            ICollection<ElementId> selectionIds = null)
        {
            var results = new Dictionary<string, int>();

            using (var workbook = new XLWorkbook())
            {
                if (splitByCategory)
                {
                    ExportSplitByCategory(doc, workbook, selectedCategories, outputParameters, results, scope, activeView, selectionIds);
                }
                else
                {
                    ExportSingleSheet(doc, workbook, selectedCategories, outputParameters, results, scope, activeView, selectionIds);
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
            Dictionary<string, int> results,
            ExportScope scope,
            View activeView,
            ICollection<ElementId> selectionIds)
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

                // ヘッダー行（1行目）を固定してスクロール時も常に表示
                worksheet.SheetView.FreezeRows(1);

                // データ行を作成
                var elements = RevitCategoryHelper.GetElementsByCategory(
                    doc, category.BuiltInCategory, scope, activeView, selectionIds);

                // 列幅計算用（ヘッダー幅で初期化。2回目の全走査を避けるため書き込みと同時に集計）
                int totalCols = categoryParams.Count + 2;
                double[] colWidths = new double[totalCols];
                for (int c = 0; c < totalCols; c++)
                    colWidths[c] = CalculateTextWidth(worksheet.Cell(1, c + 1).GetString() ?? "") + 8;

                // タイプパラメータ値のキャッシュ（同一タイプのインスタンスは値が同じ）
                var typeElemCache = new Dictionary<long, Element>();
                var typeValueCache = new Dictionary<string, string>();

                int row = 2;

                foreach (var elem in elements)
                {
                    long elemId = ElementIdToLong(elem.Id);
                    worksheet.Cell(row, 1).Value = elemId;
                    UpdateColWidth(colWidths, 0, elemId.ToString());

                    worksheet.Cell(row, 2).Value = category.Name;
                    UpdateColWidth(colWidths, 1, category.Name);

                    for (int i = 0; i < categoryParams.Count; i++)
                    {
                        var paramInfo = categoryParams[i];
                        string value = ResolveParameterValue(
                            elem, paramInfo, doc, typeElemCache, typeValueCache);

                        if (double.TryParse(value, out double numValue))
                        {
                            worksheet.Cell(row, i + 3).Value = numValue;
                        }
                        else
                        {
                            worksheet.Cell(row, i + 3).Value = value;
                        }
                        UpdateColWidth(colWidths, i + 2, value);
                    }

                    row++;
                }

                ApplyColumnWidths(worksheet, colWidths);

                // オートフィルタを設定（全走査を避けるため既知の範囲を直接指定）
                if (row > 2)
                {
                    worksheet.Range(1, 1, row - 1, totalCols).SetAutoFilter();
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
            Dictionary<string, int> results,
            ExportScope scope,
            View activeView,
            ICollection<ElementId> selectionIds)
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

            // ヘッダー行（1行目）を固定してスクロール時も常に表示
            worksheet.SheetView.FreezeRows(1);

            // 列幅計算用（ヘッダー幅で初期化）
            int totalCols = allParams.Count + 2;
            double[] colWidths = new double[totalCols];
            for (int c = 0; c < totalCols; c++)
                colWidths[c] = CalculateTextWidth(worksheet.Cell(1, c + 1).GetString() ?? "") + 8;

            // タイプパラメータ値のキャッシュ（同一タイプのインスタンスは値が同じ）
            var typeElemCache = new Dictionary<long, Element>();
            var typeValueCache = new Dictionary<string, string>();

            int row = 2;

            foreach (var category in selectedCategories)
            {
                var categoryParams = outputParameters
                    .Where(p => p.CategoryName == category.Name)
                    .ToList();

                if (categoryParams.Count == 0)
                    continue;

                // ヘッダー列（DisplayName）→ このカテゴリのパラメータ の対応を事前構築
                var paramByDisplay = new Dictionary<string, ParameterInfo>();
                foreach (var p in categoryParams)
                    if (!paramByDisplay.ContainsKey(p.DisplayName))
                        paramByDisplay[p.DisplayName] = p;

                var elements = RevitCategoryHelper.GetElementsByCategory(
                    doc, category.BuiltInCategory, scope, activeView, selectionIds);

                foreach (var elem in elements)
                {
                    long elemId = ElementIdToLong(elem.Id);
                    worksheet.Cell(row, 1).Value = elemId;
                    UpdateColWidth(colWidths, 0, elemId.ToString());

                    worksheet.Cell(row, 2).Value = category.Name;
                    UpdateColWidth(colWidths, 1, category.Name);

                    // 各パラメータの値を対応する列に書き込む
                    for (int i = 0; i < allParams.Count; i++)
                    {
                        // このカテゴリに該当パラメータがあるか確認
                        if (!paramByDisplay.TryGetValue(allParams[i].DisplayName, out var matchParam))
                            continue;

                        string value = ResolveParameterValue(
                            elem, matchParam, doc, typeElemCache, typeValueCache);

                        if (double.TryParse(value, out double numValue))
                        {
                            worksheet.Cell(row, i + 3).Value = numValue;
                        }
                        else
                        {
                            worksheet.Cell(row, i + 3).Value = value;
                        }
                        UpdateColWidth(colWidths, i + 2, value);
                    }

                    row++;
                }

                results[category.Name] = elements.Count;
            }

            ApplyColumnWidths(worksheet, colWidths);

            // オートフィルタを設定（全走査を避けるため既知の範囲を直接指定）
            if (row > 2)
            {
                worksheet.Range(1, 1, row - 1, totalCols).SetAutoFilter();
            }
        }

        /// <summary>
        /// ElementId を long に変換（バージョン差異を吸収）
        /// </summary>
        private static long ElementIdToLong(ElementId id)
        {
#if REVIT2026
            return id.Value;
#else
            return id.IntValue();
#endif
        }

        /// <summary>
        /// パラメータ値を文字列で取得（タイプパラメータはキャッシュを利用）
        /// </summary>
        /// <remarks>
        /// タイプパラメータは同一タイプの全インスタンスで値が同じため、
        /// (タイプID, パラメータ名) 単位でキャッシュして再計算を避ける。
        /// インスタンスパラメータは <see cref="Element.LookupParameter"/> で直接引く
        /// （全パラメータの線形走査を回避）。
        /// </remarks>
        private static string ResolveParameterValue(
            Element elem,
            ParameterInfo paramInfo,
            Document doc,
            Dictionary<long, Element> typeElemCache,
            Dictionary<string, string> typeValueCache)
        {
            if (paramInfo.IsTypeParameter)
            {
                var typeId = elem.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId)
                    return "";

                long tid = ElementIdToLong(typeId);
                string cacheKey = tid + "|" + paramInfo.RawName;
                if (typeValueCache.TryGetValue(cacheKey, out string cached))
                    return cached;

                if (!typeElemCache.TryGetValue(tid, out Element typeElem))
                {
                    typeElem = doc.GetElement(typeId);
                    typeElemCache[tid] = typeElem;
                }

                string value = "";
                if (typeElem != null)
                {
                    var p = typeElem.LookupParameter(paramInfo.RawName);
                    value = ParameterService.GetParameterValueAsString(p);
                }

                typeValueCache[cacheKey] = value;
                return value;
            }
            else
            {
                var p = elem.LookupParameter(paramInfo.RawName);
                return ParameterService.GetParameterValueAsString(p);
            }
        }

        /// <summary>
        /// 指定列の最大表示幅を更新（書き込みと同時に集計するためのヘルパー）
        /// </summary>
        private static void UpdateColWidth(double[] colWidths, int colIndex, string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            double w = CalculateTextWidth(text) + 2;
            if (w > colWidths[colIndex]) colWidths[colIndex] = w;
        }

        /// <summary>
        /// 集計済みの列幅をワークシートに適用
        /// </summary>
        private static void ApplyColumnWidths(IXLWorksheet worksheet, double[] colWidths)
        {
            for (int c = 0; c < colWidths.Length; c++)
                worksheet.Column(c + 1).Width = colWidths[c];

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

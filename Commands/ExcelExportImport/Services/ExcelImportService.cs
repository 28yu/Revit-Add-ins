using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using ClosedXML.Excel;
using Tools28;
using Tools28.Commands.ExcelExportImport.Models;
using Tools28.Localization;

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

        // --- スキップの内訳（SkipCount = 3つの合計）---
        /// <summary>読み取り専用（長さ・面積などの計算値）で書き込めない</summary>
        public int SkipReadOnly { get; set; }
        /// <summary>その要素にパラメータが無い</summary>
        public int SkipNotFound { get; set; }
        /// <summary>書き込み時点で既に同じ値だった</summary>
        public int SkipUnchanged { get; set; }
        /// <summary>同名パラメータの複数の列に違う値が入っていて、どれを書くか決められない</summary>
        public int SkipConflict { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
        /// <summary>インポートに失敗したセルのキー（"ElementId|ParameterName"）</summary>
        public HashSet<string> FailedCells { get; set; } = new HashSet<string>();
    }

    /// <summary>
    /// インポートプレビュー行
    /// </summary>
    public class ImportPreviewRow : System.ComponentModel.INotifyPropertyChanged
    {
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        private bool _isSelected = true;

        /// <summary>
        /// 変更プレビューで「取り込む」にチェックされているか（既定はチェックあり）。
        /// インポートダイアログで外した行は書き込み・Excel の色付けの対象外になる。
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        /// <summary>
        /// 要素 Id。Revit 2026 で要素 Id が 64bit 化されたため long で保持する
        /// （書き出し側も long。int だと大きな Id の行が読み込めない）。
        /// </summary>
        public long ElementId { get; set; }
        public string CategoryName { get; set; }
        public string ParameterName { get; set; }
        public string CurrentValue { get; set; }
        public string NewValue { get; set; }
        public bool HasChange { get; set; }
        public bool IsReadOnly { get; set; }

        /// <summary>
        /// 書き出し時に記録したパラメータの識別番号（隠しシートから取得。0 = 不明）。
        /// 同名パラメータを取り違えないよう、書き込み時もこの番号でパラメータを特定する。
        /// </summary>
        public long ParamId { get; set; }

        /// <summary>この要素にパラメータが無い（IsReadOnly も true になる）。スキップ内訳の集計用。</summary>
        public bool ParamMissing { get; set; }

        /// <summary>
        /// 同名パラメータの複数の列（【共有#1】【共有#2】等）に違う値が入っていて書き込めない（IsReadOnly も true）。
        /// </summary>
        public bool Conflict { get; set; }

        /// <summary>
        /// Excel の要素IDの要素が開いているモデルに無い（または別カテゴリの要素だった）ことを示す目印の行。
        /// 要素1つにつき1行だけ作る（以前は列の数だけ作っていたため、件数が膨大になり内訳も分からなかった）。
        /// </summary>
        public bool ElementMissing { get; set; }

        /// <summary>要素IDは見つかったが、Excel のカテゴリと実際のカテゴリが違う（別モデルの可能性）</summary>
        public bool CategoryMismatch { get; set; }
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
            return GeneratePreview(doc, filePath, out _);
        }

        /// <summary>
        /// Excelファイルからインポートプレビューを生成し、シート名一覧も同時に取得する。
        /// （ファイルを1回開くだけで済み、GetSheetNames の別途オープン（二重読込）を避けられる）
        /// </summary>
        public static List<ImportPreviewRow> GeneratePreview(Document doc, string filePath, out List<string> sheetNames)
        {
            var preview = new List<ImportPreviewRow>();
            sheetNames = new List<string>();

            // タイプパラメータの現在値・読取専用フラグは同一タイプの全インスタンスで共通。
            // プレビューは読み取りのみなので (タイプID|パラメータ名) 単位でキャッシュして再計算を避ける。
            var typeCurrentCache = new Dictionary<string, string>();
            var typeReadOnlyCache = new Dictionary<string, bool>();
            var typeMissingCache = new Dictionary<string, bool>();

            // 変更判定の診断ログ（「変更していないのに出る／変更したのに出ない」の原因調査用）
            var diag = new PreviewDiagnostics();
            DiagLog.Write($"[ImportPreview] 開始: {filePath}");

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                // 列ごとのパラメータ識別番号（このアドインの新しい版で書き出した Excel のみ。無ければ空）
                var paramIdMap = ParameterIdSheet.Read(workbook);
                DiagLog.Write($"[ImportPreview] パラメータ識別番号の記録: {paramIdMap.Count} 件");

                foreach (var worksheet in workbook.Worksheets)
                {
                    if (ParameterIdSheet.IsMetaSheet(worksheet))
                        continue; // 識別番号の隠しシート（データではない）

                    sheetNames.Add(worksheet.Name);

                    // 1シート統合モードのシートでは、他カテゴリ用の列も同じ行に並び、
                    // その要素に関係ない列は空欄で書き出される。空欄を「値の削除」と解釈すると
                    // 触っていない値が消えてしまうため、統合シートでは空欄を常にスキップする。
                    bool isMergedSheet = ExcelHeaderNames.IsMergedSheetName(worksheet.Name);

                    var lastRow = worksheet.LastRowUsed();
                    var lastCol = worksheet.LastColumnUsed();
                    if (lastRow == null || lastCol == null)
                        continue;

                    int rowCount = lastRow.RowNumber();
                    int colCount = lastCol.ColumnNumber();

                    // 見出し行（グループ行付きで書き出した Excel は 2行目）
                    int headerRow = ExcelHeaderNames.FindHeaderRow(worksheet);

                    DiagLog.Write($"[ImportPreview] シート「{worksheet.Name}」 見出し行={headerRow} 最終行={rowCount} 最終列={colCount} 統合シート={isMergedSheet}");

                    if (rowCount <= headerRow || colCount < 3)
                        continue;

                    // ヘッダーからパラメータ名を取得（(*変更不可)サフィックスは除去）
                    var paramHeaders = new List<string>();
                    for (int col = 3; col <= colCount; col++)
                    {
                        string header = worksheet.Cell(headerRow, col).GetString();
                        paramHeaders.Add(StripReadOnlySuffix(header));
                    }

                    // 同名パラメータの列グループ（例: 「I-用途【共有#1】」「【共有#2】」「【共有#4】」）。
                    // 要素は普通そのうち1つしか持たないため、行ごとにグループ単位でまとめて判定する。
                    var sameNameGroups = BuildSameNameGroups(paramHeaders);

                    // データ行を処理
                    for (int row = headerRow + 1; row <= rowCount; row++)
                    {
                        string elementIdStr = worksheet.Cell(row, 1).GetString();
                        string categoryName = worksheet.Cell(row, 2).GetString();

                        if (!TryParseElementId(elementIdStr, out long elementIdInt))
                            continue;

                        var elementId = ToElementId(elementIdInt);
                        var elem = doc.GetElement(elementId);
                        if (elem == null)
                        {
                            // 要素が見つからない（書き出し元と別のモデル、または削除済み）→ 要素ごとに1行だけ目印を残す
                            preview.Add(new ImportPreviewRow
                            {
                                ElementId = elementIdInt,
                                CategoryName = categoryName,
                                ParameterName = "",
                                CurrentValue = "（要素が見つかりません）",
                                NewValue = "",
                                HasChange = false,
                                IsReadOnly = true,
                                ElementMissing = true
                            });
                            diag.ElementNotFound++;
                            continue;
                        }

                        // 同じ要素IDでもカテゴリが違えば別の要素（別モデルで IDがたまたま一致した等）。
                        // 誤って別の要素に書き込まないよう、この行は取り込まない。
                        string actualCategory = null;
                        try { actualCategory = elem.Category?.Name; } catch { }
                        if (!string.IsNullOrEmpty(categoryName) && !string.IsNullOrEmpty(actualCategory)
                            && actualCategory != categoryName)
                        {
                            preview.Add(new ImportPreviewRow
                            {
                                ElementId = elementIdInt,
                                CategoryName = categoryName,
                                ParameterName = "",
                                CurrentValue = actualCategory,
                                NewValue = "",
                                HasChange = false,
                                IsReadOnly = true,
                                ElementMissing = true,
                                CategoryMismatch = true
                            });
                            diag.LogCategoryMismatch(worksheet.Name, row, elementIdInt, categoryName, actualCategory);
                            continue;
                        }

                        // 同名パラメータの列グループを先に処理（要素が持つ同名パラメータが1つ以下の場合）
                        HashSet<int> handledCols = null;
                        foreach (var group in sameNameGroups)
                        {
                            if (TryProcessSameNameGroup(doc, elem, worksheet, row, elementIdInt, categoryName,
                                    paramHeaders, group, isMergedSheet, preview, diag))
                            {
                                if (handledCols == null) handledCols = new HashSet<int>();
                                foreach (int c in group) handledCols.Add(c);
                            }
                        }

                        for (int i = 0; i < paramHeaders.Count; i++)
                        {
                            if (handledCols != null && handledCols.Contains(i))
                                continue;   // 同名列グループとして処理済み

                            string headerName = paramHeaders[i];
                            string newValue = GetCellValueAsString(worksheet.Cell(row, i + 3));

                            // 同名パラメータ区別の接尾辞（【組み込み】等）も含めて解析する
                            var parsed = ParameterService.ParseDisplayName(headerName);
                            bool isTypeParam = parsed.IsTypeParameter;
                            string rawName = parsed.RawName;

                            // 同名パラメータ（【共有#2】等）は書き出し時の識別番号で正確に特定する
                            long paramId = 0;
                            if (parsed.Kind != null)
                                paramIdMap.TryGetValue(ParameterIdSheet.Key(categoryName, headerName), out paramId);

                            // 空セルの扱い:
                            //  - パラメータが要素に存在しない → N/A（シート統合モードの他カテゴリ列など）→ スキップ
                            //  - 文字列パラメータで現在値が空でない → 「値の削除（クリア）」として取り込む
                            //    （数値・ElementId 型は Revit 上で空にできないためスキップ）
                            if (string.IsNullOrEmpty(newValue))
                            {
                                if (isMergedSheet)
                                {
                                    diag.BlankSkipped++;
                                    diag.BlankMerged++;
                                    continue;
                                }

                                var clearParam = ParameterService.FindParameter(elem, rawName, isTypeParam, parsed.Kind, parsed.Index, paramId, doc);
                                if (clearParam == null || clearParam.StorageType != StorageType.String)
                                {
                                    diag.BlankSkipped++;
                                    if (clearParam == null) diag.BlankNoParam++; else diag.BlankNotText++;
                                    continue;
                                }

                                string clearCurrent = ParameterService.GetParameterValueAsString(clearParam);
                                if (string.IsNullOrEmpty(clearCurrent))
                                {
                                    diag.BlankSkipped++;
                                    diag.BlankAlreadyEmpty++;
                                    continue; // 既に空 → 変更なし
                                }

                                diag.LogChange(worksheet.Name, row, elementIdInt, headerName,
                                    worksheet.Cell(row, i + 3), clearCurrent, "", clearParam, "空欄=値の削除");

                                preview.Add(new ImportPreviewRow
                                {
                                    ElementId = elementIdInt,
                                    CategoryName = categoryName,
                                    ParameterName = headerName,
                                    CurrentValue = clearCurrent,
                                    NewValue = "",
                                    HasChange = true,
                                    IsReadOnly = clearParam.IsReadOnly && !ParameterService.IsTypeChangeParameter(clearParam),
                                    ParamId = paramId
                                });
                                continue;
                            }

                            string currentValue;
                            bool isReadOnly;
                            bool paramMissing;
                            Parameter foundParam = null;   // 診断ログ用（タイプのキャッシュ命中時は null）
                            if (isTypeParam)
                            {
                                // タイプパラメータ: (タイプID|表示名) でキャッシュ（同名区別のため表示名を使う）
                                string tkey = ElementIdToLong(elem.GetTypeId()) + "|" + headerName;
                                if (!typeCurrentCache.TryGetValue(tkey, out currentValue))
                                {
                                    var tp = ParameterService.FindParameter(elem, rawName, true, parsed.Kind, parsed.Index, paramId, doc);
                                    foundParam = tp;
                                    currentValue = ParameterService.GetParameterValueAsString(tp);
                                    paramMissing = tp == null;
                                    isReadOnly = tp == null
                                        || (tp.IsReadOnly && !ParameterService.IsTypeChangeParameter(tp));
                                    typeCurrentCache[tkey] = currentValue;
                                    typeReadOnlyCache[tkey] = isReadOnly;
                                    typeMissingCache[tkey] = paramMissing;
                                }
                                else
                                {
                                    isReadOnly = typeReadOnlyCache[tkey];
                                    paramMissing = typeMissingCache[tkey];
                                }
                            }
                            else
                            {
                                var param = ParameterService.FindParameter(elem, rawName, false, parsed.Kind, parsed.Index, paramId, doc);
                                foundParam = param;
                                currentValue = ParameterService.GetParameterValueAsString(param);
                                paramMissing = param == null;
                                // タイプ変更パラメータはIsReadOnlyでもChangeTypeIdで変更可能
                                isReadOnly = param == null
                                    || (param.IsReadOnly && !ParameterService.IsTypeChangeParameter(param));
                            }
                            bool hasChange = !ValuesAreEqual(currentValue, newValue);

                            if (hasChange)
                            {
                                string reason = !isReadOnly ? "表示"
                                    : paramMissing ? "パラメータが見つからない→非表示"
                                    : "読み取り専用→非表示";
                                diag.LogChange(worksheet.Name, row, elementIdInt, headerName,
                                    worksheet.Cell(row, i + 3), currentValue, newValue, foundParam, reason);
                            }
                            else if (currentValue != newValue)
                            {
                                // 文字としては違うが「同じ」と判定した（数値比較で一致 等）→「変更したのに出ない」原因の候補
                                diag.LogEqualButDifferent(worksheet.Name, row, elementIdInt, headerName,
                                    worksheet.Cell(row, i + 3), currentValue, newValue);
                            }

                            preview.Add(new ImportPreviewRow
                            {
                                ElementId = elementIdInt,
                                CategoryName = categoryName,
                                ParameterName = headerName,
                                CurrentValue = currentValue,
                                NewValue = newValue,
                                HasChange = hasChange,
                                IsReadOnly = isReadOnly,
                                ParamId = paramId,
                                ParamMissing = paramMissing
                            });
                        }
                    }
                }
            }

            diag.WriteSummary(preview);
            return preview;
        }

        /// <summary>
        /// 見出しから「同名パラメータの列グループ」を作る（同じ I-/T-・同じ名前・同じ種別で、番号だけ違う列）。
        /// 2列以上あるものだけを返す。各要素は列番号（paramHeaders の添字）のリスト。
        /// </summary>
        private static List<List<int>> BuildSameNameGroups(List<string> paramHeaders)
        {
            var byKey = new Dictionary<string, List<int>>();
            var order = new List<string>();
            for (int i = 0; i < paramHeaders.Count; i++)
            {
                var parsed = ParameterService.ParseDisplayName(paramHeaders[i]);
                if (parsed.Kind == null) continue;   // 番号付きの列だけが対象
                string key = (parsed.IsTypeParameter ? "T" : "I") + "|" + parsed.RawName + "|" + parsed.Kind.Value;
                if (!byKey.TryGetValue(key, out var list))
                {
                    list = new List<int>();
                    byKey[key] = list;
                    order.Add(key);
                }
                list.Add(i);
            }
            return order.Select(k => byKey[k]).Where(l => l.Count >= 2).ToList();
        }

        /// <summary>
        /// 同名パラメータの列グループを1行分まとめて判定する。
        ///
        /// 【背景】同じ名前の共有パラメータが GUID 違いで複数あると、見出しは【共有#1】【共有#2】…と分かれるが、
        /// 1つの要素が持つのは普通そのうち1つだけ。利用者が Excel 上で値を別の番号の列へ移動・統合しても、
        /// Revit 上の書き込み先はその要素が持つ1つしか無い。列ごとに判定すると
        /// 「移動元の空欄で値が消え、移動先は“パラメータなし”で書けない」＝値が失われる。
        ///
        /// 【判定】要素が持つ同名パラメータが
        ///  - 0個 → 値の入った列はすべて「パラメータなし」
        ///  - 1個 → グループ内の値を1つにまとめてそのパラメータへ書く
        ///          （値が1種類ならその値、全部空欄なら削除、違う値が混在すれば「食い違い」で書かない）
        ///  - 2個以上 → 列ごとの通常判定に任せる（false を返す）
        /// </summary>
        /// <returns>このグループを処理した（＝列ごとの通常判定は不要）なら true</returns>
        private static bool TryProcessSameNameGroup(
            Document doc, Element elem, IXLWorksheet worksheet, int row, long elementId, string categoryName,
            List<string> paramHeaders, List<int> group, bool isMergedSheet,
            List<ImportPreviewRow> preview, PreviewDiagnostics diag)
        {
            var first = ParameterService.ParseDisplayName(paramHeaders[group[0]]);
            var candidates = ParameterService.FindSameNameParameters(
                elem, first.RawName, first.IsTypeParameter, first.Kind.Value, doc);
            if (candidates.Count >= 2)
                return false;

            var cells = group
                .Select(c => new { Col = c, Header = paramHeaders[c], Value = GetCellValueAsString(worksheet.Cell(row, c + 3)) })
                .ToList();
            var nonEmpty = cells.Where(c => !string.IsNullOrEmpty(c.Value)).ToList();

            if (candidates.Count == 0)
            {
                // どの番号のパラメータも持っていない → 値の入った列はすべて「パラメータなし」
                foreach (var c in nonEmpty)
                {
                    preview.Add(new ImportPreviewRow
                    {
                        ElementId = elementId, CategoryName = categoryName, ParameterName = c.Header,
                        CurrentValue = "", NewValue = c.Value,
                        HasChange = true, IsReadOnly = true, ParamMissing = true
                    });
                    diag.LogChange(worksheet.Name, row, elementId, c.Header, worksheet.Cell(row, c.Col + 3),
                        "", c.Value, null, "同名列: 要素にパラメータなし→非表示");
                }
                diag.BlankSkipped += cells.Count - nonEmpty.Count;
                diag.BlankInGroup += cells.Count - nonEmpty.Count;
                return true;
            }

            var param = candidates[0];
            long paramId = ParameterService.ParamIdToLong(param.Id);
            string current = ParameterService.GetParameterValueAsString(param);
            bool readOnly = param.IsReadOnly && !ParameterService.IsTypeChangeParameter(param);

            // 入っている値の種類（数値として同じものは同じ値とみなす）
            var distinct = new List<string>();
            foreach (var c in nonEmpty)
                if (!distinct.Any(d => ValuesAreEqual(d, c.Value)))
                    distinct.Add(c.Value);

            if (distinct.Count > 1)
            {
                // 違う値が混在 → どれを書くか決められないので書かない（値を失わない側に倒す）
                foreach (var c in nonEmpty)
                {
                    preview.Add(new ImportPreviewRow
                    {
                        ElementId = elementId, CategoryName = categoryName, ParameterName = c.Header,
                        CurrentValue = current, NewValue = c.Value,
                        HasChange = true, IsReadOnly = true, Conflict = true, ParamId = paramId
                    });
                    diag.LogChange(worksheet.Name, row, elementId, c.Header, worksheet.Cell(row, c.Col + 3),
                        current, c.Value, param, "同名列で値が食い違い→スキップ");
                }
                return true;
            }

            string desired;
            string targetHeader;
            int targetCol;
            if (distinct.Count == 1)
            {
                desired = distinct[0];
                var hit = nonEmpty.First(c => ValuesAreEqual(c.Value, desired));
                targetHeader = hit.Header;
                targetCol = hit.Col;
            }
            else
            {
                // グループの列がすべて空欄 → 値の削除（文字のパラメータのみ。統合シートでは削除しない）
                if (isMergedSheet || param.StorageType != StorageType.String || string.IsNullOrEmpty(current))
                {
                    diag.BlankSkipped += cells.Count;
                    diag.BlankInGroup += cells.Count;
                    return true;
                }
                desired = "";
                targetHeader = cells[0].Header;
                targetCol = cells[0].Col;
            }

            bool hasChange = !ValuesAreEqual(current, desired);
            preview.Add(new ImportPreviewRow
            {
                ElementId = elementId, CategoryName = categoryName, ParameterName = targetHeader,
                CurrentValue = current, NewValue = desired,
                HasChange = hasChange, IsReadOnly = readOnly, ParamId = paramId
            });
            if (hasChange)
            {
                diag.LogChange(worksheet.Name, row, elementId, targetHeader, worksheet.Cell(row, targetCol + 3),
                    current, desired, param, readOnly ? "同名列: 読み取り専用→非表示" : (desired.Length == 0 ? "同名列: 値の削除" : "同名列: 表示"));
            }
            return true;
        }

        /// <summary>
        /// 変更プレビューの判定内容を診断ログ（C:\temp\Tools28_debug.txt）に残す。
        /// ログが肥大化しないよう、種類ごとに出力件数の上限を設ける。
        /// </summary>
        private class PreviewDiagnostics
        {
            private const int MaxChangeLogs = 300;
            private const int MaxEqualLogs = 100;

            private int _changeLogs;
            private int _equalLogs;
            public int BlankSkipped;
            public int ElementNotFound;
            public int CategoryMismatch;
            private const int MaxMismatchLogs = 30;

            public void LogCategoryMismatch(string sheet, int row, long elementId, string excelCategory, string actualCategory)
            {
                if (++CategoryMismatch > MaxMismatchLogs) return;
                DiagLog.Write($"[ImportPreview] カテゴリ不一致→取り込まない シート={sheet} 行={row} 要素={elementId} " +
                              $"Excel='{excelCategory}' モデル='{actualCategory}'");
            }
            // 空欄スキップの内訳
            public int BlankMerged, BlankNoParam, BlankNotText, BlankAlreadyEmpty, BlankInGroup;

            public void LogChange(string sheet, int row, long elementId, string header,
                IXLCell cell, string current, string newValue, Parameter param, string reason)
            {
                if (++_changeLogs > MaxChangeLogs) return;
                string storage = param != null ? param.StorageType.ToString() : "-";
                string paramId = param != null ? param.Id.IntValue().ToString() : "-";
                DiagLog.Write($"[ImportPreview] 変更判定 [{reason}] シート={sheet} 行={row} 要素={elementId} 列='{header}' " +
                    $"セル型={cell.DataType} セル生値='{RawCellText(cell)}' 新しい値='{Escape(newValue)}' 現在値='{Escape(current)}' " +
                    $"格納型={storage} パラメータID={paramId}");
            }

            public void LogEqualButDifferent(string sheet, int row, long elementId, string header,
                IXLCell cell, string current, string newValue)
            {
                if (++_equalLogs > MaxEqualLogs) return;
                DiagLog.Write($"[ImportPreview] 同値扱い（文字は不一致） シート={sheet} 行={row} 要素={elementId} 列='{header}' " +
                    $"セル型={cell.DataType} セル生値='{RawCellText(cell)}' 新しい値='{Escape(newValue)}' 現在値='{Escape(current)}'");
            }

            public void WriteSummary(List<ImportPreviewRow> preview)
            {
                int changed = preview.Count(r => r.HasChange && !r.IsReadOnly);
                int changedRo = preview.Count(r => r.HasChange && r.IsReadOnly && !r.ParamMissing && !r.Conflict);
                int changedMissing = preview.Count(r => r.HasChange && r.ParamMissing);
                int changedConflict = preview.Count(r => r.HasChange && r.Conflict);
                DiagLog.Write($"[ImportPreview] 完了: 判定 {preview.Count} 件 / 変更あり(表示) {changed} 件 / " +
                    $"変更あり(読み取り専用・非表示) {changedRo} 件 / 変更あり(パラメータなし・非表示) {changedMissing} 件 / 同名列の値の食い違い {changedConflict} 件 / 空欄スキップ {BlankSkipped} 件 / " +
                    $"同値扱い(文字不一致) {_equalLogs} 件" +
                    (_changeLogs > MaxChangeLogs ? $"（変更判定ログは先頭 {MaxChangeLogs} 件のみ出力）" : ""));
                DiagLog.Write($"[ImportPreview] 要素が見つからない {ElementNotFound} 要素 / カテゴリ不一致 {CategoryMismatch} 要素");
                DiagLog.Write($"[ImportPreview] 空欄スキップの内訳: 統合シート {BlankMerged} / パラメータなし {BlankNoParam} / " +
                    $"文字以外(数値・要素参照) {BlankNotText} / もともと空 {BlankAlreadyEmpty} / 同名列グループ {BlankInGroup}");
            }

            private static string RawCellText(IXLCell cell)
            {
                try { return Escape(cell.Value.ToString()); }
                catch { return "?"; }
            }

            /// <summary>改行・タブを見える形にする（見た目では分からない違いを確認するため）</summary>
            private static string Escape(string s)
            {
                if (s == null) return "(null)";
                return s.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
            }
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

                    // 見出し行（グループ行付きで書き出した Excel は 2行目）
                    int headerRow = ExcelHeaderNames.FindHeaderRow(worksheet);

                    if (rowCount <= headerRow || colCount < 3)
                        continue;

                    // ヘッダーからパラメータ名を取得（(*変更不可)サフィックスは除去）
                    var paramHeaders = new List<string>();
                    for (int col = 3; col <= colCount; col++)
                    {
                        string header = worksheet.Cell(headerRow, col).GetString();
                        paramHeaders.Add(StripReadOnlySuffix(header));
                    }

                    // データ行を処理
                    for (int row = headerRow + 1; row <= rowCount; row++)
                    {
                        string elementIdStr = worksheet.Cell(row, 1).GetString();

                        if (!TryParseElementId(elementIdStr, out long elementIdInt))
                        {
                            result.FailCount++;
                            result.Errors.Add($"シート '{worksheet.Name}' 行{row}: ElementIdの解析に失敗");
                            continue;
                        }

                        var elementId = ToElementId(elementIdInt);
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

                            // 空セルはスキップ（シート統合モードで該当カテゴリに存在しないパラメータ列）
                            if (string.IsNullOrEmpty(newValue))
                            {
                                result.SkipCount++;
                                continue;
                            }

                            var parsed = ParameterService.ParseDisplayName(headerName);
                            bool isTypeParam = parsed.IsTypeParameter;
                            string rawName = parsed.RawName;

                            var param = ParameterService.FindParameter(elem, rawName, isTypeParam, parsed.Kind, parsed.Index, doc);

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
                                success = ParameterService.SetParameterValue(param, newValue, doc);
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
        /// プレビュー結果を使って「変更あり かつ 書込み可能」なセルのみをインポートする。
        /// Excel の再読込・全セル走査を行わないため、大容量データでも高速。
        /// （プレビューは <see cref="GeneratePreview"/> で生成済みの値を再利用する）
        /// </summary>
        public static ImportResult ImportFromPreview(Document doc, List<ImportPreviewRow> previewRows)
        {
            return ImportFromPreview(doc, previewRows, null);
        }

        /// <summary>
        /// <see cref="ImportFromPreview(Document, List{ImportPreviewRow})"/> の除外指定付き版。
        /// excludeElementIds に含まれる要素は書き込み対象から外す。
        /// （制約エラー（部屋の高さ&gt;0 等）で失敗した要素を除いて再実行するために使う）
        /// </summary>
        public static ImportResult ImportFromPreview(
            Document doc, List<ImportPreviewRow> previewRows, HashSet<long> excludeElementIds)
        {
            var result = new ImportResult();
            if (previewRows == null)
                return result;

            // 読み取り専用で変更できないセル・パラメータが無いセルはスキップ扱いで集計
            result.SkipReadOnly = previewRows.Count(r => r.HasChange && r.IsReadOnly && !r.ParamMissing && !r.Conflict);
            result.SkipNotFound = previewRows.Count(r => r.HasChange && r.ParamMissing);
            result.SkipConflict = previewRows.Count(r => r.HasChange && r.Conflict);

            // 実際に書き込む対象（変更あり かつ 書込み可能、除外指定を除く）だけを処理
            foreach (var pr in previewRows.Where(r => r.HasChange && !r.IsReadOnly
                        && (excludeElementIds == null || !excludeElementIds.Contains(r.ElementId))))
            {
                var elem = doc.GetElement(ToElementId(pr.ElementId));
                if (elem == null)
                {
                    result.FailCount++;
                    result.Errors.Add($"要素 {pr.ElementId} が見つかりません（パラメータ '{pr.ParameterName}'）");
                    continue;
                }

                string headerName = pr.ParameterName;
                var parsed = ParameterService.ParseDisplayName(headerName);
                bool isTypeParam = parsed.IsTypeParameter;
                string rawName = parsed.RawName;

                // プレビューと同じく、書き出し時の識別番号があればそれでパラメータを特定する
                var param = ParameterService.FindParameter(elem, rawName, isTypeParam, parsed.Kind, parsed.Index, pr.ParamId, doc);
                if (param == null)
                {
                    result.Warnings.Add($"パラメータ '{headerName}' が見つかりません（要素 {pr.ElementId}）");
                    result.SkipNotFound++;
                    continue;
                }

                // 読み取り専用（タイプ変更パラメータは除く）は書込み不可
                if (param.IsReadOnly && !ParameterService.IsTypeChangeParameter(param))
                {
                    result.SkipReadOnly++;
                    continue;
                }

                // プレビュー生成後にモデルが変わった場合に備えて現在値を再確認
                // （同じタイプの複数行で同じタイプパラメータを変えた場合、2行目以降はここでスキップになる）
                string currentValue = ParameterService.GetParameterValueAsString(param);
                if (ValuesAreEqual(currentValue, pr.NewValue))
                {
                    result.SkipUnchanged++;
                    DiagLog.Write($"[Import] skip(既に同じ値) elem={pr.ElementId} param='{headerName}' current='{currentValue}'");
                    continue;
                }

                string bip = "";
                try { bip = param.Definition is InternalDefinition intDef ? intDef.BuiltInParameter.ToString() : "(shared/custom)"; }
                catch { }
                DiagLog.Write($"[Import] elem={pr.ElementId} param='{headerName}' storage={param.StorageType} " +
                    $"shared={param.IsShared} readonly={param.IsReadOnly} bip={bip} " +
                    $"isType={ParameterService.IsTypeChangeParameter(param)} current='{currentValue}' new='{pr.NewValue}'");

                bool success = ParameterService.IsTypeChangeParameter(param)
                    ? ParameterService.ChangeElementType(elem, pr.NewValue, doc)
                    : ParameterService.SetParameterValue(param, pr.NewValue, doc);

                DiagLog.Write($"[Import] -> result={success}");

                if (success)
                {
                    result.SuccessCount++;
                }
                else
                {
                    result.FailCount++;
                    result.FailedCells.Add(pr.ElementId.ToString() + "|" + headerName);
                    result.Errors.Add(BuildSetFailureMessage(param, headerName, pr.ElementId, pr.NewValue, doc));
                }
            }

            result.SkipCount = result.SkipReadOnly + result.SkipNotFound + result.SkipUnchanged + result.SkipConflict;
            return result;
        }

        /// <summary>取り込めなかったセルの塗りつぶし色（オレンジ）。COM 経由の色付けと共通。</summary>
        internal const int SkippedR = 255, SkippedG = 192, SkippedB = 120;

        /// <summary>
        /// インポートで変更されたセルにExcelファイル上で色を付ける
        /// Excelが開いている場合はCOM経由で直接色付け、閉じている場合はClosedXMLで上書き
        /// </summary>
        /// <returns>色付けファイルの保存先パス。COM経由成功時は元ファイルパス。色付け不要/失敗の場合はnull</returns>
        public static string MarkImportedCells(string filePath, List<ImportPreviewRow> previewRows, out string colorMethod, HashSet<string> failedSet = null)
        {
            colorMethod = null;

            // 変更・追加（値あり）で成功したセル → 文字を青字にする
            var changedSet = new HashSet<string>(
                previewRows
                    .Where(r => r.HasChange && !r.IsReadOnly && !string.IsNullOrEmpty(r.NewValue))
                    .Select(r => r.ElementId.ToString() + "|" + r.ParameterName));

            // 値を削除（空欄化）して成功したセル → セルを青で塗りつぶす（文字が無いため）
            var clearedSet = new HashSet<string>(
                previewRows
                    .Where(r => r.HasChange && !r.IsReadOnly && string.IsNullOrEmpty(r.NewValue))
                    .Select(r => r.ElementId.ToString() + "|" + r.ParameterName));

            // 成功セットから失敗セルを除外
            if (failedSet != null && failedSet.Count > 0)
            {
                changedSet.ExceptWith(failedSet);
                clearedSet.ExceptWith(failedSet);
            }

            // 取り込めなかったセル（読み取り専用・パラメータなし・同名列の値の食い違い）→ オレンジで塗りつぶす。
            // 色が付かないと「反映されたのか分からない」ため、取り込まなかったことを Excel 上で見えるようにする。
            var skippedSet = new HashSet<string>(
                previewRows
                    .Where(r => r.HasChange && r.IsReadOnly)
                    .Select(r => r.ElementId.ToString() + "|" + r.ParameterName));

            if (changedSet.Count == 0 && clearedSet.Count == 0 && skippedSet.Count == 0
                && (failedSet == null || failedSet.Count == 0))
                return null;

            // まずCOM経由（開いているExcelに直接色付け）を試行
            if (ExcelProcessHelper.MarkCellsViaCom(filePath, changedSet, clearedSet, failedSet, skippedSet))
            {
                colorMethod = "COM";
                return filePath;
            }

            // Excelが開いていない or COM失敗の場合、ClosedXMLでファイルを直接編集
            colorMethod = "ClosedXML";
            return MarkCellsViaClosedXml(filePath, changedSet, clearedSet, failedSet, skippedSet);
        }

        /// <summary>
        /// インポートで変更されたセルにExcelファイル上で色を付ける（互換用オーバーロード）
        /// </summary>
        public static string MarkImportedCells(string filePath, List<ImportPreviewRow> previewRows)
        {
            return MarkImportedCells(filePath, previewRows, out _);
        }

        /// <summary>
        /// Excel セルの文字列から要素 Id を読む。
        /// 数値セルが指数表記（1.23E+09）で返る場合にも対応する。
        /// </summary>
        private static bool TryParseElementId(string text, out long id)
        {
            id = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;

            text = text.Trim();
            if (long.TryParse(text, out id)) return true;

            if (double.TryParse(text, System.Globalization.NumberStyles.Float,
                                System.Globalization.CultureInfo.InvariantCulture, out double d))
            {
                id = (long)d;
                return true;
            }
            return false;
        }

        /// <summary>
        /// long から ElementId を作る（バージョン差異を吸収）。
        /// Revit 2025 以前は要素 Id が int の範囲に収まる。
        /// </summary>
        private static ElementId ToElementId(long value)
        {
#if REVIT2026
            return new ElementId(value);
#else
            return new ElementId((int)value);
#endif
        }

        /// <summary>
        /// ElementId を long に変換（バージョン差異を吸収）
        /// </summary>
        private static long ElementIdToLong(ElementId id)
        {
            if (id == null) return -1;
#if REVIT2026
            return id.Value;
#else
            return id.IntValue();
#endif
        }

        /// <summary>
        /// 値設定に失敗した理由を、パラメータ型に応じて分かりやすいメッセージにする。
        /// 特に ElementId（要素参照）型は文字値を直接設定できないことを明示する。
        /// </summary>
        private static string BuildSetFailureMessage(Parameter param, string headerName, long elementId, string value, Document doc)
        {
            if (ParameterService.IsTypeChangeParameter(param))
                return $"タイプ変更に失敗（要素 {elementId}, 値: '{value}'）— 一致するタイプが見つかりません";

            if (param.StorageType == StorageType.ElementId)
            {
                bool isImage = false;
                try
                {
                    isImage = param.Definition is InternalDefinition idf
                        && (idf.BuiltInParameter == BuiltInParameter.ALL_MODEL_IMAGE
                            || idf.BuiltInParameter == BuiltInParameter.ALL_MODEL_TYPE_IMAGE);
                }
                catch { }

                if (isImage)
                    return $"パラメータ '{headerName}' は画像参照（イメージ）型のため文字値 '{value}' は設定できません" +
                           $"（要素 {elementId}）。設定するにはその名前の画像がプロジェクトに存在する必要があります。";

                string msg = $"パラメータ '{headerName}' は要素参照型のため文字値 '{value}' は設定できません" +
                             $"（要素 {elementId}）。'{value}' という名前の要素が見つかりません。";

                // 似た名前の候補（例: 「7FL」→「6階(7FL)」）を添える。自動では置き換えない
                var candidates = ParameterService.SuggestReferenceNames(param, value, doc);
                if (candidates.Count > 0)
                {
                    msg += string.Format(Loc.S("Import.RefCandidates"), string.Join(" / ", candidates));
                    DiagLog.Write($"[Import] 名前の候補 '{value}' -> {string.Join(" / ", candidates)}");
                }
                return msg;
            }

            return $"パラメータ '{headerName}' の値設定に失敗（要素 {elementId}, 値: '{value}'）";
        }

        /// <summary>
        /// ヘッダー名から編集可否マーカー（変更不可/画像参照/要素参照）を除去する
        /// </summary>
        private static string StripReadOnlySuffix(string headerName)
        {
            return ParameterHeaderMarker.Strip(headerName);
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
        private static string MarkCellsViaClosedXml(string filePath, HashSet<string> changedSet, HashSet<string> clearedSet = null, HashSet<string> failedSet = null, HashSet<string> skippedSet = null)
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

            // 成功セル: 青(R79,G129,B189)/太字、失敗セル: 赤/太字、取り込めなかったセル: オレンジ塗り
            var blueColor = XLColor.FromArgb(79, 129, 189);
            var orangeColor = XLColor.FromArgb(SkippedR, SkippedG, SkippedB);

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

                    // 見出し行（グループ行付きで書き出した Excel は 2行目）
                    int headerRow = ExcelHeaderNames.FindHeaderRow(worksheet);

                    if (rowCount <= headerRow || colCount < 3)
                        continue;

                    var paramHeaders = new List<string>();
                    for (int col = 3; col <= colCount; col++)
                    {
                        paramHeaders.Add(StripReadOnlySuffix(worksheet.Cell(headerRow, col).GetString()));
                    }

                    for (int row = headerRow + 1; row <= rowCount; row++)
                    {
                        // 要素 Id の解釈は他の経路と同じヘルパーに統一する
                        // （旧実装はここだけ double 経由で int に丸めていた）。
                        string elementIdStr = worksheet.Cell(row, 1).GetString();
                        if (TryParseElementId(elementIdStr, out long idValue))
                            elementIdStr = idValue.ToString();
                        else
                            elementIdStr = elementIdStr.Trim();

                        // セル単位で成功(変更)/削除(クリア)/失敗を判定
                        var successCols = new HashSet<int>();
                        var clearedCols = new HashSet<int>();
                        var failedCols = new HashSet<int>();
                        var skippedCols = new HashSet<int>();
                        for (int i = 0; i < paramHeaders.Count; i++)
                        {
                            string key = elementIdStr + "|" + paramHeaders[i];
                            if (clearedSet != null && clearedSet.Contains(key))
                                clearedCols.Add(i + 3);
                            else if (changedSet.Contains(key))
                                successCols.Add(i + 3);
                            else if (failedSet != null && failedSet.Contains(key))
                                failedCols.Add(i + 3);
                            else if (skippedSet != null && skippedSet.Contains(key))
                                skippedCols.Add(i + 3);
                        }

                        // 変更・削除・失敗・取り込めなかったセルがある行は全列に背景色、セル単位で色分け
                        if (successCols.Count > 0 || clearedCols.Count > 0 || failedCols.Count > 0 || skippedCols.Count > 0)
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
                            // 削除（空欄）セルは文字が無いため、セルを青で塗りつぶす
                            foreach (int col in clearedCols)
                            {
                                worksheet.Cell(row, col).Style.Fill.BackgroundColor = blueColor;
                            }
                            foreach (int col in failedCols)
                            {
                                worksheet.Cell(row, col).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(row, col).Style.Font.Bold = true;
                            }
                            foreach (int col in skippedCols)
                            {
                                worksheet.Cell(row, col).Style.Fill.BackgroundColor = orangeColor;
                            }
                            anyMarked = true;
                        }
                    }
                }

                if (!anyMarked)
                    return null;

                // 各シートの見出し行（最終列の次）に凡例を追加
                foreach (var worksheet in workbook.Worksheets)
                {
                    if (ParameterIdSheet.IsMetaSheet(worksheet))
                        continue; // 識別番号の隠しシートには凡例を付けない

                    var lastCol = worksheet.LastColumnUsed();
                    if (lastCol == null) continue;
                    int legendCol = lastCol.ColumnNumber() + 1;

                    var legendCell = worksheet.Cell(ExcelHeaderNames.FindHeaderRow(worksheet), legendCol);
                    var richText = legendCell.CreateRichText();
                    richText.AddText("(*");
                    var bluePart = richText.AddText("青字・青セル");
                    bluePart.SetFontColor(blueColor);
                    bluePart.SetBold(true);
                    richText.AddText("はインポート成功（青セルは値の削除）、");
                    var redPart = richText.AddText("赤字");
                    redPart.SetFontColor(XLColor.Red);
                    redPart.SetBold(true);
                    richText.AddText("はインポート失敗、");
                    var orangePart = richText.AddText("オレンジのセル");
                    orangePart.SetFontColor(orangeColor);
                    orangePart.SetBold(true);
                    richText.AddText("は取り込めなかった値（読み取り専用・パラメータなし・同名列の値の食い違い）)");
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

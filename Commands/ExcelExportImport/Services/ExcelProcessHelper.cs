using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// 開いているExcelファイルを検出・操作するヘルパー
    /// P/Invoke で oleaut32.dll の GetActiveObject を直接呼び出すことで、
    /// .NET Framework 4.8 (Revit 2021-2024) と .NET 8 (Revit 2025-2026) の両方で動作する
    /// </summary>
    public static class ExcelProcessHelper
    {
        // P/Invoke: oleaut32.dll の GetActiveObject（Marshal.GetActiveObject の内部実装と同等）
        // .NET 8 では Marshal.GetActiveObject が削除されたため、直接呼び出す
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(
            ref Guid rclsid,
            IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
            out Guid lpclsid);

        /// <summary>
        /// 現在Excelで開いているxlsxファイルのパス一覧を取得
        /// </summary>
        public static List<string> GetOpenExcelFiles()
        {
            var result = new List<string>();

            try
            {
                dynamic app = GetExcelApplication();
                if (app == null)
                    return result;

                try
                {
                    dynamic workbooks = app.Workbooks;
                    int count = workbooks.Count;

                    for (int i = 1; i <= count; i++)
                    {
                        dynamic wb = workbooks[i];
                        try
                        {
                            string fullName = wb.FullName;
                            // 拡張子はブック名で判定する（クラウド上のブックは FullName が URL で、
                            // 末尾に "?web=1" などが付くことがあるため）
                            string wbName = wb.Name;
                            DiagLog.Write($"[ExcelImport] 開いているブック: {wbName} / {fullName}");
                            if (wbName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                wbName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                            {
                                result.Add(fullName);
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(wb);
                        }
                    }

                    Marshal.ReleaseComObject(workbooks);
                }
                finally
                {
                    Marshal.ReleaseComObject(app);
                }
            }
            catch
            {
                // COM関連のエラーは無視
            }

            return result;
        }

        /// <summary>凡例セルの中の指定文字列部分を、指定色の太字にする（COM 経由）</summary>
        private static void ColorLegendPart(dynamic legendCell, string legendText, string part, int color)
        {
            int idx = legendText.IndexOf(part, StringComparison.Ordinal);
            if (idx < 0) return;
            dynamic chars = legendCell.Characters[idx + 1, part.Length];
            chars.Font.Color = color;
            chars.Font.Bold = true;
            Marshal.ReleaseComObject(chars);
        }

        /// <summary>
        /// 開いているExcelブックの指定セルに背景色を設定する（COM経由）
        /// </summary>
        /// <param name="filePath">対象ファイルパス</param>
        /// <param name="changedSet">変更・追加で成功したセルのキー（"ElementId|ParameterName"）→ 青字</param>
        /// <param name="clearedSet">値を削除（空欄化）して成功したセルのキー → 青塗り</param>
        /// <param name="failedSet">失敗したセルのキー → 赤字</param>
        /// <returns>色付けに成功した場合true</returns>
        public static bool MarkCellsViaCom(string filePath, HashSet<string> changedSet, HashSet<string> clearedSet = null, HashSet<string> failedSet = null, HashSet<string> skippedSet = null)
        {
            if ((changedSet == null || changedSet.Count == 0)
                && (clearedSet == null || clearedSet.Count == 0)
                && (failedSet == null || failedSet.Count == 0)
                && (skippedSet == null || skippedSet.Count == 0))
                return false;

            dynamic app = null;
            try
            {
                app = GetExcelApplication();
                if (app == null)
                    return false;

                // ファイルパスに一致するワークブックを検索（パスを正規化して比較）
                dynamic workbooks = app.Workbooks;
                dynamic targetWb = null;
                int wbCount = workbooks.Count;
                string normalizedFilePath = NormalizePath(filePath);

                // まずフルパス完全一致で検索
                for (int i = 1; i <= wbCount; i++)
                {
                    dynamic wb = workbooks[i];
                    try
                    {
                        string fullName = wb.FullName;
                        if (string.Equals(NormalizePath(fullName), normalizedFilePath, StringComparison.OrdinalIgnoreCase))
                        {
                            targetWb = wb;
                            break;
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                    if (targetWb == null || !ReferenceEquals(targetWb, wb))
                    {
                        Marshal.ReleaseComObject(wb);
                    }
                }

                // フルパスで見つからない場合、ファイル名のみで再検索
                // （OneDriveパス仮想化等でパスが異なる場合への対応）
                if (targetWb == null)
                {
                    string targetFileName = System.IO.Path.GetFileName(filePath);
                    for (int i = 1; i <= wbCount; i++)
                    {
                        dynamic wb = workbooks[i];
                        try
                        {
                            string wbName = wb.Name;
                            if (string.Equals(wbName, targetFileName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetWb = wb;
                                break;
                            }
                        }
                        catch
                        {
                            // ignore
                        }
                        if (targetWb == null || !ReferenceEquals(targetWb, wb))
                        {
                            Marshal.ReleaseComObject(wb);
                        }
                    }
                }

                Marshal.ReleaseComObject(workbooks);

                if (targetWb == null)
                    return false;

                bool anyMarked = false;
                try
                {
                    // 画面更新・イベントを一時停止（COM往復削減＋最後に一括反映）
                    try { app.ScreenUpdating = false; } catch { }
                    try { app.EnableEvents = false; } catch { }

                    // Excel COM の Interior.Color は R + G*256 + B*65536 形式
                    // R=255, G=255, B=153
                    int excelColor = 255 + 255 * 256 + 153 * 256 * 256;

                    // Excel COM の Font.Color は R + G*256 + B*65536 形式
                    // 青色: R=79, G=129, B=189
                    int blueColor = 79 + 129 * 256 + 189 * 256 * 256;
                    // 赤色: R=255, G=0, B=0（失敗セル用）
                    int redColor = 255 + 0 * 256 + 0 * 256 * 256;
                    // オレンジ（取り込めなかったセル用。ClosedXML 側と同じ色）
                    int orangeColor = ExcelImportService.SkippedR
                                      + ExcelImportService.SkippedG * 256
                                      + ExcelImportService.SkippedB * 256 * 256;

                    int sheetCount = targetWb.Sheets.Count;
                    for (int s = 1; s <= sheetCount; s++)
                    {
                        dynamic sheet = targetWb.Sheets[s];
                        try
                        {
                            // パラメータ識別番号の隠しシートは色付け対象外
                            string sheetName = Convert.ToString((object)sheet.Name);
                            if (ParameterIdSheet.IsMetaSheetName(sheetName))
                                continue;

                            // 使用範囲を取得（開始行・列も考慮）
                            dynamic usedRange = sheet.UsedRange;
                            int startRow = (int)usedRange.Row;
                            int startCol = (int)usedRange.Column;
                            int rowCount = startRow + (int)usedRange.Rows.Count - 1;
                            int colCount = startCol + (int)usedRange.Columns.Count - 1;
                            Marshal.ReleaseComObject(usedRange);

                            if (rowCount < 2 || colCount < 3)
                                continue;

                            // ヘッダー行とID列を「一括」で読み取る（セル単位のCOM往復を回避）
                            // ※ 従来は全セルを1つずつ読んでいたため、大きな表で膨大なCOM呼び出しとなりフリーズしていた
                            // ID列は1行目から読み、見出し行（グループ行付きなら2行目）を判定する
                            var idValues = ReadColumnValues(sheet, 1, rowCount, 1);
                            int headerRow = ExcelHeaderNames.FindHeaderRow(
                                Convert.ToString(idValues[0] ?? ""),
                                Convert.ToString(idValues[1] ?? ""));
                            if (rowCount <= headerRow)
                                continue;
                            var paramHeaders = ReadRowValues(sheet, headerRow, 3, colCount);

                            for (int i = 0; i < paramHeaders.Count; i++)
                            {
                                // 編集可否マーカー（変更不可/画像参照/要素参照）を除去して素の名前に合わせる
                                paramHeaders[i] = ParameterHeaderMarker.Strip(paramHeaders[i]);
                            }

                            // データ行を走査（メモリ上で判定し、色付けが必要な行だけCOM操作）
                            for (int row = headerRow + 1; row <= rowCount; row++)
                            {
                                object idValue = idValues[row - 1];
                                if (idValue == null)
                                    continue;

                                // ElementIdを整数文字列に正規化
                                string elementIdStr;
                                if (idValue is double dVal)
                                    elementIdStr = ((int)dVal).ToString();
                                else
                                    elementIdStr = Convert.ToString(idValue).Trim();

                                // セル単位で成功(変更)/削除(クリア)/失敗を判定
                                var successCols = new List<int>();
                                var clearedCols = new List<int>();
                                var failedCols = new List<int>();
                                var skippedCols = new List<int>();
                                for (int i = 0; i < paramHeaders.Count; i++)
                                {
                                    string key = elementIdStr + "|" + paramHeaders[i];
                                    if (clearedSet != null && clearedSet.Contains(key))
                                        clearedCols.Add(i + 3);
                                    else if (changedSet != null && changedSet.Contains(key))
                                        successCols.Add(i + 3);
                                    else if (failedSet != null && failedSet.Contains(key))
                                        failedCols.Add(i + 3);
                                    else if (skippedSet != null && skippedSet.Contains(key))
                                        skippedCols.Add(i + 3);
                                }

                                if (successCols.Count == 0 && clearedCols.Count == 0 && failedCols.Count == 0 && skippedCols.Count == 0)
                                    continue;

                                try
                                {
                                    // 背景色: 行全体を1回のCOM呼び出しで着色
                                    dynamic c1 = sheet.Cells[row, 1];
                                    dynamic c2 = sheet.Cells[row, colCount];
                                    dynamic rowRange = sheet.Range[c1, c2];
                                    rowRange.Interior.Color = excelColor;
                                    Marshal.ReleaseComObject(rowRange);
                                    Marshal.ReleaseComObject(c2);
                                    Marshal.ReleaseComObject(c1);

                                    // フォント色: 変更のあったセルのみ（通常は少数）
                                    foreach (int col in successCols)
                                    {
                                        dynamic cell = sheet.Cells[row, col];
                                        cell.Font.Color = blueColor;
                                        cell.Font.Bold = true;
                                        Marshal.ReleaseComObject(cell);
                                    }
                                    // 削除（空欄）セルは文字が無いため、セルを青で塗りつぶす（行の黄色を上書き）
                                    foreach (int col in clearedCols)
                                    {
                                        dynamic cell = sheet.Cells[row, col];
                                        cell.Interior.Color = blueColor;
                                        Marshal.ReleaseComObject(cell);
                                    }
                                    foreach (int col in failedCols)
                                    {
                                        dynamic cell = sheet.Cells[row, col];
                                        cell.Font.Color = redColor;
                                        cell.Font.Bold = true;
                                        Marshal.ReleaseComObject(cell);
                                    }
                                    // 取り込めなかったセル（読み取り専用・パラメータなし・値の食い違い）はオレンジで塗る
                                    foreach (int col in skippedCols)
                                    {
                                        dynamic cell = sheet.Cells[row, col];
                                        cell.Interior.Color = orangeColor;
                                        Marshal.ReleaseComObject(cell);
                                    }
                                    anyMarked = true;
                                }
                                catch
                                {
                                    // 個別行の色付け失敗は無視して次の行へ
                                }
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(sheet);
                        }
                    }

                    // 各シートの見出し行（最終列の次）に凡例を追加
                    if (anyMarked)
                    {
                        for (int s2 = 1; s2 <= sheetCount; s2++)
                        {
                            dynamic sheet2 = targetWb.Sheets[s2];
                            try
                            {
                                if (ParameterIdSheet.IsMetaSheetName(Convert.ToString((object)sheet2.Name)))
                                    continue; // 識別番号の隠しシートには凡例を付けない

                                dynamic usedRange2 = sheet2.UsedRange;
                                int lastColNum = (int)usedRange2.Column + (int)usedRange2.Columns.Count - 1;
                                Marshal.ReleaseComObject(usedRange2);

                                int legendCol = lastColNum + 1;
                                // 凡例は見出し行（グループ行付きなら2行目）に置く
                                var col1Head = ReadColumnValues(sheet2, 1, 2, 1);
                                int legendRow = ExcelHeaderNames.FindHeaderRow(
                                    Convert.ToString(col1Head[0] ?? ""),
                                    Convert.ToString(col1Head[1] ?? ""));
                                dynamic legendCell = sheet2.Cells[legendRow, legendCol];
                                string legendText = "(*青字・青セルはインポート成功（青セルは値の削除）、赤字はインポート失敗、" +
                                                    "オレンジのセルは取り込めなかった値（読み取り専用・パラメータなし・同名列の値の食い違い）)";
                                legendCell.Value = legendText;

                                // 色の説明部分を、その色の太字にする（Characters の開始位置は 1 始まり）
                                ColorLegendPart(legendCell, legendText, "青字・青セル", blueColor);
                                ColorLegendPart(legendCell, legendText, "赤字", redColor);
                                ColorLegendPart(legendCell, legendText, "オレンジのセル", orangeColor);

                                Marshal.ReleaseComObject(legendCell);
                            }
                            catch
                            {
                                // 凡例追加失敗は無視
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(sheet2);
                            }
                        }
                    }

                    return anyMarked;
                }
                finally
                {
                    // 画面更新・イベントを再開（これにより色付けが画面に反映される）
                    try { app.EnableEvents = true; } catch { }
                    try { app.ScreenUpdating = true; } catch { }
                    Marshal.ReleaseComObject(targetWb);
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                if (app != null)
                {
                    try { Marshal.ReleaseComObject(app); } catch { }
                }
            }
        }

        /// <summary>
        /// Excel で開いているブック（FullName 一致）の現在の内容を、一時フォルダに複製する（COM 経由）。
        /// クラウド上のブックで同期フォルダの実ファイルが見つからない場合の読み込み用。
        /// ファイル名は元のブックと同じにする（インポート後の色付けでブックを名前で探すため）。
        /// </summary>
        /// <returns>複製したファイルのパス。失敗した場合は null</returns>
        public static string SaveOpenWorkbookCopy(string fullName)
            => SaveOpenWorkbookCopy(fullName, matchByPath: false);

        /// <summary>
        /// 指定ファイルが Excel で開かれていて、保存されていない変更があるかを調べる（COM 経由）。
        /// インポートはディスク上のファイルを読むため、未保存の編集はそのままでは取り込まれない。
        /// </summary>
        /// <returns>開いていて未保存の変更がある場合 true（開いていない・判定できない場合は false）</returns>
        public static bool HasUnsavedChanges(string filePath)
        {
            dynamic app = null;
            dynamic workbooks = null;
            try
            {
                app = GetExcelApplication();
                if (app == null)
                    return false;

                string target = NormalizePath(filePath);
                workbooks = app.Workbooks;
                int count = workbooks.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic wb = workbooks[i];
                    try
                    {
                        string fullName = wb.FullName;
                        if (!string.Equals(NormalizePath(fullName), target, StringComparison.OrdinalIgnoreCase))
                            continue;
                        bool saved = wb.Saved;
                        return !saved;
                    }
                    catch
                    {
                        // クラウド上のブック等で判定できない場合は次へ
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(wb);
                    }
                }
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[ExcelImport] 未保存の判定に失敗: {ex.Message}");
            }
            finally
            {
                if (workbooks != null) Marshal.ReleaseComObject(workbooks);
                if (app != null) Marshal.ReleaseComObject(app);
            }
            return false;
        }

        /// <summary>
        /// Excel で開いているブックの現在の内容（未保存の編集を含む）を一時フォルダに複製する。
        /// matchByPath=true ならローカルパスを正規化して照合する（FullName の表記ゆれ対策）。
        /// </summary>
        public static string SaveOpenWorkbookCopy(string fullName, bool matchByPath)
        {
            dynamic app = null;
            dynamic workbooks = null;
            try
            {
                app = GetExcelApplication();
                if (app == null)
                    return null;

                workbooks = app.Workbooks;
                int count = workbooks.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic wb = workbooks[i];
                    try
                    {
                        string wbFullName = wb.FullName;
                        bool match = matchByPath
                            ? string.Equals(NormalizePath(wbFullName), NormalizePath(fullName), StringComparison.OrdinalIgnoreCase)
                            : string.Equals(wbFullName, fullName, StringComparison.OrdinalIgnoreCase);
                        if (!match)
                            continue;

                        string dir = Path.Combine(Path.GetTempPath(), "Tools28", "ExcelImport");
                        Directory.CreateDirectory(dir);
                        string copyPath = Path.Combine(dir, (string)wb.Name);
                        if (File.Exists(copyPath))
                            File.Delete(copyPath);

                        wb.SaveCopyAs(copyPath);
                        return File.Exists(copyPath) ? copyPath : null;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(wb);
                    }
                }
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[ExcelImport] ブックの複製に失敗: {ex.Message}");
            }
            finally
            {
                if (workbooks != null) Marshal.ReleaseComObject(workbooks);
                if (app != null) Marshal.ReleaseComObject(app);
            }
            return null;
        }

        /// <summary>
        /// Excel.Application の COM オブジェクトを取得
        /// oleaut32.dll の GetActiveObject を直接呼び出す
        /// </summary>
        private static dynamic GetExcelApplication()
        {
            try
            {
                Guid clsid;
                int hr = CLSIDFromProgID("Excel.Application", out clsid);
                if (hr != 0)
                    return null;

                object app;
                GetActiveObject(ref clsid, IntPtr.Zero, out app);
                return app;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 1行分のセル値を一括読み取り（col1..col2 を1回のCOM呼び出しで取得）
        /// </summary>
        private static List<string> ReadRowValues(dynamic sheet, int row, int col1, int col2)
        {
            var list = new List<string>();
            dynamic c1 = sheet.Cells[row, col1];
            dynamic c2 = sheet.Cells[row, col2];
            dynamic range = sheet.Range[c1, c2];
            object val = range.Value2;
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(c2);
            Marshal.ReleaseComObject(c1);

            if (val is object[,] arr)
            {
                int lo1 = arr.GetLowerBound(0);
                int lo2 = arr.GetLowerBound(1);
                int cols = arr.GetLength(1);
                for (int c = 0; c < cols; c++)
                    list.Add(Convert.ToString(arr[lo1, lo2 + c] ?? ""));
            }
            else
            {
                // 単一セル（col1 == col2）
                list.Add(Convert.ToString(val ?? ""));
            }
            return list;
        }

        /// <summary>
        /// 1列分のセル値を一括読み取り（row1..row2 を1回のCOM呼び出しで取得）
        /// </summary>
        private static object[] ReadColumnValues(dynamic sheet, int row1, int row2, int col)
        {
            int n = row2 - row1 + 1;
            var values = new object[n];
            dynamic c1 = sheet.Cells[row1, col];
            dynamic c2 = sheet.Cells[row2, col];
            dynamic range = sheet.Range[c1, c2];
            object val = range.Value2;
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(c2);
            Marshal.ReleaseComObject(c1);

            if (val is object[,] arr)
            {
                int lo1 = arr.GetLowerBound(0);
                int lo2 = arr.GetLowerBound(1);
                for (int r = 0; r < n; r++)
                    values[r] = arr[lo1 + r, lo2];
            }
            else
            {
                // 単一セル（row1 == row2）
                values[0] = val;
            }
            return values;
        }

        /// <summary>
        /// ファイルパスを正規化して比較しやすくする
        /// </summary>
        private static string NormalizePath(string path)
        {
            try
            {
                return Path.GetFullPath(path).TrimEnd(
                    Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
            catch
            {
                return path;
            }
        }
    }
}

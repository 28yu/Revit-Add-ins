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
                            if (fullName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                fullName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
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

        /// <summary>
        /// 開いているExcelブックの指定セルに背景色を設定する（COM経由）
        /// </summary>
        /// <param name="filePath">対象ファイルパス</param>
        /// <param name="changedSet">色付け対象セルのキー（"ElementId|ParameterName"）</param>
        /// <returns>色付けに成功した場合true</returns>
        public static bool MarkCellsViaCom(string filePath, HashSet<string> changedSet, HashSet<string> failedSet = null)
        {
            if ((changedSet == null || changedSet.Count == 0) && (failedSet == null || failedSet.Count == 0))
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
                    // 画面更新を一時停止（パフォーマンス向上＋最後に一括更新）
                    try { app.ScreenUpdating = false; } catch { }

                    // Excel COM の Interior.Color は R + G*256 + B*65536 形式
                    // R=255, G=255, B=153
                    int excelColor = 255 + 255 * 256 + 153 * 256 * 256;

                    // Excel COM の Font.Color は R + G*256 + B*65536 形式
                    // 青色: R=79, G=129, B=189
                    int blueColor = 79 + 129 * 256 + 189 * 256 * 256;
                    // 赤色: R=255, G=0, B=0（失敗セル用）
                    int redColor = 255 + 0 * 256 + 0 * 256 * 256;

                    int sheetCount = targetWb.Sheets.Count;
                    for (int s = 1; s <= sheetCount; s++)
                    {
                        dynamic sheet = targetWb.Sheets[s];
                        try
                        {
                            // 使用範囲を取得（開始行・列も考慮）
                            dynamic usedRange = sheet.UsedRange;
                            int startRow = (int)usedRange.Row;
                            int startCol = (int)usedRange.Column;
                            int rowCount = startRow + (int)usedRange.Rows.Count - 1;
                            int colCount = startCol + (int)usedRange.Columns.Count - 1;
                            Marshal.ReleaseComObject(usedRange);

                            if (rowCount < 2 || colCount < 3)
                                continue;

                            // ヘッダー行からパラメータ名を取得（(*変更不可)サフィックスは除去）
                            var paramHeaders = new List<string>();
                            for (int col = 3; col <= colCount; col++)
                            {
                                dynamic cell = sheet.Cells[1, col];
                                string headerText = Convert.ToString(cell.Value ?? "");
                                if (headerText.EndsWith("(*変更不可)"))
                                    headerText = headerText.Substring(0, headerText.Length - "(*変更不可)".Length);
                                paramHeaders.Add(headerText);
                                Marshal.ReleaseComObject(cell);
                            }

                            // データ行を走査して変更がある行全体に色を付ける
                            for (int row = 2; row <= rowCount; row++)
                            {
                                try
                                {
                                    dynamic idCell = sheet.Cells[row, 1];
                                    object idValue = idCell.Value;
                                    Marshal.ReleaseComObject(idCell);

                                    if (idValue == null)
                                        continue;

                                    // ElementIdを整数文字列に正規化
                                    string elementIdStr;
                                    if (idValue is double dVal)
                                        elementIdStr = ((int)dVal).ToString();
                                    else
                                        elementIdStr = Convert.ToString(idValue).Trim();

                                    // セル単位で成功/失敗を判定
                                    var successCols = new HashSet<int>();
                                    var failedCols = new HashSet<int>();
                                    for (int i = 0; i < paramHeaders.Count; i++)
                                    {
                                        string key = elementIdStr + "|" + paramHeaders[i];
                                        if (changedSet != null && changedSet.Contains(key))
                                            successCols.Add(i + 3);
                                        else if (failedSet != null && failedSet.Contains(key))
                                            failedCols.Add(i + 3);
                                    }

                                    // 変更・失敗がある行は全列に背景色、セル単位で色分け
                                    if (successCols.Count > 0 || failedCols.Count > 0)
                                    {
                                        for (int col = 1; col <= colCount; col++)
                                        {
                                            dynamic cell = sheet.Cells[row, col];
                                            cell.Interior.Color = excelColor;
                                            Marshal.ReleaseComObject(cell);
                                        }
                                        foreach (int col in successCols)
                                        {
                                            dynamic cell = sheet.Cells[row, col];
                                            cell.Font.Color = blueColor;
                                            cell.Font.Bold = true;
                                            Marshal.ReleaseComObject(cell);
                                        }
                                        foreach (int col in failedCols)
                                        {
                                            dynamic cell = sheet.Cells[row, col];
                                            cell.Font.Color = redColor;
                                            cell.Font.Bold = true;
                                            Marshal.ReleaseComObject(cell);
                                        }
                                        anyMarked = true;
                                    }
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

                    // 各シートの1行目（最終列の次）に凡例を追加
                    if (anyMarked)
                    {
                        for (int s2 = 1; s2 <= sheetCount; s2++)
                        {
                            dynamic sheet2 = targetWb.Sheets[s2];
                            try
                            {
                                dynamic usedRange2 = sheet2.UsedRange;
                                int lastColNum = (int)usedRange2.Column + (int)usedRange2.Columns.Count - 1;
                                Marshal.ReleaseComObject(usedRange2);

                                int legendCol = lastColNum + 1;
                                dynamic legendCell = sheet2.Cells[1, legendCol];
                                string legendText = "(*青字はインポート成功、赤字はインポート失敗)";
                                legendCell.Value = legendText;

                                // "青字" 部分 (3文字目から2文字) を青色太字に
                                dynamic blueChars = legendCell.Characters[3, 2];
                                blueChars.Font.Color = blueColor;
                                blueChars.Font.Bold = true;
                                Marshal.ReleaseComObject(blueChars);

                                // "赤字" 部分 (14文字目から2文字) を赤色太字に
                                dynamic redChars = legendCell.Characters[14, 2];
                                redChars.Font.Color = redColor;
                                redChars.Font.Bold = true;
                                Marshal.ReleaseComObject(redChars);

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
                    // 画面更新を再開（これにより色付けが画面に反映される）
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

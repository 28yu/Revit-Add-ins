using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// 開いているExcelファイルを検出するヘルパー
    /// </summary>
    public static class ExcelProcessHelper
    {
        /// <summary>
        /// 現在Excelで開いているxlsxファイルのパス一覧を取得
        /// </summary>
        public static List<string> GetOpenExcelFiles()
        {
            var result = new List<string>();

            try
            {
                // Excel.Application の ROT (Running Object Table) から取得を試行
                object excelApp = null;
                try
                {
                    excelApp = Marshal.GetActiveObject("Excel.Application");
                }
                catch (COMException)
                {
                    // Excelが起動していない場合
                    return result;
                }

                if (excelApp == null)
                    return result;

                try
                {
                    // excelApp.Workbooks をリフレクションで取得
                    dynamic app = excelApp;
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
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch
            {
                // COM関連のエラーは無視
            }

            return result;
        }
    }
}

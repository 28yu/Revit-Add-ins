using System;
using System.Diagnostics;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.ExcelExportImport.Services;
using Tools28.Commands.ExcelExportImport.Views;
using Tools28.Localization;

namespace Tools28.Commands.ExcelExportImport
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ExcelImportCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // インポートダイアログを表示
                var dialog = new ImportDialog(doc);
                bool? result = dialog.ShowDialog();

                if (result != true || !dialog.ImportExecuted)
                    return Result.Cancelled;

                // ダイアログで生成済みのプレビューデータを再利用（色付け用）
                var previewRows = dialog.PreviewRows;

                // トランザクション内でインポート実行
                ImportResult importResult;
                using (var trans = new Transaction(doc, "Excelインポート"))
                {
                    trans.Start();

                    importResult = ExcelImportService.Import(doc, dialog.SelectedFilePath);

                    if (importResult.SuccessCount > 0)
                    {
                        trans.Commit();
                    }
                    else
                    {
                        trans.RollBack();
                    }
                }

                // インポート成功時、Excelの変更セルに色を付ける
                string markedFilePath = null;
                string colorMethod = null;
                if ((importResult.SuccessCount > 0 || importResult.FailedCells.Count > 0) && previewRows != null)
                {
                    try
                    {
                        markedFilePath = ExcelImportService.MarkImportedCells(
                            dialog.SelectedFilePath, previewRows, out colorMethod, importResult.FailedCells);

                        // 別名保存された場合、色付きファイルを自動で開く
                        if (markedFilePath != null && markedFilePath != dialog.SelectedFilePath)
                        {
                            Process.Start(new ProcessStartInfo(markedFilePath) { UseShellExecute = true });
                        }
                    }
                    catch
                    {
                        // Excel色付けの失敗はインポート結果に影響させない
                    }
                }

                // 結果表示
                var sb = new StringBuilder();
                sb.AppendLine("インポートが完了しました。\n");
                sb.AppendLine($"成功: {importResult.SuccessCount}件");
                sb.AppendLine($"失敗: {importResult.FailCount}件");
                sb.AppendLine($"スキップ: {importResult.SkipCount}件");

                // 色付け結果を表示
                if (markedFilePath != null)
                {
                    if (markedFilePath != dialog.SelectedFilePath)
                    {
                        sb.AppendLine($"\n※ 変更行に色を付けたファイルを別名で保存しました:");
                        sb.AppendLine(markedFilePath);
                    }
                    else if (colorMethod == "COM")
                    {
                        sb.AppendLine("\n※ 開いているExcelの変更行に色を付けました。");
                    }
                    else
                    {
                        sb.AppendLine("\n※ Excelファイルの変更行に色を付けました。");
                    }
                }

                if (importResult.Errors.Count > 0)
                {
                    sb.AppendLine("\n--- エラー詳細 ---");
                    int maxErrors = Math.Min(importResult.Errors.Count, 20);
                    for (int i = 0; i < maxErrors; i++)
                    {
                        sb.AppendLine(importResult.Errors[i]);
                    }
                    if (importResult.Errors.Count > 20)
                    {
                        sb.AppendLine($"... 他 {importResult.Errors.Count - 20}件");
                    }
                }

                if (importResult.Warnings.Count > 0)
                {
                    sb.AppendLine("\n--- 警告 ---");
                    int maxWarnings = Math.Min(importResult.Warnings.Count, 10);
                    for (int i = 0; i < maxWarnings; i++)
                    {
                        sb.AppendLine(importResult.Warnings[i]);
                    }
                    if (importResult.Warnings.Count > 10)
                    {
                        sb.AppendLine($"... 他 {importResult.Warnings.Count - 10}件");
                    }
                }

                TaskDialog.Show(Loc.S("Import.ResultTitle"), sb.ToString());

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message + "\n\nマニュアル: https://28tools.com/addins.html";
                return Result.Failed;
            }
        }
    }
}

using System;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.ExcelExportImport.Services;
using Tools28.Commands.ExcelExportImport.Views;

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

                // 結果表示
                var sb = new StringBuilder();
                sb.AppendLine("インポートが完了しました。\n");
                sb.AppendLine($"成功: {importResult.SuccessCount}件");
                sb.AppendLine($"失敗: {importResult.FailCount}件");
                sb.AppendLine($"スキップ: {importResult.SkipCount}件");

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

                TaskDialog.Show("インポート結果", sb.ToString());

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message + "\n\nマニュアル: https://28yu.github.io/28tools-manual/";
                return Result.Failed;
            }
        }
    }
}

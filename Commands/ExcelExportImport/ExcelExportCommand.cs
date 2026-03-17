using System;
using System.Diagnostics;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using Tools28.Commands.ExcelExportImport.Services;
using Tools28.Commands.ExcelExportImport.Views;

namespace Tools28.Commands.ExcelExportImport
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ExcelExportCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // エクスポートダイアログを表示
                var dialog = new ExportDialog(doc);
                bool? result = dialog.ShowDialog();

                if (result != true)
                    return Result.Cancelled;

                // 保存先を選択
                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excelファイル (*.xlsx)|*.xlsx",
                    DefaultExt = ".xlsx",
                    FileName = $"{doc.Title}_パラメータ"
                };

                if (saveDialog.ShowDialog() != true)
                    return Result.Cancelled;

                // エクスポート実行
                ExcelExportService.Export(
                    doc,
                    saveDialog.FileName,
                    dialog.SelectedCategories,
                    dialog.OutputParameters);

                // エクスポートしたExcelファイルを自動で開く
                Process.Start(new ProcessStartInfo(saveDialog.FileName) { UseShellExecute = true });

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

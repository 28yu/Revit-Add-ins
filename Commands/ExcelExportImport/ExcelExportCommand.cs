using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using Tools28.Commands.ExcelExportImport.Models;
using Tools28.Commands.ExcelExportImport.Services;
using Tools28.Commands.ExcelExportImport.Views;
using Tools28.Localization;

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
                // 現在のビュー・選択状態を取得
                View activeView = doc.ActiveView;
                bool hasActiveView = activeView != null && !(activeView is ViewSchedule);

                ICollection<ElementId> selectionIds = uidoc.Selection.GetElementIds();
                bool hasSelection = selectionIds != null && selectionIds.Count > 0;

                // 範囲選択ダイアログ
                var scopeDialog = new ScopeSelectionDialog(hasActiveView, hasSelection);
                if (scopeDialog.ShowDialog() != true)
                    return Result.Cancelled;

                ExportScope scope = scopeDialog.SelectedScope;

                // エクスポートダイアログを表示（スコープを渡す）
                var dialog = new ExportDialog(doc, scope, activeView, selectionIds);
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

                // エクスポート実行（スコープを渡す）
                ExcelExportService.Export(
                    doc,
                    saveDialog.FileName,
                    dialog.SelectedCategories,
                    dialog.OutputParameters,
                    dialog.SplitByCategory,
                    scope,
                    activeView,
                    selectionIds);

                // エクスポートしたExcelファイルを自動で開く
                Process.Start(new ProcessStartInfo(saveDialog.FileName) { UseShellExecute = true });

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

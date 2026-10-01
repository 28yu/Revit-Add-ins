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
                scopeDialog.SetRevitOwner(commandData);
                if (scopeDialog.ShowDialog() != true)
                    return Result.Cancelled;

                ExportScope scope = scopeDialog.SelectedScope;

                // エクスポートダイアログを表示（スコープを渡す）
                var dialog = new ExportDialog(doc, scope, activeView, selectionIds);
                dialog.SetRevitOwner(commandData);
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

                // エクスポート実行（スコープを渡す）。
                // 大容量モデルでは時間がかかるため、進み具合の画面を出してキャンセルできるようにする。
                var timings = new ParameterTimingTracker();
                var progressWindow = new ExportProgressWindow();
                progressWindow.SetRevitOwner(commandData);
                var totalWatch = Stopwatch.StartNew();
                try
                {
                    progressWindow.Show();
                    // 処理中は Revit 本体を操作できないようにする（終了時に自動で元に戻る）
                    using (progressWindow.BlockRevitInput())
                    {
                        ExcelExportService.Export(
                            doc,
                            saveDialog.FileName,
                            dialog.SelectedCategories,
                            dialog.OutputParameters,
                            dialog.SplitByCategory,
                            scope,
                            activeView,
                            selectionIds,
                            dialog.IncludeParamGroup,
                            progressWindow,
                            timings);
                    }
                }
                catch (OperationCanceledException)
                {
                    progressWindow.Finish();
                    LogSlowParameters(timings, totalWatch, "キャンセル");
                    TaskDialog.Show(Loc.S("Export.Title"), BuildCancelledMessage(timings));
                    return Result.Cancelled;
                }
                finally
                {
                    progressWindow.Finish();
                }

                LogSlowParameters(timings, totalWatch, "完了");

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

        /// <summary>キャンセル時のメッセージ。時間のかかっていたパラメータがあれば一緒に示す。</summary>
        private static string BuildCancelledMessage(ParameterTimingTracker timings)
        {
            string msg = Loc.S("Export.Cancelled");

            // 1要素あたり 5ms 以上かかっていたものだけを候補として挙げる
            var slow = timings.Top(5).Where(e => e.AverageMs >= 5.0).ToList();
            if (slow.Count == 0)
                return msg;

            var lines = slow.Select(e => string.Format(Loc.S("Export.Cancelled.SlowParamLine"),
                e.Label, e.AverageMs.ToString("0.0"), e.TotalSeconds.ToString("0")));
            return msg + "\n\n" + Loc.S("Export.Cancelled.SlowParams") + "\n" + string.Join("\n", lines);
        }

        /// <summary>原因調査用に、読み取りに時間のかかったパラメータ上位をログに残す。</summary>
        private static void LogSlowParameters(ParameterTimingTracker timings, Stopwatch totalWatch, string outcome)
        {
            DiagLog.Write($"[ExcelExport] {outcome}: 全体 {totalWatch.ElapsedMilliseconds} ms");
            foreach (var e in timings.Top(10))
                DiagLog.Write($"[ExcelExport]   {e.Label}: 合計 {e.TotalSeconds:0.0} 秒 / {e.Count} 件, 平均 {e.AverageMs:0.00} ms");
        }
    }
}

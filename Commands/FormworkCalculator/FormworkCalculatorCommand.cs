using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.FormworkCalculator.Engine;
using Tools28.Commands.FormworkCalculator.Models;
using Tools28.Commands.FormworkCalculator.Output;
using Tools28.Commands.FormworkCalculator.Views;
using Tools28.Localization;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;
using WinFormsDialogResult = System.Windows.Forms.DialogResult;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace Tools28.Commands.FormworkCalculator
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FormworkCalculatorCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            try
            {
                var dialog = new FormworkDialog();
                if (dialog.ShowDialog() != true)
                    return Result.Cancelled;

                var settings = dialog.Settings;

                if (settings.ExportToExcel)
                {
                    using (var sfd = new SaveFileDialog
                    {
                        Title = Loc.S("Formwork.SaveExcelTitle"),
                        Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                        FileName = $"FormworkQuantity_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                    })
                    {
                        if (sfd.ShowDialog() != WinFormsDialogResult.OK)
                            return Result.Cancelled;
                        settings.ExcelOutputPath = sfd.FileName;
                    }
                }

                var calc = new FormworkCalcEngine(doc, settings, activeView);
                FormworkResult result;
                try
                {
                    result = calc.Run();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show(Loc.S("Common.Error"),
                        string.Format(Loc.S("Formwork.CalcFailed"), ex.Message));
                    return Result.Failed;
                }

                if (result == null || result.ProcessedElementCount == 0)
                {
                    TaskDialog.Show(Loc.S("Common.Warning"), Loc.S("Formwork.NoElements"));
                    return Result.Cancelled;
                }

                if (settings.ExportToExcel && !string.IsNullOrEmpty(settings.ExcelOutputPath))
                {
                    try
                    {
                        ExcelExporter.Export(settings.ExcelOutputPath, settings, result);
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show(Loc.S("Common.Error"),
                            string.Format(Loc.S("Formwork.ExcelFailed"), ex.Message));
                    }
                }

                // Revit 側の出力（共有パラメータ → DirectShape → 集計表 → 他ビュー非表示）
                ElementId scheduleViewId = null;
                ElementId view3DId = null;
                List<ElementId> createdShapeIds = null;

                if (settings.CreateSchedule || settings.Create3DView)
                {
                    // Step 1: 共有パラメータを用意
                    using (var tParams = new Transaction(doc, "型枠数量算出 - 共有パラメータ"))
                    {
                        tParams.Start();
                        try
                        {
                            FormworkParameterManager.EnsureParameters(doc, uiapp.Application);
                            tParams.Commit();
                        }
                        catch
                        {
                            tParams.RollBack();
                        }
                    }

                    // Step 2: DirectShape 作成 + 集計表作成
                    using (var t = new Transaction(doc, "型枠数量算出 - ビュー作成"))
                    {
                        t.Start();
                        try
                        {
                            if (settings.Create3DView)
                            {
                                var v3d = FormworkVisualizer.CreateVisualization(doc, result, settings);
                                if (v3d?.AnalysisView != null)
                                {
                                    view3DId = v3d.AnalysisView.Id;
                                    createdShapeIds = v3d.CreatedShapeIds;
                                }
                            }

                            if (settings.CreateSchedule)
                            {
                                scheduleViewId = ScheduleCreator.CreateSchedule(doc);
                            }

                            t.Commit();
                        }
                        catch (Exception ex)
                        {
                            t.RollBack();
                            TaskDialog.Show(Loc.S("Common.Error"),
                                string.Format(Loc.S("Formwork.ViewFailed"), ex.Message));
                        }
                    }

                    // Step 3: 他ビュー非表示（作成した DirectShape を既存ビューで隠す）
                    if (createdShapeIds != null && createdShapeIds.Count > 0)
                    {
                        using (var tHide = new Transaction(doc, "型枠数量算出 - 他ビュー非表示"))
                        {
                            tHide.Start();
                            try
                            {
                                FormworkVisualizer.HideInOtherViews(doc, createdShapeIds, view3DId);
                                tHide.Commit();
                            }
                            catch
                            {
                                tHide.RollBack();
                            }
                        }
                    }
                }

                string summary =
                    string.Format(Loc.S("Formwork.DoneMsg"),
                        result.ProcessedElementCount,
                        result.TotalFormworkArea,
                        result.TotalDeductedArea,
                        result.InclinedFaceArea);

                if (settings.ExportToExcel && !string.IsNullOrEmpty(settings.ExcelOutputPath))
                    summary += "\n\n" + string.Format(Loc.S("Formwork.ExcelAt"), settings.ExcelOutputPath);

                if (createdShapeIds != null && createdShapeIds.Count > 0)
                    summary += "\n\n" + string.Format(Loc.S("Formwork.ShapesCreated"), createdShapeIds.Count);

                if (scheduleViewId != null)
                    summary += "\n" + Loc.S("Formwork.ScheduleCreated");

                if (result.Errors.Count > 0)
                    summary += "\n\n" + string.Format(Loc.S("Formwork.ErrorCount"), result.Errors.Count);

                TaskDialog.Show(Loc.S("Formwork.DoneTitle"), summary);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = string.Format(Loc.S("Formwork.Fatal"), ex.Message)
                    + "\n\n" + ex.StackTrace;
                return Result.Failed;
            }
        }
    }
}

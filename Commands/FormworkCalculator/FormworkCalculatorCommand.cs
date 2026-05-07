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
                var dialog = new FormworkDialog(doc);
                if (dialog.ShowDialog() != true)
                    return Result.Cancelled;

                var settings = dialog.Settings;

                if (settings.ExportToExcel)
                {
                    using (var sfd = new SaveFileDialog
                    {
                        Title = Loc.S("Formwork.SaveExcelTitle"),
                        Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                        FileName = "型枠数量集計.xlsx",
                    })
                    {
                        if (sfd.ShowDialog() != WinFormsDialogResult.OK)
                            return Result.Cancelled;
                        settings.ExcelOutputPath = sfd.FileName;
                    }
                }

                FormworkResult result;
                try
                {
                    var calc = new FormworkCalcEngine(doc, settings, activeView);
                    result = calc.Run();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show(Loc.S("Common.Error"),
                        string.Format(Loc.S("Formwork.CalcFailed"), ex.Message));
                    return Result.Failed;
                }

                if (result == null ||
                    (result.ProcessedElementCount == 0 && result.ExcludedResults.Count == 0))
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

                ElementId scheduleViewId = null;
                ElementId summaryScheduleId = null;
                ElementId view3DId = null;
                List<ElementId> createdShapeIds = null;

                if (settings.CreateSchedule || settings.Create3DView)
                {
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

                    using (var t = new Transaction(doc, "型枠数量算出 - ビュー作成"))
                    {
                        t.Start();
                        try
                        {
                            if (settings.Create3DView)
                            {
                                var v3d = FormworkVisualizer.CreateVisualization(doc, result, settings, activeView);
                                if (v3d?.AnalysisView != null)
                                {
                                    view3DId = v3d.AnalysisView.Id;
                                    createdShapeIds = v3d.CreatedShapeIds;
                                }
                            }

                            if (settings.CreateSchedule)
                            {
                                scheduleViewId = ScheduleCreator.CreateSchedule(doc, result);
                                // 動的合計サマリ集計表 (IsItemized=false で 1 行で全件集約)
                                summaryScheduleId = ScheduleCreator.CreateSummarySchedule(doc);
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

                // ビュータブを順に開く: 3D ビュー (タブを開く) → サマリ集計表 (最終アクティブ)。
                // メイン集計表は Project Browser から手動で開く。
                if (view3DId != null && view3DId != ElementId.InvalidElementId)
                {
                    try
                    {
                        var v = doc.GetElement(view3DId) as View;
                        if (v != null) uidoc.ActiveView = v;
                    }
                    catch { }
                }
                ElementId finalViewId =
                    (summaryScheduleId != null && summaryScheduleId != ElementId.InvalidElementId)
                        ? summaryScheduleId
                        : ((scheduleViewId != null && scheduleViewId != ElementId.InvalidElementId)
                            ? scheduleViewId
                            : view3DId);
                if (finalViewId != null && finalViewId != ElementId.InvalidElementId)
                {
                    try
                    {
                        var v = doc.GetElement(finalViewId) as View;
                        if (v != null) uidoc.ActiveView = v;
                    }
                    catch { }
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
                if (summaryScheduleId != null)
                    summary += "\n" + Loc.S("Formwork.SummaryScheduleCreated");

                if (result.ExcludedResults != null && result.ExcludedResults.Count > 0)
                {
                    int steelN = 0, deckN = 0;
                    foreach (var ex in result.ExcludedResults)
                    {
                        if (ex.Kind == ExclusionKind.Steel) steelN++;
                        else if (ex.Kind == ExclusionKind.DeckSlab) deckN++;
                    }
                    if (steelN > 0)
                        summary += "\n\n" + string.Format(Loc.S("Formwork.SteelExcluded"), steelN);
                    if (deckN > 0)
                        summary += "\n" + string.Format(Loc.S("Formwork.DeckSlabExcluded"), deckN);
                    summary += "\n" + Loc.S("Formwork.ExcludedFilterNote");
                }

                if (result.Errors.Count > 0)
                    summary += "\n\n" + string.Format(Loc.S("Formwork.ErrorCount"), result.Errors.Count);

                TaskDialog.Show(Loc.S("Formwork.DoneTitle"), summary);
                FormworkDebugLog.Close();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Close();
                message = string.Format(Loc.S("Formwork.Fatal"), ex.Message)
                    + "\n\n" + ex.StackTrace;
                return Result.Failed;
            }
        }
    }
}

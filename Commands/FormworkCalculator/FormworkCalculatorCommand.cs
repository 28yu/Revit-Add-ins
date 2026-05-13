using System;
using System.Collections.Generic;
using System.Linq;
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
                dialog.SetRevitOwner(commandData);
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
                ElementId sheetId = null;
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
                                // 結果に含まれる全ソース (ホスト・各リンク) を列挙し
                                // それぞれに対して個別の集計表を作成する。
                                // - リンクなし: 「型枠数量集計」(従来通り)
                                // - リンクあり: 「型枠数量集計 - ホスト」「型枠数量集計 - 構造_A棟」等
                                var sources = result.ElementResults
                                    .Select(er => string.IsNullOrEmpty(er.SourceName)
                                        ? ElementSourceRegistry.HostSourceName
                                        : er.SourceName)
                                    .Distinct()
                                    .OrderBy(s => s != ElementSourceRegistry.HostSourceName)
                                    .ThenBy(s => s)
                                    .ToList();

                                if (sources.Count <= 1)
                                {
                                    // ホストのみ (リンク無し or リンク要素なし)
                                    scheduleViewId = ScheduleCreator.CreateSchedule(doc, result);
                                }
                                else
                                {
                                    // ホスト + リンク: ソース毎に個別の集計表を作る
                                    // メインの scheduleViewId はホストの集計表を指す
                                    foreach (var srcName in sources)
                                    {
                                        var id = ScheduleCreator.CreateSchedule(doc, result, srcName);
                                        if (srcName == ElementSourceRegistry.HostSourceName)
                                            scheduleViewId = id;
                                    }
                                }
                                // 動的合計サマリ集計表 (ホスト・リンク別の小計 + 全体合計)
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

                    // 3Dビューと集計表が両方できていればシートに自動配置 (別トランザクション:
                    // ScheduleSheetInstance.Create は元の集計表がコミット済みである必要があるため)
                    bool haveAnyOutput =
                        (view3DId != null && view3DId != ElementId.InvalidElementId) ||
                        (scheduleViewId != null && scheduleViewId != ElementId.InvalidElementId) ||
                        (summaryScheduleId != null && summaryScheduleId != ElementId.InvalidElementId);
                    if (settings.CreateSheet && haveAnyOutput)
                    {
                        using (var tSheet = new Transaction(doc, "型枠数量算出 - シート作成"))
                        {
                            tSheet.Start();
                            try
                            {
                                sheetId = FormworkSheetCreator.CreateSheet(
                                    doc, view3DId, scheduleViewId, summaryScheduleId);
                                tSheet.Commit();
                            }
                            catch (Exception ex)
                            {
                                tSheet.RollBack();
                                FormworkDebugLog.Log($"  [Sheet] CreateSheet EX: {ex.Message}");
                            }
                        }
                    }
                }

                // ビュータブを順に開く: 集計表系 → 最後に 3D ビュー (最終アクティブ)。
                // 3D ビューがユーザーにとってのメインビューなので、実行後はそこに遷移する。
                // メイン集計表は Project Browser から手動で開く。
                if (scheduleViewId != null && scheduleViewId != ElementId.InvalidElementId)
                {
                    try
                    {
                        var v = doc.GetElement(scheduleViewId) as View;
                        if (v != null) uidoc.ActiveView = v;
                    }
                    catch { }
                }
                if (summaryScheduleId != null && summaryScheduleId != ElementId.InvalidElementId)
                {
                    try
                    {
                        var v = doc.GetElement(summaryScheduleId) as View;
                        if (v != null) uidoc.ActiveView = v;
                    }
                    catch { }
                }
                if (view3DId != null && view3DId != ElementId.InvalidElementId)
                {
                    try
                    {
                        var v = doc.GetElement(view3DId) as View;
                        if (v != null) uidoc.ActiveView = v;
                    }
                    catch { }
                }
                // シートを最終アクティブにする (3Dビュー・集計表の上に配置した全体像が見える)
                if (sheetId != null && sheetId != ElementId.InvalidElementId)
                {
                    try
                    {
                        var v = doc.GetElement(sheetId) as View;
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
                if (sheetId != null && sheetId != ElementId.InvalidElementId)
                {
                    try
                    {
                        var sh = doc.GetElement(sheetId) as ViewSheet;
                        if (sh != null)
                            summary += "\n" + string.Format(
                                Loc.S("Formwork.SheetCreated"),
                                sh.SheetNumber, sh.Name);
                    }
                    catch { }
                }

                // リンクモデル取り込み件数を表示
                if (settings.IncludeLinkedModels)
                {
                    int linkedElemCount = result.ElementResults.Count(er =>
                        !string.IsNullOrEmpty(er.SourceName)
                        && er.SourceName != ElementSourceRegistry.HostSourceName);
                    if (linkedElemCount > 0 || result.LinkedInstanceCount > 0)
                    {
                        summary += "\n\n" + string.Format(
                            Loc.S("Formwork.LinkedIncluded"),
                            linkedElemCount, result.LinkedInstanceCount);
                    }
                }

                if (result.ExcludedResults != null && result.ExcludedResults.Count > 0)
                {
                    int steelN = 0, deckN = 0, sweepN = 0, steelStairN = 0, lgsN = 0;
                    foreach (var ex in result.ExcludedResults)
                    {
                        if (ex.Kind == ExclusionKind.Steel) steelN++;
                        else if (ex.Kind == ExclusionKind.DeckSlab) deckN++;
                        else if (ex.Kind == ExclusionKind.WallSweep) sweepN++;
                        else if (ex.Kind == ExclusionKind.SteelStair) steelStairN++;
                        else if (ex.Kind == ExclusionKind.LgsWall) lgsN++;
                    }
                    if (steelN > 0)
                        summary += "\n\n" + string.Format(Loc.S("Formwork.SteelExcluded"), steelN);
                    if (deckN > 0)
                        summary += "\n" + string.Format(Loc.S("Formwork.DeckSlabExcluded"), deckN);
                    if (sweepN > 0)
                        summary += "\n" + string.Format(Loc.S("Formwork.WallSweepExcluded"), sweepN);
                    if (steelStairN > 0)
                        summary += "\n" + string.Format(Loc.S("Formwork.SteelStairExcluded"), steelStairN);
                    if (lgsN > 0)
                        summary += "\n" + string.Format(Loc.S("Formwork.LgsExcluded"), lgsN);
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

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
                        FileName = "型枠数量集計.xlsx",
                    })
                    {
                        if (sfd.ShowDialog() != WinFormsDialogResult.OK)
                            return Result.Cancelled;
                        settings.ExcelOutputPath = sfd.FileName;
                    }
                }

                // ReferenceIntersector に必要な View3D を確保（既存利用 or 一時作成）
                View3D rayView = null;
                bool tempViewCreated = false;
                using (var tRayView = new Transaction(doc, "型枠数量算出 - レイビュー確保"))
                {
                    tRayView.Start();
                    try
                    {
                        rayView = FindOrCreateRayView(doc, out tempViewCreated);
                        tRayView.Commit();
                    }
                    catch
                    {
                        tRayView.RollBack();
                    }
                }

                if (rayView == null)
                {
                    TaskDialog.Show(Loc.S("Common.Error"),
                        "ReferenceIntersector 用の 3D ビューを確保できませんでした。");
                    return Result.Failed;
                }

                FormworkResult result;
                try
                {
                    var calc = new FormworkCalcEngine(doc, settings, activeView, rayView);
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
                    DeleteTempRayView(doc, rayView, tempViewCreated);
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

                // 一時 rayView を削除（最終 3D ビューを残す）
                DeleteTempRayView(doc, rayView, tempViewCreated);

                // 自動作成した 3D ビューをアクティブ化
                if (view3DId != null && view3DId != ElementId.InvalidElementId)
                {
                    try
                    {
                        var v = doc.GetElement(view3DId) as View;
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

        /// <summary>
        /// ReferenceIntersector 用の 3D ビューを常に新規作成する。
        /// 既存ビューを流用すると、ビューテンプレート・セクションボックス・
        /// カテゴリ非表示などでレイキャストが正しく機能しないため、
        /// 必ず解析専用の一時ビューを使う。
        /// </summary>
        private static View3D FindOrCreateRayView(Document doc, out bool created)
        {
            created = false;
            var vftype = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(v => v.ViewFamily == ViewFamily.ThreeDimensional);
            if (vftype == null) return null;

            var v3d = View3D.CreateIsometric(doc, vftype.Id);
            try { v3d.Name = "Tools28_型枠_RayView_" + DateTime.Now.Ticks; } catch { }

            // ビューテンプレート・セクションボックスを解除し、全カテゴリを可視化
            try { v3d.ViewTemplateId = ElementId.InvalidElementId; } catch { }
            try { v3d.IsSectionBoxActive = false; } catch { }
            try { v3d.DetailLevel = ViewDetailLevel.Fine; } catch { }
            try { v3d.DisplayStyle = DisplayStyle.Shading; } catch { }

            created = true;
            return v3d;
        }

        private static void DeleteTempRayView(Document doc, View3D view, bool wasCreated)
        {
            if (!wasCreated || view == null) return;
            using (var t = new Transaction(doc, "型枠数量算出 - 一時ビュー削除"))
            {
                t.Start();
                try
                {
                    doc.Delete(view.Id);
                    t.Commit();
                }
                catch
                {
                    t.RollBack();
                }
            }
        }
    }
}

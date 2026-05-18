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
                // ダイアログ表示前にプロジェクトブラウザの 3D ビュー選択を検出。
                // 検出した場合、ダイアログの「選択のビュー」を有効化・初期選択にする。
                var selectedView3Ds = CollectSelectedView3Ds(uidoc) ?? new List<View3D>();

                var dialog = new FormworkDialog(doc, defaults: null,
                    selectedViewsCount: selectedView3Ds.Count);
                dialog.SetRevitOwner(commandData);
                if (dialog.ShowDialog() != true)
                    return Result.Cancelled;

                var settings = dialog.Settings;

                // 計算範囲モードに従って対象ソースビュー群を確定する:
                //   - SelectedViews: プロジェクトブラウザ選択の 3D ビュー群
                //   - CurrentView / EntireProject: アクティブビュー (3D の場合のみ)
                List<View3D> sourceViews;
                if (settings.Scope == CalculationScope.SelectedViews)
                {
                    sourceViews = selectedView3Ds;
                }
                else
                {
                    sourceViews = new List<View3D>();
                    if (activeView is View3D v3dActive)
                        sourceViews.Add(v3dActive);
                }

                // [2] 分析ビューが選択/アクティブな場合、元のソースビューを解決する
                sourceViews = ResolveAnalysisViews(doc, sourceViews);

                if (sourceViews.Count == 0)
                {
                    TaskDialog.Show(Loc.S("Common.Warning"),
                        "型枠数量算出は 3D ビューで実行してください。\nプロジェクトブラウザで複数の 3D ビューを選択すると、それらをまとめて 1 シートに集約できます。");
                    return Result.Cancelled;
                }

                // [2] 更新モード判定: 対象ソースビューに対応する既存の分析3Dビューまたは
                // ビュー別集計表が見つかった場合、ユーザーに「更新 / 再作成 / キャンセル」を確認する。
                // 更新モード時: 既存の分析ビュー・集計表を再利用し、シートと合計集計表は手を付けない。
                // 再作成モード時: 従来通り全て削除・再作成する。
                bool hasExistingOutputs = false;
                foreach (var sv in sourceViews)
                {
                    if (FormworkOutputFinder.FindAnalysisView(doc, sv.Name) != null
                        || FormworkOutputFinder.FindMainSchedule(doc, sv.Name) != null)
                    {
                        hasExistingOutputs = true;
                        break;
                    }
                }

                bool updateMode = false;
                if (hasExistingOutputs)
                {
                    var td = new TaskDialog(Loc.S("Formwork.UpdateConfirm.Title"))
                    {
                        MainInstruction = Loc.S("Formwork.UpdateConfirm.MainInstruction"),
                        MainContent = Loc.S("Formwork.UpdateConfirm.Content"),
                        CommonButtons = TaskDialogCommonButtons.Cancel,
                    };
                    td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                        Loc.S("Formwork.UpdateConfirm.UpdateLabel"),
                        Loc.S("Formwork.UpdateConfirm.UpdateDesc"));
                    td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                        Loc.S("Formwork.UpdateConfirm.RecreateLabel"),
                        Loc.S("Formwork.UpdateConfirm.RecreateDesc"));
                    var tdResult = td.Show();
                    if (tdResult == TaskDialogResult.CommandLink1) updateMode = true;
                    else if (tdResult == TaskDialogResult.CommandLink2) updateMode = false;
                    else return Result.Cancelled;
                }

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

                // 各ソースビューの計算結果と出力IDを保持
                var perViewResults = new List<FormworkResult>();
                var perViewSourceNames = new List<string>();
                var perViewAnalysisViewIds = new List<ElementId>();
                var perViewScheduleIds = new List<ElementId>();
                // 分析ビュー Id → そのビューで作成された DirectShape Id 群
                var perAnalysisViewShapeIds = new Dictionary<ElementId, List<ElementId>>();

                // エンジン用に正規化したスコープを使う:
                // SelectedViews モードは、ループ内で各ビューを CurrentView として処理する。
                bool isMultiViewMode = settings.Scope == CalculationScope.SelectedViews;
                FormworkResult firstResult = null;
                foreach (var sv in sourceViews)
                {
                    var engineSettings = settings;
                    if (isMultiViewMode)
                    {
                        // SelectedViews は CurrentView と同じ視点フィルタリングを使う
                        engineSettings = CloneWithScope(settings, CalculationScope.CurrentView);
                    }

                    FormworkResult r;
                    try
                    {
                        var calc = new FormworkCalcEngine(doc, engineSettings, sv);
                        r = calc.Run();
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show(Loc.S("Common.Error"),
                            string.Format(Loc.S("Formwork.CalcFailed"), $"[{sv.Name}] {ex.Message}"));
                        return Result.Failed;
                    }
                    if (r == null) continue;
                    perViewResults.Add(r);
                    perViewSourceNames.Add(sv.Name);
                    if (firstResult == null) firstResult = r;
                }

                if (firstResult == null ||
                    perViewResults.All(r => r.ProcessedElementCount == 0 && r.ExcludedResults.Count == 0))
                {
                    TaskDialog.Show(Loc.S("Common.Warning"), Loc.S("Formwork.NoElements"));
                    return Result.Cancelled;
                }

                if (settings.ExportToExcel && !string.IsNullOrEmpty(settings.ExcelOutputPath))
                {
                    try
                    {
                        // Excel 出力は最初のビューの結果を使う (複数ビューの場合は最初の選択のみ)
                        ExcelExporter.Export(settings.ExcelOutputPath, settings, firstResult);
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show(Loc.S("Common.Error"),
                            string.Format(Loc.S("Formwork.ExcelFailed"), ex.Message));
                    }
                }

                ElementId summaryScheduleId = null;
                ElementId sheetId = null;
                int totalShapesCreated = 0;
                // 集計表IDをシート配置用に保持
                var allMainScheduleIds = new List<ElementId>();

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
                            // 選択された各 3D ビューについて個別に処理する。
                            // ParamSourceView でタグ付けされるため、各ビューの DirectShape は分離管理可能。
                            for (int i = 0; i < perViewResults.Count; i++)
                            {
                                var r = perViewResults[i];
                                var sv = sourceViews[i];
                                var svName = perViewSourceNames[i];

                                if (settings.Create3DView)
                                {
                                    var v3d = FormworkVisualizer.CreateVisualization(doc, r, settings, sv, updateMode);
                                    if (v3d?.AnalysisView != null)
                                    {
                                        perViewAnalysisViewIds.Add(v3d.AnalysisView.Id);
                                        if (v3d.CreatedShapeIds != null)
                                        {
                                            totalShapesCreated += v3d.CreatedShapeIds.Count;
                                            perAnalysisViewShapeIds[v3d.AnalysisView.Id] =
                                                new List<ElementId>(v3d.CreatedShapeIds);
                                        }
                                    }
                                }

                                if (settings.CreateSchedule)
                                {
                                    // 1ビュー1集計表: ParamSourceView でフィルタ
                                    // 名前: 「型枠数量集計 - {ビュー名}」
                                    ElementId sid = null;
                                    try
                                    {
                                        sid = ScheduleCreator.CreateSchedule(doc, r,
                                            sourceFilter: null, sourceViewFilter: svName,
                                            updateMode: updateMode);
                                    }
                                    catch (Exception schEx)
                                    {
                                        FormworkDebugLog.Log(
                                            $"  [Sched] FAILED schedule for view '{svName}': {schEx.Message}");
                                    }
                                    if (sid != null && sid != ElementId.InvalidElementId)
                                    {
                                        perViewScheduleIds.Add(sid);
                                        allMainScheduleIds.Add(sid);
                                    }
                                }
                            }

                            if (settings.CreateSchedule && !updateMode)
                            {
                                // 動的合計サマリ集計表 (全ビュー横断の合計)
                                // 更新モードでは既存の合計表をそのまま使う ([2])
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

                    // 全ソースビューの DirectShape を、それぞれの分析ビュー以外で非表示にする。
                    // 過去実行で作成された他ビューの DirectShape (ParamSourceView でタグ済) も含めて
                    // 再ラベリングする。これにより、本実行で新規作成された分析ビューに
                    // 過去の他ビュー DirectShape が紛れ込む問題を防ぐ。
                    using (var tHide = new Transaction(doc, "型枠数量算出 - 他ビュー非表示"))
                    {
                        tHide.Start();
                        try
                        {
                            FormworkVisualizer.HideAllFormworkShapesInOtherViews(doc);
                            tHide.Commit();
                        }
                        catch (Exception ex)
                        {
                            tHide.RollBack();
                            FormworkDebugLog.Log($"  [Hide] HideAllFormworkShapesInOtherViews EX: {ex.Message}");
                        }
                    }

                    // シートにはプロジェクト内に存在する全型枠分析ビュー + 全型枠集計表を配置する。
                    // これにより [2] 「特定ビューのみ再実行」時も、過去に作成した他ビューの
                    // 解析ビュー・集計表をそのまま保持しつつシートも再構成できる。
                    bool haveAnyOutput =
                        perViewAnalysisViewIds.Count > 0 ||
                        allMainScheduleIds.Count > 0 ||
                        (summaryScheduleId != null && summaryScheduleId != ElementId.InvalidElementId);
                    // 更新モードでは既存シートをそのまま保持する ([2])
                    if (settings.CreateSheet && haveAnyOutput && !updateMode)
                    {
                        using (var tSheet = new Transaction(doc, "型枠数量算出 - シート作成"))
                        {
                            tSheet.Start();
                            try
                            {
                                // プロジェクト内の全型枠分析ビュー・集計表を収集
                                var allAnalysisViewIds = FormworkSheetCreator.CollectAllAnalysisViewIds(doc);
                                var allScheduleIds = FormworkSheetCreator.CollectAllPerViewScheduleIds(doc);
                                sheetId = FormworkSheetCreator.CreateOrUpdateSheet(
                                    doc, allAnalysisViewIds, allScheduleIds, summaryScheduleId);
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
                ElementId firstScheduleId = perViewScheduleIds.FirstOrDefault();
                ElementId firstAnalysisViewId = perViewAnalysisViewIds.FirstOrDefault();
                if (firstScheduleId != null && firstScheduleId != ElementId.InvalidElementId)
                {
                    try
                    {
                        var v = doc.GetElement(firstScheduleId) as View;
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
                if (firstAnalysisViewId != null && firstAnalysisViewId != ElementId.InvalidElementId)
                {
                    try
                    {
                        var v = doc.GetElement(firstAnalysisViewId) as View;
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

                // 全ビューの合算で集計を表示
                int totalProcessed = perViewResults.Sum(r => r.ProcessedElementCount);
                double totalFormworkArea = perViewResults.Sum(r => r.TotalFormworkArea);
                double totalDeductedArea = perViewResults.Sum(r => r.TotalDeductedArea);
                double totalInclinedArea = perViewResults.Sum(r => r.InclinedFaceArea);

                string summary =
                    string.Format(Loc.S("Formwork.DoneMsg"),
                        totalProcessed,
                        totalFormworkArea,
                        totalDeductedArea,
                        totalInclinedArea);

                if (sourceViews.Count > 1)
                    summary = $"対象3Dビュー: {sourceViews.Count}件\n[{string.Join(", ", perViewSourceNames)}]\n\n" + summary;

                if (settings.ExportToExcel && !string.IsNullOrEmpty(settings.ExcelOutputPath))
                    summary += "\n\n" + string.Format(Loc.S("Formwork.ExcelAt"), settings.ExcelOutputPath);

                if (totalShapesCreated > 0)
                    summary += "\n\n" + string.Format(Loc.S("Formwork.ShapesCreated"), totalShapesCreated);

                if (perViewScheduleIds.Count > 0)
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
                    int linkedElemCount = perViewResults.Sum(r => r.ElementResults.Count(er =>
                        !string.IsNullOrEmpty(er.SourceName)
                        && er.SourceName != ElementSourceRegistry.HostSourceName));
                    int linkedInstanceCount = perViewResults.Sum(r => r.LinkedInstanceCount);
                    if (linkedElemCount > 0 || linkedInstanceCount > 0)
                    {
                        summary += "\n\n" + string.Format(
                            Loc.S("Formwork.LinkedIncluded"),
                            linkedElemCount, linkedInstanceCount);
                    }
                }

                var allExcluded = perViewResults
                    .Where(r => r.ExcludedResults != null)
                    .SelectMany(r => r.ExcludedResults)
                    .ToList();
                if (allExcluded.Count > 0)
                {
                    int steelN = 0, deckN = 0, sweepN = 0, steelStairN = 0, lgsN = 0;
                    foreach (var ex in allExcluded)
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

                int totalErrors = perViewResults.Sum(r => r.Errors?.Count ?? 0);
                if (totalErrors > 0)
                    summary += "\n\n" + string.Format(Loc.S("Formwork.ErrorCount"), totalErrors);

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

        /// <summary>
        /// 分析ビュー ("3D_型枠数量 - XXX" 等) が含まれている場合、
        /// タグから元のソースビュー名を取得して実際のソースビューに置き換える。
        /// これにより、分析ビューをアクティブにした状態でボタンを実行しても
        /// 更新モード検出が正常に機能する。
        /// </summary>
        private static List<View3D> ResolveAnalysisViews(Document doc, List<View3D> views)
        {
            if (views == null || views.Count == 0) return views ?? new List<View3D>();

            var allView3Ds = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate)
                .ToDictionary(v => v.Name, v => v);

            var resolved = new List<View3D>();
            foreach (var v in views)
            {
                if (FormworkVisualizer.IsAnalysisViewName(v.Name))
                {
                    var sourceViewName = FormworkParameterManager.GetRelatedSourceView(v);
                    if (!string.IsNullOrEmpty(sourceViewName)
                        && allView3Ds.TryGetValue(sourceViewName, out var sourceView))
                    {
                        resolved.Add(sourceView);
                    }
                    else
                    {
                        // タグなし / ソースビュー不明の場合はそのまま追加
                        resolved.Add(v);
                    }
                }
                else
                {
                    resolved.Add(v);
                }
            }
            return resolved;
        }

        /// <summary>
        /// Scope のみ変更したコピーを返す。元の settings は変更しない。
        /// </summary>
        private static FormworkSettings CloneWithScope(FormworkSettings src, CalculationScope newScope)
        {
            return new FormworkSettings
            {
                Scope = newScope,
                GroupByCategory = src.GroupByCategory,
                GroupByZone = src.GroupByZone,
                ZoneParameterName = src.ZoneParameterName,
                GroupByFormworkType = src.GroupByFormworkType,
                FormworkTypeParameterName = src.FormworkTypeParameterName,
                ExportToExcel = src.ExportToExcel,
                CreateSchedule = src.CreateSchedule,
                Create3DView = src.Create3DView,
                CreateSheet = src.CreateSheet,
                ColorScheme = src.ColorScheme,
                ShowDeductedFaces = src.ShowDeductedFaces,
                UseGLDeduction = src.UseGLDeduction,
                GLElevationMeters = src.GLElevationMeters,
                IncludeLinkedModels = src.IncludeLinkedModels,
                EnableDebugLog = src.EnableDebugLog,
                ExcludeSteelMembers = src.ExcludeSteelMembers,
                ExcludeSteelStairs = src.ExcludeSteelStairs,
                ExcludeLgsWalls = src.ExcludeLgsWalls,
                SlopedWallTopWidthThresholdMm = src.SlopedWallTopWidthThresholdMm,
                IncludedCategories = src.IncludedCategories,
                ExcelOutputPath = src.ExcelOutputPath,
            };
        }

        /// <summary>
        /// プロジェクトブラウザで選択されている 3D ビューを収集する。
        /// 1 つも 3D ビューが選択されていない場合は null を返す。
        /// </summary>
        private static List<View3D> CollectSelectedView3Ds(UIDocument uidoc)
        {
            try
            {
                var sel = uidoc.Selection.GetElementIds();
                if (sel == null || sel.Count == 0) return null;

                var doc = uidoc.Document;
                var views = new List<View3D>();
                foreach (var id in sel)
                {
                    var v3d = doc.GetElement(id) as View3D;
                    if (v3d != null && !v3d.IsTemplate) views.Add(v3d);
                }
                return views.Count > 0 ? views : null;
            }
            catch
            {
                return null;
            }
        }
    }
}

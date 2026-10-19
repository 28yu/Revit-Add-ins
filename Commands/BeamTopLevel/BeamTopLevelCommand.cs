using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitApp = Autodesk.Revit.ApplicationServices.Application;
using Tools28.Commands.BeamCommon;
using Tools28.Localization;

namespace Tools28.Commands.BeamTopLevel
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class BeamTopLevelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // ビューテンプレートの解除から色分けの確定までを1つにまとめる。
            // 設定ダイアログでキャンセルした・処理に失敗した場合は、解除も元に戻す
            TransactionGroup templateGroup = null;

            try
            {
                View activeView = doc.ActiveView;

                // チェック1: ビュータイプ（平面ビュー = FloorPlan / 構造伏図 = EngineeringPlan）
                if (activeView.ViewType != ViewType.FloorPlan &&
                    activeView.ViewType != ViewType.EngineeringPlan)
                {
                    TaskDialog.Show(Loc.S("Common.Error"),
                        Loc.S("BeamTop.NotPlanView"));
                    return Result.Cancelled;
                }

                // チェック2: ビューテンプレート
                if (activeView.ViewTemplateId != ElementId.InvalidElementId)
                {
                    TaskDialogResult templateResult = TaskDialog.Show(Loc.S("BeamTop.ViewTemplate"),
                        Loc.S("BeamTop.ViewTemplateMsg"),
                        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                    if (templateResult == TaskDialogResult.Yes)
                    {
                        templateGroup = new TransactionGroup(doc, Loc.S("Common.Txn.ClearViewTemplate"));
                        templateGroup.Start();
                        using (Transaction t = new Transaction(doc, Loc.S("Common.Txn.ClearViewTemplate")))
                        {
                            t.Start();
                            activeView.ViewTemplateId = ElementId.InvalidElementId;
                            t.Commit();
                        }
                    }
                    else
                    {
                        return Result.Cancelled;
                    }
                }

                // チェック3: 梁の取得
                var beams = new FilteredElementCollector(doc, activeView.Id)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    // ⚠ Cast<> は使わない。構造フレームカテゴリには DirectShape 等の
                    //    FamilyInstance でない要素が混ざりうるため、1つでもあると例外で落ちる。
                    .OfType<FamilyInstance>()
                    .ToList();

                if (beams.Count == 0)
                {
                    TaskDialog.Show(Loc.S("Common.Error"), Loc.S("BeamTop.NoBeams"));
                    return Result.Cancelled;
                }

                // 参照レベルを取得
                Level refLevel = activeView.GenLevel;
                if (refLevel == null)
                {
                    TaskDialog.Show(Loc.S("Common.Error"), Loc.S("BeamTop.NoRefLevel"));
                    return Result.Cancelled;
                }

                // ファミリ毎にグループ化
                var beamsByFamily = beams
                    .GroupBy(b => b.Symbol.Family.Name)
                    .ToDictionary(g => g.Key, g => g.ToList());

                // ファミリ毎のパラメータ候補を検索
                var topLevelParamCandidates = BeamCalculator.FindTopLevelParameterCandidates(beamsByFamily);
                var additionalLevelParams = BeamCalculator.FindAdditionalLevelParameters(
                    beamsByFamily, topLevelParamCandidates);

                // TextNoteType一覧を収集
                var textNoteTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .OrderBy(t => t.Name)
                    .Select(t => new TextNoteTypeItem(t))
                    .ToList();

                // ダイアログ表示
                var dialogData = new BeamTopLevelDialogData
                {
                    ViewName = activeView.Name,
                    BeamCount = beams.Count,
                    RefLevel = refLevel,
                    BeamsByFamily = beamsByFamily,
                    TopLevelParamCandidates = topLevelParamCandidates,
                    AdditionalLevelParams = additionalLevelParams,
                    TextNoteTypes = textNoteTypes,
                    // 表記する文字の重複（別の階で同じ文字を使った）を画面で知らせるため
                    LegendExists = name => BeamLegendViews.Find(doc, LegendManager.GetLegendViewName(name)) != null
                };

                var dialog = new BeamTopLevelDialog(dialogData);
                dialog.SetRevitOwner(commandData);
                bool? dialogResult = dialog.ShowDialog();

                if (dialogResult != true)
                    return Result.Cancelled;

                // ダイアログから設定を取得
                Dictionary<string, string> familyTopLevelParamSelection = dialog.FamilyTopLevelParamSelection;
                bool overwriteExisting = dialog.OverwriteExistingFilters;
                ElementId selectedTextNoteTypeId = dialog.SelectedTextNoteTypeId;
                ElementId selectedBeamLabelTypeId = dialog.SelectedBeamLabelTypeId;
                string labelText = dialog.LevelLabelText;

                // 処理実行
                int successCount = 0;
                int failureCount = 0;
                var failedBeams = new List<string>();
                var calculationResults = new Dictionary<ElementId, BeamCalculationResult>();

                // 各梁の計算
                foreach (var beam in beams)
                {
                    string familyName = beam.Symbol.Family.Name;
                    if (!familyTopLevelParamSelection.ContainsKey(familyName))
                    {
                        calculationResults[beam.Id] = new BeamCalculationResult
                        {
                            Success = false,
                            Error = "パラメータが未選択"
                        };
                        failureCount++;
                        failedBeams.Add($"ID:{beam.Id} ({familyName}) - {Loc.S("Beam.ParamNotSelected")}");
                        continue;
                    }

                    var result = BeamCalculator.Calculate(beam,
                        refLevel.Name, labelText, familyTopLevelParamSelection[familyName]);
                    calculationResults[beam.Id] = result;

                    if (result.Success)
                        successCount++;
                    else
                    {
                        failureCount++;
                        failedBeams.Add($"ID:{beam.Id} ({familyName}) - {result.Error}");
                    }
                }

                // レベル別グループ（トランザクション外で計算）
                var levelGroups = calculationResults
                    .Where(kv => kv.Value.Success)
                    .GroupBy(kv => kv.Value.DisplayValue)
                    .ToDictionary(g => g.Key, g => g.Count());

                // パラメータ作成・値の書き込み・フィルタ作成をトランザクションで実行
                BeamLegendResult legend = null;
                using (Transaction trans = new Transaction(doc, Loc.S("BeamTop.Txn.Apply")))
                {
                    trans.Start();

                    // 共有パラメータの作成・バインド
                    RevitApp revitApp = uiapp.Application;
                    ParameterManager.EnsureSharedParameters(doc, revitApp);

                    // 梁への値書き込み
                    foreach (var beam in beams)
                    {
                        var result = calculationResults[beam.Id];
                        ParameterManager.WriteValues(beam, result);
                    }

                    // 梁ラベル（TextNote）の作成
                    if (selectedBeamLabelTypeId != null)
                    {
                        BeamLabelManager.CreateBeamLabels(doc, activeView,
                            beams, calculationResults, selectedBeamLabelTypeId,
                            overwriteExisting);
                    }

                    // フィルタ作成・色分け
                    FilterManager.CreateFiltersAndColorize(doc, activeView,
                        levelGroups, overwriteExisting, failureCount);

                    // 凡例の製図ビューを作成
                    // 同名の凡例を上書きすると画面で確認済みなら、上書きのチェックが無くても作り直す
                    legend = LegendManager.CreateLegendView(
                        doc, levelGroups, overwriteExisting || dialog.OverwriteLegendConfirmed, failureCount,
                        selectedTextNoteTypeId, labelText);

                    trans.Commit();
                    if (templateGroup != null && trans.GetStatus() == TransactionStatus.Committed)
                        templateGroup.Assimilate();
                }

                // 完了ダイアログ
                string legendName = legend?.View != null ? legend.View.Name : LegendManager.GetLegendViewName(labelText);
                string legendInfo;
                if (legend?.Error != null)
                    legendInfo = string.Format(Loc.S("Beam.LegendFailed"), legend.Error);
                else if (legend?.View != null)
                {
                    legendInfo = string.Format(Loc.S("Beam.LegendCreated"), legendName);
                    if (legend.UsedDraftingFallback)
                        legendInfo += Loc.S("Beam.LegendDraftingFallback");
                    if (legend.ReusedDraftingView)
                        legendInfo += Loc.S("Beam.LegendReusedDrafting");
                }
                else if (legend != null && legend.Skipped)
                    legendInfo = string.Format(Loc.S("Beam.LegendSkippedNamed"), legendName);
                else
                    legendInfo = string.Format(Loc.S("Beam.LegendFailed"), "-");

                // 作成した凡例を開く（どこにできたか分かるように）。開けなくても処理は成功扱い
                if (legend?.View != null && legend.Error == null)
                {
                    try { uidoc.ActiveView = legend.View; } catch (Exception) { }
                }

                string resultMessage = string.Format(Loc.S("Beam.Result"),
                    beams.Count, successCount, failureCount, levelGroups.Count, legendInfo);

                if (failureCount > 0)
                {
                    resultMessage += Loc.S("Beam.FailedHeader");
                    foreach (var info in failedBeams.Take(20))
                    {
                        resultMessage += $"\n{info}";
                    }
                    if (failedBeams.Count > 20)
                    {
                        resultMessage += "\n" + string.Format(Loc.S("Import.Done.More"), failedBeams.Count - 20);
                    }
                    resultMessage += Loc.S("Beam.FailedNote");
                }

                TaskDialog.Show(Loc.S("BeamTop.DoneTitle"), resultMessage);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = string.Format(Loc.S("Common.ErrorWithManual"),
                    string.Format(Loc.S("BeamTop.ProcessError"), ex.Message));
                return Result.Failed;
            }
            finally
            {
                if (templateGroup != null)
                {
                    if (templateGroup.GetStatus() == TransactionStatus.Started)
                        templateGroup.RollBack();
                    templateGroup.Dispose();
                }
            }
        }
    }
}

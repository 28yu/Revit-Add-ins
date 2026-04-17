using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitApp = Autodesk.Revit.ApplicationServices.Application;

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

            try
            {
                View activeView = doc.ActiveView;

                // チェック1: ビュータイプ（平面ビュー = FloorPlan / 構造伏図 = EngineeringPlan）
                if (activeView.ViewType != ViewType.FloorPlan &&
                    activeView.ViewType != ViewType.EngineeringPlan)
                {
                    TaskDialog.Show("エラー",
                        "アクティブビューが平面ビューまたは構造伏図ではありません。\n平面ビューまたは構造伏図を開いた状態で実行してください。");
                    return Result.Cancelled;
                }

                // チェック2: ビューテンプレート
                if (activeView.ViewTemplateId != ElementId.InvalidElementId)
                {
                    TaskDialogResult templateResult = TaskDialog.Show("ビューテンプレート確認",
                        "ビューテンプレートが設定されています。\nフィルタを作成するにはテンプレートを解除する必要があります。\n\nテンプレートを解除しますか？",
                        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                    if (templateResult == TaskDialogResult.Yes)
                    {
                        using (Transaction t = new Transaction(doc, "ビューテンプレート解除"))
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
                    .Cast<FamilyInstance>()
                    .ToList();

                if (beams.Count == 0)
                {
                    TaskDialog.Show("エラー", "ビュー内に梁（構造フレーム）が見つかりません。");
                    return Result.Cancelled;
                }

                // 参照レベルを取得
                Level refLevel = activeView.GenLevel;
                if (refLevel == null)
                {
                    TaskDialog.Show("エラー", "ビューの参照レベルが取得できません。");
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
                    TextNoteTypes = textNoteTypes
                };

                var dialog = new BeamTopLevelDialog(dialogData);
                bool? dialogResult = dialog.ShowDialog();

                if (dialogResult != true)
                    return Result.Cancelled;

                // ダイアログから設定を取得
                Dictionary<string, string> familyTopLevelParamSelection = dialog.FamilyTopLevelParamSelection;
                bool overwriteExisting = dialog.OverwriteExistingFilters;
                ElementId selectedTextNoteTypeId = dialog.SelectedTextNoteTypeId;
                ElementId selectedBeamLabelTypeId = dialog.SelectedBeamLabelTypeId;

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
                        failedBeams.Add($"ID:{beam.Id} ({familyName}) - パラメータ未選択");
                        continue;
                    }

                    var result = BeamCalculator.Calculate(beam,
                        refLevel.Name, familyTopLevelParamSelection[familyName]);
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
                ElementId legendViewId = null;
                using (Transaction trans = new Transaction(doc, "梁天端色分け"))
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
                    legendViewId = LegendManager.CreateLegendDraftingView(
                        doc, levelGroups, overwriteExisting, failureCount,
                        selectedTextNoteTypeId);

                    trans.Commit();
                }

                // 完了ダイアログ
                string legendInfo = legendViewId != null
                    ? "凡例ビュー: 「梁天端色分け凡例」を作成しました"
                    : "凡例ビュー: 作成をスキップしました（既存あり）";

                string resultMessage = $"処理が完了しました。\n\n" +
                    $"対象梁数: {beams.Count}\n" +
                    $"成功: {successCount}\n" +
                    $"失敗: {failureCount}\n" +
                    $"作成フィルタ数: {levelGroups.Count}\n" +
                    $"{legendInfo}";

                if (failureCount > 0)
                {
                    resultMessage += "\n\n--- 失敗した梁 ---\n";
                    foreach (var info in failedBeams.Take(20))
                    {
                        resultMessage += $"\n{info}";
                    }
                    if (failedBeams.Count > 20)
                    {
                        resultMessage += $"\n...他 {failedBeams.Count - 20} 件";
                    }
                    resultMessage += "\n\n※ 失敗した梁は赤色フィルタで表示されています。";
                }

                TaskDialog.Show("梁天端色分け - 完了", resultMessage);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"梁天端色分け処理中にエラーが発生しました。\n\n{ex.Message}" +
                    "\n\nマニュアル: https://28tools.com/addins.html" +
                    "\n配布サイト: https://28yu.github.io/28tools-download/";
                return Result.Failed;
            }
        }
    }
}

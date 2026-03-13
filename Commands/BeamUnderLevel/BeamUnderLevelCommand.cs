using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitApp = Autodesk.Revit.ApplicationServices.Application;

namespace Tools28.Commands.BeamUnderLevel
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class BeamUnderLevelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                View activeView = doc.ActiveView;

                // チェック1: ビュータイプ（天井伏図 = CeilingPlan）
                if (activeView.ViewType != ViewType.CeilingPlan)
                {
                    TaskDialog.Show("エラー",
                        "アクティブビューが天井伏図ではありません。\n天井伏図を開いた状態で実行してください。");
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

                // レベル情報を収集
                var allLevels = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(l => l.Elevation)
                    .ToList();

                // アクティブビューの参照レベルを取得
                Level refLevel = activeView.GenLevel;
                if (refLevel == null)
                {
                    TaskDialog.Show("エラー", "ビューの参照レベルが取得できません。");
                    return Result.Cancelled;
                }

                // 上位レベル候補（参照レベルより上のレベル）
                var upperLevels = allLevels
                    .Where(l => l.Elevation > refLevel.Elevation)
                    .OrderBy(l => l.Elevation)
                    .ToList();

                if (upperLevels.Count == 0)
                {
                    TaskDialog.Show("エラー", "参照レベルより上のレベルが見つかりません。");
                    return Result.Cancelled;
                }

                // ファミリ毎にグループ化
                var beamsByFamily = beams
                    .GroupBy(b => b.Symbol.Family.Name)
                    .ToDictionary(g => g.Key, g => g.ToList());

                // ファミリ毎のパラメータ候補を検索
                var paramCandidates = BeamCalculator.FindHeightParameterCandidates(beamsByFamily);

                // ダイアログ表示
                var dialogData = new BeamUnderLevelDialogData
                {
                    ViewName = activeView.Name,
                    BeamCount = beams.Count,
                    RefLevel = refLevel,
                    LowerLevels = upperLevels,
                    DefaultLowerLevel = upperLevels.First(),
                    BeamsByFamily = beamsByFamily,
                    ParamCandidates = paramCandidates
                };

                var dialog = new BeamUnderLevelDialog(dialogData);
                bool? dialogResult = dialog.ShowDialog();

                if (dialogResult != true)
                    return Result.Cancelled;

                // ダイアログから設定を取得
                Level selectedLowerLevel = dialog.SelectedLowerLevel;
                Dictionary<string, string> familyParamSelection = dialog.FamilyParamSelection;
                bool overwriteExisting = dialog.OverwriteExistingFilters;

                // 処理実行
                int successCount = 0;
                int failureCount = 0;
                var failedBeams = new List<string>();
                var calculationResults = new Dictionary<ElementId, BeamCalculationResult>();

                // 階高計算（上位レベル - 参照レベル）
                double floorHeight = selectedLowerLevel.Elevation - refLevel.Elevation;

                // 各梁の計算
                foreach (var beam in beams)
                {
                    string familyName = beam.Symbol.Family.Name;
                    if (!familyParamSelection.ContainsKey(familyName))
                    {
                        calculationResults[beam.Id] = new BeamCalculationResult
                        {
                            Success = false,
                            Error = "梁高さパラメータが未選択"
                        };
                        failureCount++;
                        failedBeams.Add($"ID:{beam.Id} ({familyName}) - パラメータ未選択");
                        continue;
                    }

                    var result = BeamCalculator.Calculate(beam, floorHeight,
                        refLevel.Name, familyParamSelection[familyName]);
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
                using (Transaction trans = new Transaction(doc, "梁下端色分け"))
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

                    // フィルタ作成・色分け
                    FilterManager.CreateFiltersAndColorize(doc, activeView,
                        levelGroups, overwriteExisting);

                    trans.Commit();
                }

                // 完了ダイアログ
                string resultMessage = $"処理が完了しました。\n\n" +
                    $"対象梁数: {beams.Count}\n" +
                    $"成功: {successCount}\n" +
                    $"失敗: {failureCount}\n" +
                    $"作成フィルタ数: {levelGroups.Count}";

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

                TaskDialog.Show("梁下端色分け - 完了", resultMessage);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"梁下端色分け処理中にエラーが発生しました。\n\n{ex.Message}" +
                    "\n\nマニュアル: https://28yu.github.io/28tools-manual/" +
                    "\n配布サイト: https://28yu.github.io/28tools-download/";
                return Result.Failed;
            }
        }
    }
}

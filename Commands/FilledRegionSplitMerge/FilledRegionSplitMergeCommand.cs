using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.FilledRegionSplitMerge
{
    [Transaction(TransactionMode.Manual)]
    public class FilledRegionSplitMergeCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // ステップ1-3: 選択取得と検証
                var selectedIds = uidoc.Selection.GetElementIds();

                // 選択がない場合、ユーザーに選択を促す
                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("塗潰し領域 分割/統合",
                        "塗り潰し領域を選択してから、このコマンドを実行してください。");
                    return Result.Cancelled;
                }

                // ステップ4: 選択された領域の分析
                var analysis = FilledRegionHelper.AnalyzeSelection(doc, selectedIds);

                if (analysis.FilledRegions.Count == 0)
                {
                    TaskDialog.Show("塗潰し領域 分割/統合",
                        "塗り潰し領域が選択されていません。\n塗り潰し領域を選択してから、このコマンドを実行してください。");
                    return Result.Cancelled;
                }

                if (!analysis.CanSplit && !analysis.CanMerge)
                {
                    TaskDialog.Show("塗潰し領域 分割/統合",
                        "分割または統合できる塗り潰し領域がありません。\n\n" +
                        "• 分割: 複数のエリアを持つ領域を選択してください\n" +
                        "• 統合: 複数の領域を選択してください");
                    return Result.Cancelled;
                }

                // ステップ5: ダイアログ表示
                var dialog = new FilledRegionSplitMergeDialog(analysis);
                var dialogResult = dialog.ShowDialog();

                if (dialogResult != true)
                {
                    return Result.Cancelled;
                }

                // ステップ7-9: トランザクション開始と処理実行
                using (Transaction trans = new Transaction(doc, "塗潰し領域 分割/統合"))
                {
                    trans.Start();

                    try
                    {
                        int processedCount = 0;
                        string resultMessage = "";

                        if (dialog.SelectedOperation == FilledRegionSplitMergeDialog.OperationType.Split)
                        {
                            // 分割処理
                            var regionsToSplit = analysis.FilledRegions
                                .Where(fr => fr.GetBoundaries().Count > 1)
                                .ToList();

                            foreach (var fr in regionsToSplit)
                            {
                                int count = FilledRegionHelper.SplitFilledRegion(doc, fr);
                                processedCount += count;
                            }

                            resultMessage = $"{regionsToSplit.Count}個の領域を{processedCount}個の独立した領域に分割しました。";
                        }
                        else // Merge
                        {
                            // 統合処理
                            processedCount = FilledRegionHelper.MergeFilledRegions(
                                doc,
                                analysis.FilledRegions,
                                dialog.SelectedPattern.Id);

                            resultMessage = $"{processedCount}個の領域を1個の領域に統合しました。";
                        }

                        trans.Commit();

                        // ステップ10: 完了メッセージ表示
                        TaskDialog.Show("処理完了", resultMessage);

                        // デバッグログ出力
                        LogToFile($"[{DateTime.Now}] {resultMessage}");

                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;

                // エラーログ出力
                LogToFile($"[{DateTime.Now}] エラー: {ex.Message}\n{ex.StackTrace}");

                TaskDialog.Show("エラー",
                    $"処理中にエラーが発生しました。\n\n{ex.Message}");

                return Result.Failed;
            }
        }

        /// <summary>
        /// デバッグログをファイルに出力
        /// </summary>
        private void LogToFile(string logMessage)
        {
            try
            {
                string logPath = @"C:\temp\Tools28_debug.txt";
                System.IO.File.AppendAllText(logPath, logMessage + "\n");
            }
            catch
            {
                // ログ出力失敗は無視
            }
        }
    }
}

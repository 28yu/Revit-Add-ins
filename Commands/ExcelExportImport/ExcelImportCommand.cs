using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.ExcelExportImport.Services;
using Tools28.Commands.ExcelExportImport.Views;
using Tools28.Localization;

namespace Tools28.Commands.ExcelExportImport
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ExcelImportCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // インポートダイアログを表示
                var dialog = new ImportDialog(doc);
                dialog.SetRevitOwner(commandData);
                bool? result = dialog.ShowDialog();

                if (result != true || !dialog.ImportExecuted)
                    return Result.Cancelled;

                // ダイアログで生成済みのプレビューデータを再利用（色付け用）
                var previewRows = dialog.PreviewRows;

                // トランザクション内でインポート実行。
                // 「部屋の高さは0より大きい必要がある」等の“無視できないエラー”が起きると
                // 1トランザクションでは全体がロールバックされ有効な行も反映されない。
                // そこで失敗要素を収集→除外して再実行し、有効な要素だけを確定する。
                var importResult = new ImportResult();
                var failingIds = new HashSet<long>();
                bool committed = false;

                if (previewRows != null)
                {
                    const int maxPasses = 5;
                    for (int pass = 0; pass < maxPasses; pass++)
                    {
                        using (var trans = new Transaction(doc, "Excelインポート"))
                        {
                            trans.Start();

                            var preproc = new ImportFailurePreprocessor();
                            var opts = trans.GetFailureHandlingOptions();
                            opts.SetFailuresPreprocessor(preproc);
                            opts.SetClearAfterRollback(true);
                            trans.SetFailureHandlingOptions(opts);

                            importResult = ExcelImportService.ImportFromPreview(doc, previewRows, failingIds);

                            // 書き込む変更が無ければコミットしない
                            if (importResult.SuccessCount == 0 && importResult.FailCount == 0)
                            {
                                trans.RollBack();
                                committed = false;
                                break;
                            }

                            var status = trans.Commit(); // ここで失敗プリプロセッサが走る
                            if (status == TransactionStatus.Committed)
                            {
                                committed = true;
                                break;
                            }

                            // 無視できないエラーでロールバック → 失敗要素を除外して再試行
                            int before = failingIds.Count;
                            foreach (var id in preproc.FailingElementIds)
                                failingIds.Add(id);
                            if (failingIds.Count == before)
                                break; // これ以上除外できる要素が無い（収束せず）
                        }
                    }
                }

                // 制約により反映できなかった（除外した）要素を結果に反映
                if (previewRows != null && failingIds.Count > 0)
                {
                    foreach (var pr in previewRows)
                    {
                        if (pr.HasChange && !pr.IsReadOnly && failingIds.Contains(pr.ElementId))
                        {
                            importResult.FailCount++;
                            importResult.FailedCells.Add(pr.ElementId + "|" + pr.ParameterName);
                        }
                    }
                    importResult.Errors.Add(
                        $"{failingIds.Count}個の要素は Revit の制約により変更できずスキップしました" +
                        "（例: 部屋の高さは0より大きい必要があります）。他の要素は反映しました。");
                }

                // コミットできていない場合、成功は0（モデルには反映されていない）
                if (!committed)
                {
                    importResult.SuccessCount = 0;
                }

                // インポート成功時、Excelの変更セルに色を付ける
                string markedFilePath = null;
                string colorMethod = null;
                if ((importResult.SuccessCount > 0 || importResult.FailedCells.Count > 0) && previewRows != null)
                {
                    try
                    {
                        markedFilePath = ExcelImportService.MarkImportedCells(
                            dialog.SelectedFilePath, previewRows, out colorMethod, importResult.FailedCells);

                        // 別名保存された場合、色付きファイルを自動で開く
                        if (markedFilePath != null && markedFilePath != dialog.SelectedFilePath)
                        {
                            Process.Start(new ProcessStartInfo(markedFilePath) { UseShellExecute = true });
                        }
                    }
                    catch
                    {
                        // Excel色付けの失敗はインポート結果に影響させない
                    }
                }

                // 結果表示
                var sb = new StringBuilder();
                sb.AppendLine("インポートが完了しました。\n");
                sb.AppendLine($"成功: {importResult.SuccessCount}件");
                sb.AppendLine($"失敗: {importResult.FailCount}件");
                sb.AppendLine($"スキップ: {importResult.SkipCount}件");

                // 色付け結果を表示
                if (markedFilePath != null)
                {
                    if (markedFilePath != dialog.SelectedFilePath)
                    {
                        sb.AppendLine($"\n※ 変更行に色を付けたファイルを別名で保存しました:");
                        sb.AppendLine(markedFilePath);
                    }
                    else if (colorMethod == "COM")
                    {
                        sb.AppendLine("\n※ 開いているExcelの変更行に色を付けました。");
                    }
                    else
                    {
                        sb.AppendLine("\n※ Excelファイルの変更行に色を付けました。");
                    }
                }

                if (importResult.Errors.Count > 0)
                {
                    sb.AppendLine("\n--- エラー詳細 ---");
                    int maxErrors = Math.Min(importResult.Errors.Count, 20);
                    for (int i = 0; i < maxErrors; i++)
                    {
                        sb.AppendLine(importResult.Errors[i]);
                    }
                    if (importResult.Errors.Count > 20)
                    {
                        sb.AppendLine($"... 他 {importResult.Errors.Count - 20}件");
                    }
                }

                if (importResult.Warnings.Count > 0)
                {
                    sb.AppendLine("\n--- 警告 ---");
                    int maxWarnings = Math.Min(importResult.Warnings.Count, 10);
                    for (int i = 0; i < maxWarnings; i++)
                    {
                        sb.AppendLine(importResult.Warnings[i]);
                    }
                    if (importResult.Warnings.Count > 10)
                    {
                        sb.AppendLine($"... 他 {importResult.Warnings.Count - 10}件");
                    }
                }

                TaskDialog.Show(Loc.S("Import.ResultTitle"), sb.ToString());

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message + "\n\nマニュアル: https://28tools.com/addins.html";
                return Result.Failed;
            }
        }
    }
}

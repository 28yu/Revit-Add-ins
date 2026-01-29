using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.SheetCreation
{
    /// <summary>
    /// シート一括作成コマンド
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ExecuteSheetCreationCommand : IExternalCommand
    {
        // デバッグログ用
        private void LogDebug(string message)
        {
            try
            {
                System.IO.Directory.CreateDirectory(@"C:\temp");
                System.IO.File.AppendAllText(@"C:\temp\Tools28_debug.txt",
                    DateTime.Now.ToString("HH:mm:ss.fff") + ": " + message + "\n");
            }
            catch { }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                LogDebug("========================================");
                LogDebug("=== シート一括作成コマンド実行開始 ===");
                LogDebug("========================================");

                // 図枠の存在確認
                var titleBlocks = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .ToList();

                LogDebug($"プロジェクト内の図枠数: {titleBlocks.Count}");

                if (titleBlocks.Count == 0)
                {
                    LogDebug("エラー: 図枠が見つかりません");

                    TaskDialog errorDialog = new TaskDialog("エラー");
                    errorDialog.MainInstruction = "図枠ファミリがロードされていません";
                    errorDialog.MainContent = "プロジェクトに図枠ファミリをロードしてから実行してください。\n\n" +
                                              "「挿入」タブ → 「ファミリのロード」→ タイトルブロック";
                    errorDialog.CommonButtons = TaskDialogCommonButtons.Ok;
                    errorDialog.Show();
                    return Result.Cancelled;
                }

                // ダイアログを表示
                LogDebug("ダイアログ表示");
                SheetCreationDialog dialog = new SheetCreationDialog(doc);
                bool? dialogResult = dialog.ShowDialog();

                LogDebug($"ダイアログ結果: {dialogResult}");

                if (dialogResult != true)
                {
                    LogDebug("ユーザーがキャンセル");
                    return Result.Cancelled;
                }

                // null チェック
                if (dialog.SelectedTitleBlock == null)
                {
                    LogDebug("エラー: SelectedTitleBlock が null");
                    message = "図枠が選択されていません。";
                    return Result.Failed;
                }

                // 選択された設定を取得
                FamilySymbol titleBlock = dialog.SelectedTitleBlock.Symbol;

                if (titleBlock == null)
                {
                    LogDebug("エラー: titleBlock.Symbol が null");
                    message = "図枠の取得に失敗しました。";
                    return Result.Failed;
                }

                int sheetCount = dialog.SheetCount;
                string prefix = dialog.Prefix ?? "";

                LogDebug($"図枠: {titleBlock.FamilyName} - {titleBlock.Name}");
                LogDebug($"作成枚数: {sheetCount}");
                LogDebug($"プレフィックス: '{prefix}'");

                // シートを作成
                using (Transaction trans = new Transaction(doc, "シート一括作成"))
                {
                    trans.Start();
                    LogDebug("トランザクション開始");

                    // 図枠をアクティブ化
                    if (!titleBlock.IsActive)
                    {
                        LogDebug("図枠をアクティブ化");
                        titleBlock.Activate();
                        doc.Regenerate();
                    }

                    // 次のシート番号を取得
                    int nextNumber = GetNextSheetNumber(doc, prefix);
                    LogDebug($"開始シート番号: {nextNumber}");

                    // シートを作成
                    List<ViewSheet> createdSheets = new List<ViewSheet>();
                    for (int i = 0; i < sheetCount; i++)
                    {
                        int currentNumber = nextNumber + i;
                        string sheetNumber = FormatSheetNumber(prefix, currentNumber);
                        string sheetName = $"新規シート {sheetNumber}";

                        LogDebug($"シート作成中 [{i + 1}/{sheetCount}]: {sheetNumber}");

                        ViewSheet sheet = ViewSheet.Create(doc, titleBlock.Id);
                        sheet.SheetNumber = sheetNumber;
                        sheet.Name = sheetName;

                        createdSheets.Add(sheet);
                    }

                    trans.Commit();
                    LogDebug("トランザクションコミット完了");

                    // 結果を表示
                    ShowResultDialog(createdSheets, titleBlock);
                    LogDebug("結果ダイアログ表示完了");
                }

                LogDebug("=== シート一括作成コマンド正常終了 ===");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                LogDebug($"=== エラー発生 ===");
                LogDebug($"メッセージ: {ex.Message}");
                LogDebug($"スタックトレース: {ex.StackTrace}");

                message = GetErrorMessageWithManualUrl($"シート作成中にエラーが発生しました。\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        /// <summary>
        /// シート番号をフォーマット
        /// </summary>
        private string FormatSheetNumber(string prefix, int number)
        {
            if (string.IsNullOrEmpty(prefix))
            {
                return $"- {number}";
            }
            else
            {
                return $"{prefix} - {number}";
            }
        }

        /// <summary>
        /// 次のシート番号を取得
        /// </summary>
        private int GetNextSheetNumber(Document doc, string prefix)
        {
            var existingSheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Select(s => s.SheetNumber)
                .ToList();

            int maxNumber = 0;

            foreach (string sheetNumber in existingSheets)
            {
                string numberPart = ExtractNumberFromSheetNumber(sheetNumber, prefix);

                if (int.TryParse(numberPart, out int number))
                {
                    maxNumber = Math.Max(maxNumber, number);
                }
            }

            return maxNumber + 1;
        }

        /// <summary>
        /// シート番号から数値部分を抽出
        /// </summary>
        private string ExtractNumberFromSheetNumber(string sheetNumber, string prefix)
        {
            if (!string.IsNullOrEmpty(prefix))
            {
                string searchPattern = $"{prefix} - ";
                if (sheetNumber.StartsWith(searchPattern))
                {
                    return sheetNumber.Substring(searchPattern.Length).Trim();
                }
            }

            if (sheetNumber.StartsWith("- "))
            {
                return sheetNumber.Substring(2).Trim();
            }

            if (!string.IsNullOrEmpty(prefix))
            {
                string oldPattern = $"{prefix}- ";
                if (sheetNumber.StartsWith(oldPattern))
                {
                    return sheetNumber.Substring(oldPattern.Length).Trim();
                }
            }

            int hyphenIndex = sheetNumber.LastIndexOf('-');
            if (hyphenIndex >= 0)
            {
                string afterHyphen = sheetNumber.Substring(hyphenIndex + 1).Trim();
                string trimmed = afterHyphen.TrimStart('0');
                return string.IsNullOrEmpty(trimmed) ? "1" : trimmed;
            }

            string result = sheetNumber.TrimStart('0');
            return string.IsNullOrEmpty(result) ? "1" : result;
        }

        /// <summary>
        /// 結果ダイアログを表示
        /// </summary>
        private void ShowResultDialog(List<ViewSheet> createdSheets, FamilySymbol titleBlock)
        {
            TaskDialog resultDialog = new TaskDialog("シート作成完了");
            resultDialog.MainInstruction = $"{createdSheets.Count}枚のシートを作成しました";

            string detailText = $"図枠: {titleBlock.FamilyName} - {titleBlock.Name}\n";
            detailText += $"シート番号: {createdSheets.First().SheetNumber}";
            if (createdSheets.Count > 1)
            {
                detailText += $" ～ {createdSheets.Last().SheetNumber}";
            }

            resultDialog.MainContent = detailText;
            resultDialog.CommonButtons = TaskDialogCommonButtons.Ok;
            resultDialog.Show();
        }

        /// <summary>
        /// エラーメッセージにマニュアルURLを含める
        /// </summary>
        private string GetErrorMessageWithManualUrl(string errorMessage)
        {
            return $"{errorMessage}\n\nマニュアル: https://28yu.github.io/28tools-manual/\n配布サイト: https://28yu.github.io/28tools-download/\nFor English: Click 🌐 button on the manual page";
        }
    }
}
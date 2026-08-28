using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Localization;

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
            DiagLog.Cmd("SheetCreation", "Execute 開始");
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                DiagLog.Cmd("SheetCreation", "図枠検索開始");
                // 図枠の存在確認
                var titleBlocks = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .ToList();

                DiagLog.Cmd("SheetCreation", $"図枠数: {titleBlocks.Count}");

                if (titleBlocks.Count == 0)
                {
                    LogDebug("エラー: 図枠が見つかりません");

                    TaskDialog errorDialog = new TaskDialog(Loc.S("Common.Error"));
                    errorDialog.MainInstruction = Loc.S("Sheet.NoTitleBlockLoaded.Main");
                    errorDialog.MainContent = Loc.S("Sheet.NoTitleBlockLoaded.Content");
                    errorDialog.CommonButtons = TaskDialogCommonButtons.Ok;
                    errorDialog.Show();
                    return Result.Cancelled;
                }

                DiagLog.Cmd("SheetCreation", "Dialog コンストラクタ 直前");
                SheetCreationDialog dialog = new SheetCreationDialog(doc);
                DiagLog.Cmd("SheetCreation", "Dialog コンストラクタ 完了");
                dialog.SetRevitOwner(commandData);
                DiagLog.Cmd("SheetCreation", "SetRevitOwner 完了, ShowDialog 直前");
                bool? dialogResult = dialog.ShowDialog();
                DiagLog.Cmd("SheetCreation", $"ShowDialog 戻り: {dialogResult}");

                if (dialogResult != true)
                {
                    LogDebug("ユーザーがキャンセル");
                    return Result.Cancelled;
                }

                // null チェック
                if (dialog.SelectedTitleBlock == null)
                {
                    LogDebug("エラー: SelectedTitleBlock が null");
                    message = Loc.S("Sheet.NoTitleBlockSelected");
                    return Result.Failed;
                }

                // 選択された設定を取得
                FamilySymbol titleBlock = dialog.SelectedTitleBlock.Symbol;

                if (titleBlock == null)
                {
                    LogDebug("エラー: titleBlock.Symbol が null");
                    message = Loc.S("Sheet.TitleBlockFailed");
                    return Result.Failed;
                }

                int sheetCount = dialog.SheetCount;
                string prefix = dialog.Prefix ?? "";

                // シート名はモデルに保存され、そのモデルを開く全員が見る文字列。
                // アドインの言語設定ではなく Revit 本体の UI 言語に合わせる
                // （カテゴリ名や既定ビュー名など Revit が生成する他の名前と揃えるため）。
                string modelLang = RevitUiLanguage.Resolve(commandData.Application);

                LogDebug($"図枠: {titleBlock.FamilyName} - {titleBlock.Name}");
                LogDebug($"作成枚数: {sheetCount}");
                LogDebug($"プレフィックス: '{prefix}'");

                // シートを作成
                using (Transaction trans = new Transaction(doc, Loc.S("Sheet.Txn.Create")))
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
                        string sheetName = string.Format(Loc.S("Sheet.NewSheetName", modelLang), sheetNumber);

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

                message = GetErrorMessageWithManualUrl(string.Format(Loc.S("Sheet.ProcessError"), ex.Message));
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
        /// 指定プレフィックスにおける次のシート番号を取得する。
        /// ⚠ 対象は「同じプレフィックスのシート」だけ。
        ///    他プレフィックス（"B - 57"）や独自形式（"S-101"）まで数えると採番が飛ぶ。
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
                if (numberPart == null) continue;   // プレフィックスが違う＝対象外

                if (int.TryParse(numberPart, out int number))
                {
                    maxNumber = Math.Max(maxNumber, number);
                }
            }

            return maxNumber + 1;
        }

        /// <summary>
        /// シート番号から数値部分を抽出する。
        /// 指定プレフィックスの形式に一致しない場合は null を返す（採番の対象外）。
        /// </summary>
        private string ExtractNumberFromSheetNumber(string sheetNumber, string prefix)
        {
            if (string.IsNullOrEmpty(sheetNumber)) return null;

            if (!string.IsNullOrEmpty(prefix))
            {
                string current = prefix + " - ";
                if (sheetNumber.StartsWith(current))
                    return sheetNumber.Substring(current.Length).Trim();

                string legacy = prefix + "- ";   // 旧形式（スペース無し）
                if (sheetNumber.StartsWith(legacy))
                    return sheetNumber.Substring(legacy.Length).Trim();

                return null;
            }

            // プレフィックス無し（"- 12" 形式）
            if (sheetNumber.StartsWith("- "))
                return sheetNumber.Substring(2).Trim();

            return null;
        }

        /// <summary>
        /// 結果ダイアログを表示
        /// </summary>
        private void ShowResultDialog(List<ViewSheet> createdSheets, FamilySymbol titleBlock)
        {
            if (createdSheets == null || createdSheets.Count == 0) return;

            TaskDialog resultDialog = new TaskDialog(Loc.S("Sheet.Result.Title"));
            resultDialog.MainInstruction = string.Format(Loc.S("Sheet.Result.Main"), createdSheets.Count);

            string titleBlockName = $"{titleBlock.FamilyName} - {titleBlock.Name}";
            resultDialog.MainContent = createdSheets.Count > 1
                ? string.Format(Loc.S("Sheet.Result.DetailRange"), titleBlockName,
                    createdSheets.First().SheetNumber, createdSheets.Last().SheetNumber)
                : string.Format(Loc.S("Sheet.Result.Detail"), titleBlockName,
                    createdSheets.First().SheetNumber);

            resultDialog.CommonButtons = TaskDialogCommonButtons.Ok;
            resultDialog.Show();
        }

        /// <summary>
        /// エラーメッセージにマニュアルURLを含める
        /// </summary>
        private string GetErrorMessageWithManualUrl(string errorMessage)
        {
            return $"{errorMessage}\n\nマニュアル: https://28tools.com/addins.html\n配布サイト: https://28yu.github.io/28tools-download/\nFor English: Click 🌐 button on the manual page";
        }
    }
}
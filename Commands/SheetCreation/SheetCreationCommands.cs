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
        /// <summary>デバッグログ。出力先・ローテーション・ON/OFF は DiagLog に一元化している。</summary>
        private void LogDebug(string message) => DiagLog.Write("[SheetCreation] " + message);

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

                    List<ViewSheet> createdSheets;
                    var skippedNumbers = new List<string>();

                    if (dialog.UseListMode)
                    {
                        LogDebug($"リストモード: {dialog.SheetList.Count} 行");
                        createdSheets = CreateSheetsFromList(
                            doc, titleBlock, dialog.SheetList, modelLang, skippedNumbers);
                    }
                    else
                    {
                        // 次のシート番号を取得
                        int nextNumber = GetNextSheetNumber(doc, prefix);
                        LogDebug($"開始シート番号: {nextNumber}");
                        createdSheets = CreateSheetsByCount(
                            doc, titleBlock, prefix, sheetCount, nextNumber, modelLang);
                    }

                    trans.Commit();
                    LogDebug("トランザクションコミット完了");

                    // 結果を表示
                    ShowResultDialog(createdSheets, titleBlock, skippedNumbers);
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
        /// 「枚数＋図面No」モード。連番でシートを作る。トランザクション内で呼ぶこと。
        /// </summary>
        private List<ViewSheet> CreateSheetsByCount(
            Document doc, FamilySymbol titleBlock, string prefix, int sheetCount,
            int nextNumber, string modelLang)
        {
            var created = new List<ViewSheet>();

            for (int i = 0; i < sheetCount; i++)
            {
                string sheetNumber = FormatSheetNumber(prefix, nextNumber + i);
                string sheetName = string.Format(Loc.S("Sheet.NewSheetName", modelLang), sheetNumber);

                LogDebug($"シート作成中 [{i + 1}/{sheetCount}]: {sheetNumber}");

                ViewSheet sheet = ViewSheet.Create(doc, titleBlock.Id);
                sheet.SheetNumber = sheetNumber;
                sheet.Name = sheetName;
                created.Add(sheet);
            }

            return created;
        }

        /// <summary>
        /// 「リスト入力」モード。指定された番号・名前でシートを作る。トランザクション内で呼ぶこと。
        /// すでに同じシート番号が存在する行はスキップし、skipped に記録する
        /// （Revit はシート番号の重複を許さないため、そのまま設定すると例外になる）。
        /// シート名が空の行は連番モードと同じ既定名にする。
        /// </summary>
        private List<ViewSheet> CreateSheetsFromList(
            Document doc, FamilySymbol titleBlock, List<SheetListEntry> entries,
            string modelLang, List<string> skipped)
        {
            var created = new List<ViewSheet>();
            if (entries == null || entries.Count == 0) return created;

            // 既存シート番号（大文字小文字を区別しない）。Revit の重複判定に合わせる。
            var existingNumbers = new HashSet<string>(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Select(v => v.SheetNumber ?? ""),
                StringComparer.CurrentCultureIgnoreCase);

            for (int i = 0; i < entries.Count; i++)
            {
                var entry = entries[i];
                if (entry == null || string.IsNullOrWhiteSpace(entry.Number)) continue;

                if (!existingNumbers.Add(entry.Number))
                {
                    LogDebug($"シート番号が重複のためスキップ: {entry.Number}");
                    skipped.Add(entry.Number);
                    continue;
                }

                LogDebug($"シート作成中 [{i + 1}/{entries.Count}]: {entry.Number}");

                ViewSheet sheet = ViewSheet.Create(doc, titleBlock.Id);
                sheet.SheetNumber = entry.Number;
                sheet.Name = string.IsNullOrWhiteSpace(entry.Name)
                    ? string.Format(Loc.S("Sheet.NewSheetName", modelLang), entry.Number)
                    : entry.Name;

                created.Add(sheet);
            }

            return created;
        }

        /// <summary>
        /// シート番号をフォーマットする。
        /// 図面No（プレフィックス）が空欄なら番号だけにする。
        /// （旧実装は空欄でも先頭に "- " を付けており、"- 1" という番号になっていた）
        /// </summary>
        private string FormatSheetNumber(string prefix, int number)
        {
            return string.IsNullOrEmpty(prefix)
                ? number.ToString()
                : $"{prefix} - {number}";
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

            // プレフィックス無し。現行形式は番号のみ（"12"）。
            if (int.TryParse(sheetNumber.Trim(), out _))
                return sheetNumber.Trim();

            // 旧形式（"- 12"）。旧バージョンで作ったシートからも採番を続けられるようにする。
            if (sheetNumber.StartsWith("- "))
                return sheetNumber.Substring(2).Trim();

            return null;
        }

        /// <summary>
        /// 結果ダイアログを表示
        /// </summary>
        private void ShowResultDialog(List<ViewSheet> createdSheets, FamilySymbol titleBlock,
                                      List<string> skippedNumbers)
        {
            bool hasCreated = createdSheets != null && createdSheets.Count > 0;
            bool hasSkipped = skippedNumbers != null && skippedNumbers.Count > 0;
            if (!hasCreated && !hasSkipped) return;

            TaskDialog resultDialog = new TaskDialog(Loc.S("Sheet.Result.Title"));
            resultDialog.MainInstruction = string.Format(
                Loc.S("Sheet.Result.Main"), hasCreated ? createdSheets.Count : 0);

            string content = "";
            if (hasCreated)
            {
                string titleBlockName = $"{titleBlock.FamilyName} - {titleBlock.Name}";
                content = createdSheets.Count > 1
                    ? string.Format(Loc.S("Sheet.Result.DetailRange"), titleBlockName,
                        createdSheets.First().SheetNumber, createdSheets.Last().SheetNumber)
                    : string.Format(Loc.S("Sheet.Result.Detail"), titleBlockName,
                        createdSheets.First().SheetNumber);
            }

            // 既存と重複してスキップした番号を明示する（黙って減らさない）
            if (hasSkipped)
            {
                if (content.Length > 0) content += "\n\n";
                content += string.Format(Loc.S("Sheet.Result.Skipped"), skippedNumbers.Count)
                         + "\n" + string.Join(", ", skippedNumbers.Take(10));
                if (skippedNumbers.Count > 10) content += " …";
            }

            resultDialog.MainContent = content;
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
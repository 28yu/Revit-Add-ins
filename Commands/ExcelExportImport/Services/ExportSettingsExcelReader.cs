using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using Tools28.Commands.ExcelExportImport.Models;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// このアドインで書き出した Excel ファイルから「出力設定」（対象カテゴリ・
    /// 出力パラメータ・順序）を復元するサービス。
    /// 設定JSONを保存し忘れた場合や、他ユーザーが書き出した Excel の並びを
    /// そのまま再利用したい場合に使う。
    /// </summary>
    /// <remarks>
    /// エクスポート形式（<see cref="ExcelExportService"/>）を前提に読み取る:
    ///  - 1行目がヘッダー。1列目「要素ID」、2列目「カテゴリ」、3列目以降が
    ///    パラメータ列で、見出しは "I-"/"T-" プレフィックス付き DisplayName
    ///    （読取専用は "(*変更不可)" サフィックス付き）。
    ///  - データ行の2列目には Revit の実カテゴリ名が入る。
    ///  - カテゴリ毎シート分割: 各シート＝1カテゴリ。
    ///  - 1シート統合（"データ"）: 2列目で複数カテゴリが混在。
    /// </remarks>
    public static class ExportSettingsExcelReader
    {
        // ExcelExportService が出力する固定ヘッダー名（1列目・2列目）。
        // 列位置ではなくこの見出しで列を特定するため、列を入れ替えても読める。
        private const string ElementIdHeader = "要素ID";
        private const string CategoryHeader = "カテゴリ";

        /// <summary>
        /// エクスポート済み Excel から出力設定を復元する。
        /// パラメータ列を1つも読み取れなかった場合は OutputParameters が空になる。
        /// </summary>
        public static ExportSettings ReadFromExcel(string filePath)
        {
            var settings = new ExportSettings();

            // カテゴリ・出力パラメータを「出現順」を保ったまま重複なく蓄積する
            var categoryOrder = new List<string>();
            var categorySet = new HashSet<string>();
            var entryKeys = new HashSet<string>();

            void RegisterCategory(string cat)
            {
                if (!string.IsNullOrWhiteSpace(cat) && categorySet.Add(cat))
                    categoryOrder.Add(cat);
            }

            void AddEntry(string cat, string rawName, bool isType, string displayName)
            {
                // 同名パラメータを区別するため、キーには表示名（接尾辞込み）を使う
                string key = displayName + "|" + cat;
                if (!entryKeys.Add(key)) return;
                settings.OutputParameters.Add(new ExportParameterEntry
                {
                    RawName = rawName,
                    IsTypeParameter = isType,
                    CategoryName = cat,
                    DisplayName = displayName
                });
            }

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    // ヘッダー（1行目）だけを読む（全データ走査は行わない）。
                    // 列位置は固定で仮定せず、見出し文字列で「要素ID列」「カテゴリ列」を
                    // 特定し、それ以外で I-/T- プレフィックスを持つ列をパラメータ列とする。
                    // → ユーザーが Excel 上で列を入れ替え・並べ替えても正しく読める。
                    var headerParams = new List<ParsedHeader>();
                    int categoryCol = -1;
                    foreach (var cell in worksheet.Row(1).CellsUsed())
                    {
                        string text = cell.GetString();
                        if (text == ElementIdHeader)
                            continue; // 要素ID列（設定復元には不要）
                        if (text == CategoryHeader)
                        {
                            categoryCol = cell.Address.ColumnNumber; // カテゴリ列を記録
                            continue;
                        }
                        var parsed = ParseHeader(text);
                        if (parsed != null)
                        {
                            parsed.Column = cell.Address.ColumnNumber;
                            headerParams.Add(parsed);
                        }
                    }
                    if (headerParams.Count == 0)
                        continue; // このシートに有効なパラメータ列が無い

                    // 統合シート（ExportSingleSheet が付ける固定名 "データ"）だけは
                    // 複数カテゴリが混在するため列→カテゴリの対応判定が要る。
                    // それ以外（カテゴリ毎シート分割）は 1シート=1カテゴリなので、
                    // 先頭データ行(2行目)のカテゴリ列からカテゴリ名を1回読むだけで済む。
                    if (worksheet.Name != "データ")
                    {
                        // カテゴリ列が見つかればその値、無ければシート名（サニタイズ済みの実名）を使う
                        string cat = null;
                        if (categoryCol > 0)
                            cat = worksheet.Cell(2, categoryCol).GetString();
                        if (string.IsNullOrWhiteSpace(cat))
                            cat = worksheet.Name; // カテゴリ列なし/データ行なし → シート名で代用
                        RegisterCategory(cat);
                        foreach (var h in headerParams)
                            AddEntry(cat, h.RawName, h.IsType, h.DisplayName);
                    }
                    else
                    {
                        // カテゴリ列が見つからない場合のみ従来の2列目にフォールバック
                        int catCol = categoryCol > 0 ? categoryCol : 2;
                        ReadSingleSheet(worksheet, headerParams, catCol, RegisterCategory, AddEntry);
                    }
                }
            }

            settings.SelectedCategories = categoryOrder;
            return settings;
        }

        /// <summary>
        /// 1シート統合形式（"データ"）の読み取り。複数カテゴリが混在するため、
        /// 2列目のカテゴリと各パラメータ列の非空判定で列→カテゴリを対応付ける。
        /// （この形式は既定ではないため、必要時のみ全行走査する）
        /// </summary>
        private static void ReadSingleSheet(
            IXLWorksheet worksheet,
            List<ParsedHeader> headerParams,
            int categoryCol,
            System.Action<string> registerCategory,
            System.Action<string, string, bool, string> addEntry)
        {
            var lastRow = worksheet.LastRowUsed();
            int rowCount = lastRow?.RowNumber() ?? 1;

            var sheetCategories = new List<string>();
            var sheetCatSet = new HashSet<string>();
            var nonEmpty = new HashSet<string>(); // key: "cat|column"

            for (int row = 2; row <= rowCount; row++)
            {
                var r = worksheet.Row(row);
                string cat = r.Cell(categoryCol).GetString();
                if (string.IsNullOrWhiteSpace(cat)) continue;

                if (sheetCatSet.Add(cat))
                    sheetCategories.Add(cat);

                foreach (var h in headerParams)
                {
                    string key = cat + "|" + h.Column;
                    if (nonEmpty.Contains(key)) continue;
                    if (!string.IsNullOrEmpty(r.Cell(h.Column).GetString()))
                        nonEmpty.Add(key);
                }
            }

            foreach (var cat in sheetCategories)
            {
                registerCategory(cat);
                foreach (var h in headerParams)
                    if (nonEmpty.Contains(cat + "|" + h.Column))
                        addEntry(cat, h.RawName, h.IsType, h.DisplayName);
            }
        }

        /// <summary>ヘッダー見出しをパースする。パラメータ列でなければ null。</summary>
        private static ParsedHeader ParseHeader(string header)
        {
            if (string.IsNullOrWhiteSpace(header)) return null;

            // 編集可否マーカー（変更不可/画像参照/要素参照）を除去 → 表示名（接尾辞込み）
            string displayName = ParameterHeaderMarker.Strip(header);

            // プレフィックスが無い列（要素ID/カテゴリ等）はパラメータ列ではない
            if (!displayName.StartsWith("T-") && !displayName.StartsWith("I-"))
                return null;

            // 同名区別の接尾辞（【組み込み】等）を分離して生パラメータ名・種別を得る
            var parsed = ParameterService.ParseDisplayName(displayName);
            return new ParsedHeader
            {
                RawName = parsed.RawName,
                IsType = parsed.IsTypeParameter,
                DisplayName = displayName
            };
        }

        private class ParsedHeader
        {
            public string RawName;
            public bool IsType;
            public int Column;
            public string DisplayName;
        }
    }
}

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
        private const string ReadOnlySuffix = "(*変更不可)";

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

            void AddEntry(string cat, string rawName, bool isType)
            {
                string key = rawName + "|" + isType + "|" + cat;
                if (!entryKeys.Add(key)) return;
                settings.OutputParameters.Add(new ExportParameterEntry
                {
                    RawName = rawName,
                    IsTypeParameter = isType,
                    CategoryName = cat
                });
            }

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    // ヘッダーは1行目だけを読む（全データ走査は行わない）。
                    // 3列目以降がパラメータ列。ClosedXML のセルアクセスは重いため、
                    // Row(1).CellsUsed() で1行目の使用セルだけを一度に走査する。
                    var headerParams = new List<ParsedHeader>();
                    foreach (var cell in worksheet.Row(1).CellsUsed())
                    {
                        int col = cell.Address.ColumnNumber;
                        if (col < 3) continue; // 1列目=要素ID, 2列目=カテゴリ
                        var parsed = ParseHeader(cell.GetString());
                        if (parsed != null)
                        {
                            parsed.Column = col;
                            headerParams.Add(parsed);
                        }
                    }
                    if (headerParams.Count == 0)
                        continue; // このシートに有効なパラメータ列が無い

                    // 統合シート（ExportSingleSheet が付ける固定名 "データ"）だけは
                    // 複数カテゴリが混在するため列→カテゴリの対応判定が要る。
                    // それ以外（カテゴリ毎シート分割）は 1シート=1カテゴリなので、
                    // 先頭データ行(2行目)の2列目からカテゴリ名を1回読むだけで済む。
                    if (worksheet.Name != "データ")
                    {
                        string cat = worksheet.Cell(2, 2).GetString();
                        if (string.IsNullOrWhiteSpace(cat))
                            cat = worksheet.Name; // データ行なし（要素0件）→ シート名で代用
                        RegisterCategory(cat);
                        foreach (var h in headerParams)
                            AddEntry(cat, h.RawName, h.IsType);
                    }
                    else
                    {
                        ReadSingleSheet(worksheet, headerParams, RegisterCategory, AddEntry);
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
            System.Action<string> registerCategory,
            System.Action<string, string, bool> addEntry)
        {
            var lastRow = worksheet.LastRowUsed();
            int rowCount = lastRow?.RowNumber() ?? 1;

            var sheetCategories = new List<string>();
            var sheetCatSet = new HashSet<string>();
            var nonEmpty = new HashSet<string>(); // key: "cat|column"

            for (int row = 2; row <= rowCount; row++)
            {
                var r = worksheet.Row(row);
                string cat = r.Cell(2).GetString();
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
                        addEntry(cat, h.RawName, h.IsType);
            }
        }

        /// <summary>ヘッダー見出しをパースする。パラメータ列でなければ null。</summary>
        private static ParsedHeader ParseHeader(string header)
        {
            if (string.IsNullOrWhiteSpace(header)) return null;

            string name = header;
            if (name.EndsWith(ReadOnlySuffix))
                name = name.Substring(0, name.Length - ReadOnlySuffix.Length);

            if (name.StartsWith("T-"))
                return new ParsedHeader { RawName = name.Substring(2), IsType = true };
            if (name.StartsWith("I-"))
                return new ParsedHeader { RawName = name.Substring(2), IsType = false };

            // プレフィックスが無い列（要素ID/カテゴリ等）はパラメータ列ではない
            return null;
        }

        private class ParsedHeader
        {
            public string RawName;
            public bool IsType;
            public int Column;
        }
    }
}

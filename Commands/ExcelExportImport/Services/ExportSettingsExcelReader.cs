using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                    var lastRow = worksheet.LastRowUsed();
                    var lastCol = worksheet.LastColumnUsed();
                    if (lastRow == null || lastCol == null)
                        continue;

                    int rowCount = lastRow.RowNumber();
                    int colCount = lastCol.ColumnNumber();
                    if (colCount < 3)
                        continue; // パラメータ列（3列目以降）が無い

                    // ヘッダー（3列目以降）をパース。I-/T- プレフィックスの無い列は
                    // パラメータ列ではないとみなしてスキップ（null で場所だけ確保）。
                    var headerParams = new List<ParsedHeader>();
                    for (int col = 3; col <= colCount; col++)
                    {
                        headerParams.Add(ParseHeader(worksheet.Cell(1, col).GetString()));
                    }
                    if (headerParams.All(h => h == null))
                        continue; // このシートに有効なパラメータ列が無い

                    // データ行の2列目から実カテゴリ名を出現順に収集
                    var sheetCategories = new List<string>();
                    var sheetCatSet = new HashSet<string>();
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string cat = worksheet.Cell(row, 2).GetString();
                        if (!string.IsNullOrWhiteSpace(cat) && sheetCatSet.Add(cat))
                            sheetCategories.Add(cat);
                    }

                    // データ行が無い（要素0件）シートはシート名をカテゴリ名として代用
                    if (sheetCategories.Count == 0)
                        sheetCategories.Add(worksheet.Name);

                    if (sheetCategories.Count == 1)
                    {
                        // カテゴリ毎シート（または単一カテゴリの統合シート）
                        string cat = sheetCategories[0];
                        RegisterCategory(cat);
                        foreach (var h in headerParams)
                        {
                            if (h == null) continue;
                            AddEntry(cat, h.RawName, h.IsType);
                        }
                    }
                    else
                    {
                        // 1シート統合: どのパラメータ列がどのカテゴリに属すかは
                        // 「そのカテゴリの行に非空セルがあるか」で判定する
                        foreach (var cat in sheetCategories)
                            RegisterCategory(cat);

                        var nonEmpty = new HashSet<string>(); // key: "cat|colIndex"
                        for (int row = 2; row <= rowCount; row++)
                        {
                            string cat = worksheet.Cell(row, 2).GetString();
                            if (string.IsNullOrWhiteSpace(cat)) continue;

                            for (int i = 0; i < headerParams.Count; i++)
                            {
                                if (headerParams[i] == null) continue;
                                string key = cat + "|" + i;
                                if (nonEmpty.Contains(key)) continue;
                                if (!string.IsNullOrEmpty(worksheet.Cell(row, i + 3).GetString()))
                                    nonEmpty.Add(key);
                            }
                        }

                        foreach (var cat in sheetCategories)
                        {
                            for (int i = 0; i < headerParams.Count; i++)
                            {
                                var h = headerParams[i];
                                if (h == null) continue;
                                if (nonEmpty.Contains(cat + "|" + i))
                                    AddEntry(cat, h.RawName, h.IsType);
                            }
                        }
                    }
                }
            }

            settings.SelectedCategories = categoryOrder;
            return settings;
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
        }
    }
}

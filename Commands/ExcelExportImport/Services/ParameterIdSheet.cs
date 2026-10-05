using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using Tools28.Commands.ExcelExportImport.Models;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// 書き出した Excel の各列が「どのパラメータか」を識別番号（ParamId）で記録する隠しシート。
    ///
    /// 【目的】
    /// 同名パラメータ（例: 共有パラメータ「☆財産区分」が GUID 違いで2つ）は見出しを
    /// 「I-☆財産区分【共有#1】」「【共有#2】」のように番号で区別しているが、番号だけでは
    /// 同名パラメータを片方しか持たない要素でどちらを指すのか判断できず、値の取り違え・消去が起きた。
    /// 識別番号を残しておけば、インポート時に書き出し時と同じパラメータを正確に特定できる。
    ///
    /// 【形式】 1行目=見出し、2行目以降=「カテゴリ名 | 列見出し（マーカー除去後の表示名）| 識別番号」
    /// ⚠ シート名と見出しは検索キーなので多言語化しない（CLAUDE.md の文字列分類 C）。
    /// </summary>
    public static class ParameterIdSheet
    {
        /// <summary>隠しシート名（多言語化禁止）</summary>
        public const string SheetName = "Tools28_ParamIds";

        /// <summary>このシートがパラメータ識別番号の隠しシートか</summary>
        public static bool IsMetaSheet(IXLWorksheet worksheet)
            => worksheet != null && IsMetaSheetName(worksheet.Name);

        /// <summary>シート名がパラメータ識別番号の隠しシートか</summary>
        public static bool IsMetaSheetName(string sheetName)
            => string.Equals(sheetName, SheetName, StringComparison.OrdinalIgnoreCase);

        /// <summary>照合キー（カテゴリ名 + 列見出し）</summary>
        public static string Key(string categoryName, string header)
            => (categoryName ?? "") + "|" + (header ?? "");

        /// <summary>出力パラメータの識別番号を隠しシートに書き込む</summary>
        public static void Write(XLWorkbook workbook, IEnumerable<ParameterInfo> outputParameters)
        {
            var ws = workbook.Worksheets.Add(SheetName);
            ws.Cell(1, 1).Value = "Category";
            ws.Cell(1, 2).Value = "Header";
            ws.Cell(1, 3).Value = "ParamId";

            int row = 2;
            var written = new HashSet<string>();
            foreach (var p in outputParameters)
            {
                string key = Key(p.CategoryName, p.DisplayName);
                if (!written.Add(key)) continue;
                ws.Cell(row, 1).Value = p.CategoryName ?? "";
                ws.Cell(row, 2).Value = p.DisplayName;
                ws.Cell(row, 3).Value = p.ParamId;
                row++;
            }

            ws.Visibility = XLWorksheetVisibility.Hidden;
        }

        /// <summary>
        /// 隠しシートから「カテゴリ名|列見出し → 識別番号」を読み込む。
        /// シートが無い（旧バージョンで書き出した Excel 等）場合は空の辞書を返す。
        /// </summary>
        public static Dictionary<string, long> Read(XLWorkbook workbook)
        {
            var map = new Dictionary<string, long>();
            IXLWorksheet ws = null;
            foreach (var sheet in workbook.Worksheets)
            {
                if (IsMetaSheet(sheet)) { ws = sheet; break; }
            }
            if (ws == null) return map;

            var lastRow = ws.LastRowUsed();
            if (lastRow == null) return map;

            for (int row = 2; row <= lastRow.RowNumber(); row++)
            {
                string cat = ws.Cell(row, 1).GetString();
                string header = ws.Cell(row, 2).GetString();
                var idCell = ws.Cell(row, 3);
                long id;
                if (idCell.DataType == XLDataType.Number)
                    id = (long)idCell.GetDouble();
                else if (!long.TryParse(idCell.GetString(), out id))
                    continue;
                if (id == 0 || string.IsNullOrEmpty(header)) continue;
                map[Key(cat, header)] = id;
            }
            return map;
        }
    }
}

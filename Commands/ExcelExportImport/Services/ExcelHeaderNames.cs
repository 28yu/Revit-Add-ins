using System;
using System.Collections.Generic;
using Tools28.Localization;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// Excel の固定見出し（要素ID / カテゴリ）と統合シート名の、言語をまたいだ照合を一元管理する。
    ///
    /// 【方針】
    ///  - 書き出しは <b>現在のアドイン言語</b>で行う
    ///  - 読み戻しは <b>全言語の候補</b>と照合する
    /// これにより、ある言語で書き出した Excel を別言語環境で取り込んでも列を特定でき、
    /// 多言語化前（日本語固定）に書き出した Excel も引き続き読める。
    /// ParameterKindHelper の種別ラベル、ParameterHeaderMarker のマーカーと同じ考え方。
    ///
    /// ⚠ 候補一覧は Strings{JP,EN,CN}.cs の対応するキーと必ず一致させること。
    /// </summary>
    public static class ExcelHeaderNames
    {
        /// <summary>1列目の見出し（書き出し用・現在の言語）</summary>
        public static string ElementId => Loc.S("Export.Header.ElementId");

        /// <summary>2列目の見出し（書き出し用・現在の言語）</summary>
        public static string Category => Loc.S("Export.Header.Category");

        /// <summary>1シート統合モードのシート名（書き出し用・現在の言語）</summary>
        public static string MergedSheet => Loc.S("Export.SheetName.Merged");

        private static readonly HashSet<string> ElementIdCandidates =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "要素ID", "Element ID", "图元ID",
            };

        private static readonly HashSet<string> CategoryCandidates =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "カテゴリ", "Category", "类别",
            };

        private static readonly HashSet<string> MergedSheetCandidates =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "データ", "Data", "数据",
            };

        /// <summary>「要素ID」列の見出しか（全言語で照合）</summary>
        public static bool IsElementIdHeader(string text)
            => !string.IsNullOrEmpty(text) && ElementIdCandidates.Contains(text.Trim());

        /// <summary>「カテゴリ」列の見出しか（全言語で照合）</summary>
        public static bool IsCategoryHeader(string text)
            => !string.IsNullOrEmpty(text) && CategoryCandidates.Contains(text.Trim());

        /// <summary>1シート統合モードのシート名か（全言語で照合）</summary>
        public static bool IsMergedSheetName(string sheetName)
            => !string.IsNullOrEmpty(sheetName) && MergedSheetCandidates.Contains(sheetName.Trim());
    }
}

using System.Collections.Generic;

namespace Tools28.Commands.ExcelExportImport.Models
{
    /// <summary>
    /// エクスポート設定の保存/読込用モデル
    /// </summary>
    public class ExportSettings
    {
        /// <summary>選択されたカテゴリ名リスト</summary>
        public List<string> SelectedCategories { get; set; } = new List<string>();

        /// <summary>出力パラメータリスト（順序付き）</summary>
        public List<ExportParameterEntry> OutputParameters { get; set; } = new List<ExportParameterEntry>();
    }

    /// <summary>
    /// エクスポート対象パラメータの保存用エントリ
    /// </summary>
    public class ExportParameterEntry
    {
        /// <summary>パラメータ名（プレフィックスなし）</summary>
        public string RawName { get; set; }

        /// <summary>タイプパラメータの場合true</summary>
        public bool IsTypeParameter { get; set; }

        /// <summary>所属カテゴリ名</summary>
        public string CategoryName { get; set; }

        /// <summary>
        /// 表示名（I-/T-プレフィックス＋名前＋同名区別の接尾辞）。
        /// 同名パラメータ（例: エリアの「用途」×2）を復元時に正しく特定するための識別子。
        /// 旧設定ファイルには存在しない（null）場合があり、その際は RawName で復元する。
        /// </summary>
        public string DisplayName { get; set; }
    }
}

using Tools28.Commands.ExcelExportImport.Models;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// エクスポート時にヘッダー名へ付ける「編集可否」マーカーの定義と、
    /// インポート/設定読込時にそれを除去する処理を一元管理する。
    /// これにより「文字値を入れられない列」を書き出し時点で一目で分かるようにする。
    /// </summary>
    public static class ParameterHeaderMarker
    {
        /// <summary>読み取り専用（値を変更できない）</summary>
        public const string ReadOnly = "(*変更不可)";

        /// <summary>画像参照（文字値は設定不可。画像ピッカー専用）</summary>
        public const string Image = "(*画像参照/文字不可)";

        /// <summary>要素参照（レベル/材料/タイプ等。既存要素名が必要で自由文字は不可）</summary>
        public const string Reference = "(*要素参照/名称必須)";

        private static readonly string[] AllMarkers = { ReadOnly, Image, Reference };

        /// <summary>
        /// パラメータに対応するヘッダー用マーカーを返す（不要なら空文字）。
        /// 優先順位: 画像参照 &gt; 要素参照 &gt; 読取専用。
        /// 「タイプ」はタイプ名で変更でき運用上も頻用のためマーカーを付けない。
        /// </summary>
        public static string MarkerFor(ParameterInfo p)
        {
            if (p == null) return "";
            if (p.RawName == "タイプ") return "";      // タイプ名で変更可能（別処理）
            if (p.IsImage) return Image;
            if (p.IsElementReference) return Reference;
            if (p.IsReadOnly) return ReadOnly;
            return "";
        }

        /// <summary>このパラメータ列が「文字値を直接入れられない」列かどうか（色付け判定用）</summary>
        public static bool IsNonTextEditable(ParameterInfo p)
        {
            return !string.IsNullOrEmpty(MarkerFor(p));
        }

        /// <summary>ヘッダー名末尾のマーカーを除去して素のパラメータ表示名に戻す</summary>
        public static string Strip(string headerName)
        {
            if (string.IsNullOrEmpty(headerName)) return headerName;
            foreach (var m in AllMarkers)
            {
                if (headerName.EndsWith(m))
                    return headerName.Substring(0, headerName.Length - m.Length);
            }
            return headerName;
        }
    }
}

using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.ExcelExportImport.Models;
using Tools28.Localization;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// エクスポート時にヘッダー名へ付ける「編集可否」マーカーの定義と、
    /// インポート/設定読込時にそれを除去する処理を一元管理する。
    /// これにより「文字値を入れられない列」を書き出し時点で一目で分かるようにする。
    ///
    /// 【多言語の扱い】
    /// マーカーは <b>書き出し時は現在のアドイン言語</b>で付け、
    /// <b>読み戻し時は全言語のマーカーを候補にして除去</b>する。
    /// これにより、ある言語で書き出した Excel を別言語環境で取り込んでも
    /// パラメータ名の照合が壊れない（過去の日本語固定版で書き出した Excel も読める）。
    /// ParameterKindHelper の種別ラベルと同じ考え方。
    /// </summary>
    public static class ParameterHeaderMarker
    {
        /// <summary>読み取り専用（値を変更できない）</summary>
        public static string ReadOnly => Loc.S("Export.Marker.ReadOnly");

        /// <summary>画像参照（文字値は設定不可。画像ピッカー専用）</summary>
        public static string Image => Loc.S("Export.Marker.Image");

        /// <summary>要素参照（レベル/材料/タイプ等。既存要素名が必要で自由文字は不可）</summary>
        public static string Reference => Loc.S("Export.Marker.Reference");

        /// <summary>
        /// 除去対象のマーカー（全言語）。
        /// ⚠ Strings{JP,EN,CN}.cs の Export.Marker.* と必ず一致させること。
        ///    日本語のものは、多言語化前に書き出した Excel を読むためにも必要。
        /// </summary>
        private static readonly string[] AllMarkers =
        {
            // ReadOnly
            "(*変更不可)", "(*Read-only)", "(*不可修改)",
            // Image
            "(*画像参照/文字不可)", "(*Image ref/no text)", "(*图像引用/不可输入文字)",
            // Reference
            "(*要素参照/名称必須)", "(*Element ref/name required)", "(*图元引用/需要名称)",
        };

        /// <summary>
        /// パラメータに対応するヘッダー用マーカーを返す（不要なら空文字）。
        /// 優先順位: 画像参照 &gt; 要素参照 &gt; 読取専用。
        /// 「タイプ」はタイプ名で変更でき運用上も頻用のためマーカーを付けない。
        /// </summary>
        public static string MarkerFor(ParameterInfo p)
        {
            if (p == null) return "";
            if (IsTypeParameter(p)) return "";      // タイプ名で変更可能（別処理）
            if (p.IsImage) return Image;
            if (p.IsElementReference) return Reference;
            if (p.IsReadOnly) return ReadOnly;
            return "";
        }

        /// <summary>
        /// 「タイプ」パラメータ（ELEM_TYPE_PARAM）かどうか。
        /// ⚠ 名前で判定してはいけない。パラメータ名は Revit の言語に従うため、
        ///    英語版 Revit では "Type"、中国語版では "类型" になり、日本語literalとの
        ///    比較が外れて「タイプ」列が編集不可扱い（灰色）で書き出されてしまう。
        ///    組み込みパラメータ Id は言語に依存しないのでこちらで判定する。
        /// </summary>
        private static bool IsTypeParameter(ParameterInfo p)
        {
            if (p.ParamId == (long)BuiltInParameter.ELEM_TYPE_PARAM) return true;

            // ParamId を取得できなかった場合の保険（日本語環境のみ有効）
            return p.RawName == "タイプ";
        }

        /// <summary>このパラメータ列が「文字値を直接入れられない」列かどうか（色付け判定用）</summary>
        public static bool IsNonTextEditable(ParameterInfo p)
        {
            return !string.IsNullOrEmpty(MarkerFor(p));
        }

        /// <summary>
        /// ヘッダー名末尾のマーカーを除去して素のパラメータ表示名に戻す。
        /// どの言語で書き出された Excel でも除去できるよう、全言語のマーカーを候補にする。
        /// </summary>
        public static string Strip(string headerName)
        {
            if (string.IsNullOrEmpty(headerName)) return headerName;
            foreach (var m in AllMarkers)
            {
                if (!string.IsNullOrEmpty(m) && headerName.EndsWith(m))
                    return headerName.Substring(0, headerName.Length - m.Length);
            }
            return headerName;
        }
    }
}

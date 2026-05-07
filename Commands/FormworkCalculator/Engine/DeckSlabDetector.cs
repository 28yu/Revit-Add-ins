using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// デッキスラブ判定。床のタイプ名 (または要素名) に "DS" を含むものを
    /// デッキスラブとして型枠算出から除外する。
    /// </summary>
    internal static class DeckSlabDetector
    {
        /// <summary>
        /// 要素がデッキスラブと判定された場合 true を返し reason に判定理由を出力する。
        /// </summary>
        internal static bool IsDeckSlab(Element elem, Document doc, out string reason)
        {
            reason = string.Empty;
            if (elem == null) return false;
            if (elem.Category == null) return false;
            if (elem.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Floors) return false;

            string typeName = string.Empty;
            try
            {
                var t = doc.GetElement(elem.GetTypeId()) as ElementType;
                typeName = t?.Name ?? string.Empty;
            }
            catch { }

            string instName = string.Empty;
            try { instName = elem.Name ?? string.Empty; } catch { }

            if (ContainsDsToken(typeName))
            {
                reason = $"floor type name contains 'DS' (type='{typeName}')";
                return true;
            }
            if (ContainsDsToken(instName))
            {
                reason = $"floor name contains 'DS' (name='{instName}')";
                return true;
            }
            return false;
        }

        /// <summary>
        /// 半角 "DS" / 全角 "ＤＳ" を含めば true。
        /// （誤検出を抑えるため、大文字 "DS" のみを対象とし小文字 "ds" は除外する）
        /// </summary>
        private static bool ContainsDsToken(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            return s.Contains("DS") || s.Contains("ＤＳ");
        }
    }
}

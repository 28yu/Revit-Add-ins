using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamCommon
{
    /// <summary>
    /// 梁天端・梁下端色分けの「レベル差 → 色」の対応を管理する。
    ///
    /// 色はレベル差（"1FL-50" の "-50" の部分）だけで決め、どの階で実行しても同じ色にする。
    /// 対応はモデル内の塗潰し領域タイプ（例: "梁天端色_-50"）に保存し、一度決めた色は変えない。
    /// 新しいレベル差は、まだ使われていない色を順に割り当てる。
    /// 利用者がタイプの色を変えれば、次回以降の実行でもその色が使われる。
    /// </summary>
    public static class BeamLevelColors
    {
        private static readonly Regex OffsetRegex = new Regex(@"[+-]\d+$");

        // Revit の要素名（ビュー名・フィルタ名・タイプ名）に使えない文字
        private const string InvalidNameChars = "\\:{}[]|;<>?`~";

        /// <summary>
        /// 表示値からレベル差の部分を取り出す（例: "1FL-50" → "-50"、"2FL+0" → "+0"）
        /// </summary>
        public static string OffsetKey(string displayValue)
        {
            Match m = OffsetRegex.Match(displayValue ?? "");
            return m.Success ? m.Value : displayValue;
        }

        /// <summary>
        /// 表記する文字に、要素名に使えない文字が含まれていれば最初の1文字を返す（なければ null）
        /// </summary>
        public static char? FindInvalidNameChar(string text)
        {
            foreach (char c in text)
            {
                if (InvalidNameChars.IndexOf(c) >= 0)
                    return c;
            }
            return null;
        }

        /// <summary>
        /// 各表示値の色を決める。新しいレベル差の色は塗潰し領域タイプとして保存する（トランザクション内で呼ぶ）
        /// </summary>
        /// <param name="colorTypePrefix">保存用タイプ名の接頭辞（例: "梁天端色_"）。変更禁止の検索キー</param>
        /// <param name="displayValues">今回の表示値（例: "1FL-50"）</param>
        /// <param name="palette">割り当てに使う色の並び</param>
        public static Dictionary<string, Color> Resolve(
            Document doc,
            string colorTypePrefix,
            IEnumerable<string> displayValues,
            IList<Color> palette)
        {
            // 既存の対応（レベル差 → 色）を読む
            var keyColors = new Dictionary<string, Color>();
            var regionTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .ToList();
            foreach (var t in regionTypes.Where(t => t.Name.StartsWith(colorTypePrefix)))
            {
                string key = t.Name.Substring(colorTypePrefix.Length);
                if (!keyColors.ContainsKey(key))
                    keyColors[key] = t.ForegroundPatternColor;
            }

            ElementId solidFillId = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill)?.Id;
            FilledRegionType baseType = regionTypes.FirstOrDefault();

            var newKeys = displayValues
                .Select(OffsetKey)
                .Distinct()
                .Where(k => !keyColors.ContainsKey(k))
                .OrderBy(ParseOffset)
                .ToList();

            foreach (string key in newKeys)
            {
                Color color = PickUnusedColor(palette, keyColors.Values);
                keyColors[key] = color;
                SaveColor(doc, baseType, solidFillId, colorTypePrefix + key, color);
            }

            var result = new Dictionary<string, Color>();
            foreach (string displayValue in displayValues.Distinct())
                result[displayValue] = keyColors[OffsetKey(displayValue)];
            return result;
        }

        /// <summary>
        /// 凡例の矩形に使う塗潰し領域タイプ名（色の保存先と同じタイプを使う）
        /// </summary>
        public static string ColorTypeName(string colorTypePrefix, string displayValue)
        {
            return colorTypePrefix + OffsetKey(displayValue);
        }

        private static double ParseOffset(string key)
        {
            return double.TryParse(key, out double v) ? v : 0;
        }

        private static Color PickUnusedColor(IList<Color> palette, IEnumerable<Color> used)
        {
            var usedList = used.ToList();
            foreach (Color c in palette)
            {
                if (!usedList.Any(u => u.IsValid && u.Red == c.Red && u.Green == c.Green && u.Blue == c.Blue))
                    return c;
            }
            // 全色使用済み: 使用数に応じて先頭から順に再利用する
            return palette[usedList.Count % palette.Count];
        }

        private static void SaveColor(Document doc, FilledRegionType baseType,
            ElementId solidFillId, string typeName, Color color)
        {
            if (baseType == null || solidFillId == null)
                return;
            try
            {
                var newType = baseType.Duplicate(typeName) as FilledRegionType;
                if (newType == null)
                    return;
                newType.ForegroundPatternColor = color;
                newType.ForegroundPatternId = solidFillId;
                newType.BackgroundPatternId = ElementId.InvalidElementId;
                newType.IsMasking = false;
            }
            catch (Exception)
            {
                // 保存できなくても今回の色分けは続ける（次回は別の色になる場合がある）
            }
        }
    }
}

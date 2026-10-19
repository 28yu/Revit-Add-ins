using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.BeamCommon;

namespace Tools28.Commands.BeamUnderLevel
{
    /// <summary>
    /// 梁下端色分け用フィルタ作成・色分け管理
    /// </summary>
    public static class FilterManager
    {
        // フィルタ名のプレフィックス
        private const string FilterPrefix = "梁下_";
        private const string ErrorFilterName = "梁下_エラー";

        // レベル差ごとの色の保存先（塗潰し領域タイプ名の接頭辞）。検索キーなので変更・多言語化禁止
        internal const string ColorTypePrefix = "梁下端色_";

        /// <summary>
        /// 色分けフィルタを作成しビューに適用
        /// </summary>
        public static void CreateFiltersAndColorize(
            Document doc,
            View activeView,
            Dictionary<string, int> levelGroups,
            bool overwriteExisting,
            int errorCount = 0)
        {
            List<ElementId> categories = new List<ElementId>
            {
                new ElementId(BuiltInCategory.OST_StructuralFraming)
            };

            // ベタ塗りパターンを取得
            ElementId solidFillPatternId = GetSolidFillPatternId(doc);

            // レベル値でソート（数値部分で）
            var sortedLevels = levelGroups
                .OrderBy(kv => ExtractNumericValue(kv.Key))
                .ToList();

            // レベル差ごとに固定した色（各階で同じレベル差は同じ色）
            Dictionary<string, Color> colors = ResolveColors(doc, levelGroups.Keys);

            // 既存フィルタを削除（上書きモード）
            if (overwriteExisting)
            {
                RemoveExistingFilters(doc, activeView);
            }

            // 各レベル値のフィルタを作成
            foreach (var entry in sortedLevels)
            {
                string displayValue = entry.Key;
                string filterName = FilterPrefix + displayValue;

                try
                {
                    // 同名フィルタは同じレベルの別のビュー（意匠の平面図と構造伏図など）でも
                    // 使われているため削除しない。既存があれば再利用してこのビューに適用する
                    ParameterFilterElement filterElement = FindFilterByName(doc, filterName);
                    if (filterElement == null)
                    {
                        ElementId paramId = GetSharedParameterIdByName(doc,
                            ParameterManager.ParamDisplay);
                        if (paramId == null)
                            continue;

#if REVIT2026
                        FilterRule rule = ParameterFilterRuleFactory.CreateEqualsRule(
                            paramId, displayValue);
#else
                        FilterRule rule = ParameterFilterRuleFactory.CreateEqualsRule(
                            paramId, displayValue, false);
#endif

                        ElementParameterFilter paramFilter =
                            new ElementParameterFilter(rule);

                        filterElement = ParameterFilterElement.Create(
                            doc, filterName, categories, paramFilter);
                    }

                    if (!activeView.GetFilters().Contains(filterElement.Id))
                    {
                        activeView.AddFilter(filterElement.Id);
                    }
                    else if (!overwriteExisting)
                    {
                        // 上書きしない場合は、このビューの既存の色をそのまま残す
                        continue;
                    }

                    Color color = colors[displayValue];
                    OverrideGraphicSettings overrides = CreateColorOverrides(
                        color, solidFillPatternId);

                    activeView.SetFilterOverrides(filterElement.Id, overrides);
                }
                catch (Exception)
                {
                    // フィルタ作成失敗はスキップ
                }
            }

            // エラー梁用フィルタ（エラーがある場合のみ）
            if (errorCount > 0)
            {
                CreateErrorFilter(doc, activeView, categories, overwriteExisting);
            }
        }

        /// <summary>
        /// エラー梁用フィルタを作成（赤色表示）
        /// </summary>
        private static void CreateErrorFilter(
            Document doc, View activeView,
            List<ElementId> categories, bool overwriteExisting)
        {
            try
            {
                // 他のビューでも使われているため削除せず、既存があれば再利用する
                ParameterFilterElement errorFilterElement = FindFilterByName(doc, ErrorFilterName);
                if (errorFilterElement == null)
                {
                    ElementId errorParamId = GetSharedParameterIdByName(doc,
                        ParameterManager.ParamError);
                    if (errorParamId == null)
                        return;

#if REVIT2026
                    FilterRule errorRule = ParameterFilterRuleFactory.CreateNotEqualsRule(
                        errorParamId, "");
#else
                    FilterRule errorRule = ParameterFilterRuleFactory.CreateNotEqualsRule(
                        errorParamId, "", false);
#endif

                    ElementParameterFilter errorFilter =
                        new ElementParameterFilter(errorRule);

                    errorFilterElement = ParameterFilterElement.Create(
                        doc, ErrorFilterName, categories, errorFilter);
                }

                if (!activeView.GetFilters().Contains(errorFilterElement.Id))
                    activeView.AddFilter(errorFilterElement.Id);
                else if (!overwriteExisting)
                    return;

                Color errorColor = new Color(255, 100, 100);
                ElementId solidFillId = GetSolidFillPatternId(doc);
                OverrideGraphicSettings errorOverrides = CreateColorOverrides(
                    errorColor, solidFillId);

                activeView.SetFilterOverrides(errorFilterElement.Id, errorOverrides);
            }
            catch (Exception)
            {
                // エラーフィルタ作成失敗はスキップ
            }
        }

        /// <summary>
        /// 既存の梁下フィルタをビューから削除
        /// </summary>
        private static void RemoveExistingFilters(Document doc, View view)
        {
            var filterIds = view.GetFilters();
            foreach (ElementId filterId in filterIds)
            {
                var filter = doc.GetElement(filterId) as ParameterFilterElement;
                if (filter != null && filter.Name.StartsWith(FilterPrefix))
                {
                    view.RemoveFilter(filterId);
                }
            }
        }

        /// <summary>
        /// 名前でフィルタを検索
        /// </summary>
        private static ParameterFilterElement FindFilterByName(Document doc, string filterName)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>()
                .FirstOrDefault(f => f.Name == filterName);
        }

        /// <summary>
        /// 共有パラメータのElementIdを取得
        /// </summary>
        private static ElementId GetSharedParameterIdByName(Document doc, string paramName)
        {
            // 構造フレームカテゴリの要素からパラメータIDを取得
            var collector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .WhereElementIsNotElementType();

            foreach (Element elem in collector)
            {
                Parameter param = elem.LookupParameter(paramName);
                if (param != null)
                    return param.Id;
            }

            return null;
        }

        /// <summary>
        /// 色によるグラフィック上書き設定を生成
        /// 投影サーフェス、断面パターン、投影線すべてに色を適用
        /// </summary>
        private static OverrideGraphicSettings CreateColorOverrides(
            Color color, ElementId solidFillPatternId)
        {
            var overrides = new OverrideGraphicSettings();

            if (solidFillPatternId != null)
            {
                // 投影サーフェス（前景）のみ
                overrides.SetSurfaceForegroundPatternId(solidFillPatternId);
                overrides.SetSurfaceForegroundPatternColor(color);
            }

            return overrides;
        }

        /// <summary>
        /// ベタ塗り（Solid Fill）パターンのElementIdを取得
        /// </summary>
        private static ElementId GetSolidFillPatternId(Document doc)
        {
            var fillPattern = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);

            return fillPattern?.Id;
        }

        /// <summary>
        /// 表示値ごとの色（レベル差で固定。凡例と同じ色を使う）
        /// </summary>
        internal static Dictionary<string, Color> ResolveColors(Document doc, IEnumerable<string> displayValues)
        {
            return BeamLevelColors.Resolve(doc, ColorTypePrefix, displayValues.ToList(), GenerateColors(20));
        }

        /// <summary>
        /// 表示値から数値部分を抽出（ソート用）
        /// 例: "2FL+2800" → 2800, "1FL-600" → -600
        /// </summary>
        internal static double ExtractNumericValue(string displayValue)
        {
            int plusIndex = displayValue.LastIndexOf('+');
            int minusIndex = displayValue.LastIndexOf('-');

            int signIndex = Math.Max(plusIndex, minusIndex);
            if (signIndex >= 0)
            {
                string numStr = displayValue.Substring(signIndex);
                if (double.TryParse(numStr, out double val))
                    return val;
            }

            return 0;
        }

        /// <summary>
        /// 色パレット生成（明るく見やすい色合い、暗い色は使用しない）
        /// </summary>
        internal static List<Color> GenerateColors(int colorCount)
        {
            var baseColors = new[]
            {
                new { R = 144, G = 175, B = 197 },  // mist blue
                new { R = 165, G = 196, B = 152 },  // sage green
                new { R = 210, G = 165, B = 120 },  // warm sand
                new { R = 180, G = 150, B = 180 },  // soft mauve
                new { R = 120, G = 180, B = 190 },  // teal
                new { R = 200, G = 160, B = 150 },  // rose beige
                new { R = 160, G = 185, B = 130 },  // leaf green
                new { R = 185, G = 170, B = 140 },  // light taupe
                new { R = 150, G = 170, B = 200 },  // periwinkle
                new { R = 200, G = 185, B = 150 },  // wheat
            };

            var colors = new List<Color>();

            for (int i = 0; i < colorCount; i++)
            {
                var baseColor = baseColors[i % baseColors.Length];
                int r = baseColor.R;
                int g = baseColor.G;
                int b = baseColor.B;

                // 10色以上の場合は明度を調整
                if (i >= baseColors.Length)
                {
                    float brightness = 0.8f + (float)(i % baseColors.Length) /
                        baseColors.Length * 0.4f;
                    r = Math.Min((int)(r * brightness), 255);
                    g = Math.Min((int)(g * brightness), 255);
                    b = Math.Min((int)(b * brightness), 255);
                }

                colors.Add(new Color((byte)r, (byte)g, (byte)b));
            }

            return colors;
        }
    }
}

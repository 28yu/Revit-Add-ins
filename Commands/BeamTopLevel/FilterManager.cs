using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamTopLevel
{
    /// <summary>
    /// 梁天端色分け用フィルタ作成・色分け管理
    /// </summary>
    public static class FilterManager
    {
        private const string FilterPrefix = "梁天端_";
        private const string ErrorFilterName = "梁天端_エラー";

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

            ElementId solidFillPatternId = GetSolidFillPatternId(doc);

            var sortedLevels = levelGroups
                .OrderBy(kv => ExtractNumericValue(kv.Key))
                .ToList();

            List<Color> colors = GenerateColors(sortedLevels.Count);

            if (overwriteExisting)
            {
                RemoveExistingFilters(doc, activeView);
            }

            int colorIndex = 0;
            foreach (var entry in sortedLevels)
            {
                string displayValue = entry.Key;
                string filterName = FilterPrefix + displayValue;

                try
                {
                    ParameterFilterElement existingFilter = FindFilterByName(doc, filterName);
                    if (existingFilter != null)
                    {
                        if (overwriteExisting)
                        {
                            if (activeView.GetFilters().Contains(existingFilter.Id))
                                activeView.RemoveFilter(existingFilter.Id);
                            doc.Delete(existingFilter.Id);
                        }
                        else
                        {
                            colorIndex++;
                            continue;
                        }
                    }

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

                    ParameterFilterElement filterElement =
                        ParameterFilterElement.Create(doc, filterName, categories, paramFilter);

                    activeView.AddFilter(filterElement.Id);

                    Color color = colors[colorIndex];
                    OverrideGraphicSettings overrides = CreateColorOverrides(
                        color, solidFillPatternId);

                    activeView.SetFilterOverrides(filterElement.Id, overrides);
                }
                catch (Exception)
                {
                    // フィルタ作成失敗はスキップ
                }

                colorIndex++;
            }

            if (errorCount > 0)
            {
                CreateErrorFilter(doc, activeView, categories, overwriteExisting);
            }
        }

        private static void CreateErrorFilter(
            Document doc, View activeView,
            List<ElementId> categories, bool overwriteExisting)
        {
            try
            {
                ParameterFilterElement existingError = FindFilterByName(doc, ErrorFilterName);
                if (existingError != null)
                {
                    if (overwriteExisting)
                    {
                        if (activeView.GetFilters().Contains(existingError.Id))
                            activeView.RemoveFilter(existingError.Id);
                        doc.Delete(existingError.Id);
                    }
                    else
                    {
                        return;
                    }
                }

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

                ParameterFilterElement errorFilterElement =
                    ParameterFilterElement.Create(doc, ErrorFilterName, categories, errorFilter);

                activeView.AddFilter(errorFilterElement.Id);

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

        private static ParameterFilterElement FindFilterByName(Document doc, string filterName)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>()
                .FirstOrDefault(f => f.Name == filterName);
        }

        private static ElementId GetSharedParameterIdByName(Document doc, string paramName)
        {
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

        private static OverrideGraphicSettings CreateColorOverrides(
            Color color, ElementId solidFillPatternId)
        {
            var overrides = new OverrideGraphicSettings();

            if (solidFillPatternId != null)
            {
                overrides.SetSurfaceForegroundPatternId(solidFillPatternId);
                overrides.SetSurfaceForegroundPatternColor(color);
            }

            return overrides;
        }

        private static ElementId GetSolidFillPatternId(Document doc)
        {
            var fillPattern = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);

            return fillPattern?.Id;
        }

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

using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 型枠 DirectShape に View Filter で色付けする。
    /// Per-element override の代わりに Filter を使うことで、ユーザーが
    /// 実行後にフィルタを編集して色や可視性を調整できるようにする。
    ///
    /// Filter Rule:  28Tools_Formwork_区分 == <groupKey>
    /// 対象カテゴリ: OST_GenericModel
    /// </summary>
    internal static class FormworkFilterManager
    {
        internal const string FilterPrefix = "型枠_";

        internal static void ApplyColorFilters(
            Document doc,
            View3D view,
            Dictionary<string, (byte R, byte G, byte B)> keyColors)
        {
            var catIds = new List<ElementId>
            {
                new ElementId(BuiltInCategory.OST_GenericModel),
            };

            ElementId solidFillId = GetDraftingSolidFillPatternId(doc);

            // 区分パラメータの ElementId を取得（DirectShape から）
            ElementId paramElemId = GetSharedParameterIdFromDirectShape(
                doc, FormworkParameterManager.ParamGroupKey);
            if (paramElemId == null) return;

            // 既存の型枠_* フィルタをビューから外して再作成
            RemoveExistingFiltersFromView(doc, view);

            foreach (var kv in keyColors)
            {
                string key = kv.Key;
                var (r, g, b) = kv.Value;

                string filterName = FilterPrefix + ReplaceInvalidChars(key);

                ParameterFilterElement filter = FindFilterByName(doc, filterName);
                if (filter != null)
                {
                    try { doc.Delete(filter.Id); } catch { }
                    filter = null;
                }

                try
                {
#if REVIT2026
                    FilterRule rule = ParameterFilterRuleFactory.CreateEqualsRule(paramElemId, key);
#else
                    FilterRule rule = ParameterFilterRuleFactory.CreateEqualsRule(paramElemId, key, false);
#endif
                    var elemFilter = new ElementParameterFilter(rule);
                    filter = ParameterFilterElement.Create(doc, filterName, catIds, elemFilter);
                }
                catch
                {
                    continue;
                }

                try
                {
                    view.AddFilter(filter.Id);

                    var ogs = new OverrideGraphicSettings();
                    var color = new Color(r, g, b);
                    ogs.SetProjectionLineColor(color);
                    ogs.SetSurfaceForegroundPatternColor(color);
                    if (solidFillId != null && solidFillId != ElementId.InvalidElementId)
                        ogs.SetSurfaceForegroundPatternId(solidFillId);
                    ogs.SetSurfaceForegroundPatternVisible(true);
                    ogs.SetSurfaceBackgroundPatternVisible(false);

                    view.SetFilterOverrides(filter.Id, ogs);
                    view.SetFilterVisibility(filter.Id, true);
                }
                catch { }
            }
        }

        private static void RemoveExistingFiltersFromView(Document doc, View view)
        {
            try
            {
                foreach (var fid in view.GetFilters().ToList())
                {
                    var f = doc.GetElement(fid) as ParameterFilterElement;
                    if (f != null && f.Name.StartsWith(FilterPrefix))
                    {
                        try { view.RemoveFilter(fid); } catch { }
                    }
                }
            }
            catch { }
        }

        private static ParameterFilterElement FindFilterByName(Document doc, string name)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>()
                .FirstOrDefault(f => f.Name == name);
        }

        /// <summary>
        /// OST_GenericModel カテゴリの任意の要素から、指定名の共有パラメータの
        /// ElementId を取得する。Filter Rule 作成に必要。
        /// </summary>
        private static ElementId GetSharedParameterIdFromDirectShape(Document doc, string paramName)
        {
            var collector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .WhereElementIsNotElementType();

            foreach (Element e in collector)
            {
                var p = e.LookupParameter(paramName);
                if (p != null) return p.Id;
            }
            return null;
        }

        private static ElementId GetDraftingSolidFillPatternId(Document doc)
        {
            var fps = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>();
            foreach (var fp in fps)
            {
                var p = fp.GetFillPattern();
                if (p == null) continue;
                if (p.IsSolidFill && p.Target == FillPatternTarget.Drafting)
                    return fp.Id;
            }
            foreach (var fp in fps)
            {
                var p = fp.GetFillPattern();
                if (p == null) continue;
                if (p.IsSolidFill) return fp.Id;
            }
            return ElementId.InvalidElementId;
        }

        private static string ReplaceInvalidChars(string s)
        {
            if (string.IsNullOrEmpty(s)) return "未設定";
            var invalid = new[] { '\\', '/', ':', '{', '}', '[', ']', '|', ';', '<', '>', '?', '`', '~' };
            var chars = s.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
                if (invalid.Contains(chars[i])) chars[i] = '_';
            return new string(chars);
        }
    }
}

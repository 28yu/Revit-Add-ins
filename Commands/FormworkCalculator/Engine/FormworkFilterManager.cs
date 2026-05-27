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
                new ElementId(FormworkParameterManager.FormworkCategory),
            };

            ElementId solidFillId = GetDraftingSolidFillPatternId(doc);
            Engine.FormworkDebugLog.Log(
                $"  [Filter] solidFillPatternId={(solidFillId != null && solidFillId != ElementId.InvalidElementId ? solidFillId.IntValue().ToString() : "INVALID")}");

            // 区分パラメータの ElementId を取得（DirectShape から）
            ElementId paramElemId = GetSharedParameterIdFromDirectShape(
                doc, FormworkParameterManager.ParamGroupKey);
            if (paramElemId == null)
            {
                Engine.FormworkDebugLog.Log("  [Filter] paramElemId is null - filters cannot be created");
                return;
            }
            Engine.FormworkDebugLog.Log(
                $"  [Filter] paramElemId={paramElemId.IntValue()} (28Tools_Formwork_区分)");

            // 既存の型枠_* フィルタをビューから外して再作成
            RemoveExistingFiltersFromView(doc, view);

            foreach (var kv in keyColors)
            {
                string key = kv.Key;
                var (r, g, b) = kv.Value;

                string filterName = FilterPrefix + ReplaceInvalidChars(key);

                ParameterFilterElement filter = FindFilterByName(doc, filterName);
                // 既存フィルタが旧カテゴリ (OST_GenericModel) を対象としていたら作り直す:
                // 旧バージョンの型枠 DS は OST_GenericModel、新版は FormworkCategory (NurseCallDevices)
                // なのでカテゴリが合っていないと新規 DS にマッチしない。
                if (filter != null)
                {
                    var fCats = filter.GetCategories();
                    bool hasNewCat = false;
                    if (fCats != null)
                    {
                        int newCatInt = (int)FormworkParameterManager.FormworkCategory;
                        foreach (var c in fCats)
                            if (c.IntValue() == newCatInt) { hasNewCat = true; break; }
                    }
                    if (!hasNewCat)
                    {
                        try { doc.Delete(filter.Id); } catch { }
                        filter = null;
                    }
                }
                // 既存フィルタを再利用 (削除すると他の分析ビューの参照が切れるため。
                // ただしカテゴリ不一致の場合は上で削除済み)
                if (filter == null)
                {
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
                    ogs.SetSurfaceTransparency(0); // formwork は完全不透明で表示

                    view.SetFilterOverrides(filter.Id, ogs);

                    // 除外フィルタは既定で非表示（ユーザーが手動で ON にすると確認可能）。
                    bool defaultVisible = key != FormworkParameterManager.ExcludedGroupKey;
                    view.SetFilterVisibility(filter.Id, defaultVisible);

                    Engine.FormworkDebugLog.Log(
                        $"  [Filter] '{filterName}' rule=区分=='{key}' color=({r},{g},{b}) " +
                        $"visible={defaultVisible}");
                }
                catch (System.Exception ex)
                {
                    Engine.FormworkDebugLog.Log($"  [Filter] '{filterName}' apply EX: {ex.Message}");
                }
            }
            Engine.FormworkDebugLog.Log(
                $"  [Filter] applied {keyColors.Count} filters to view '{view.Name}'");
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
        /// 型枠 DirectShape のカテゴリの任意の要素から、指定名の共有パラメータの
        /// ElementId を取得する。Filter Rule 作成に必要。
        /// 新カテゴリで見つからなければ旧カテゴリ (OST_GenericModel) も探す (移行期対応)。
        /// </summary>
        private static ElementId GetSharedParameterIdFromDirectShape(Document doc, string paramName)
        {
            foreach (var bic in new[] { FormworkParameterManager.FormworkCategory,
                                        FormworkParameterManager.LegacyFormworkCategory })
            {
                var collector = new FilteredElementCollector(doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType();

                foreach (Element e in collector)
                {
                    var p = e.LookupParameter(paramName);
                    if (p != null) return p.Id;
                }
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

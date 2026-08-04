using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FilterManagement.Models;

namespace Tools28.Commands.FilterManagement.Services
{
    /// <summary>
    /// プロジェクト内の全フィルタ（ルールベース／選択）を列挙し、
    /// 各フィルタがどのビュー／ビューテンプレートで使用されているかを判定するサービス。
    ///
    /// 判定方針:
    ///   - フィルタ本体は FilterElement（ParameterFilterElement + SelectionFilterElement）を列挙。
    ///   - 使用状況は全 View（通常ビュー＋ビューテンプレート）を走査し View.GetFilters() を集計。
    ///     ビュー数は要素数より遥かに少ないため同期処理で十分軽量。
    /// </summary>
    public class FilterScanner
    {
        /// <summary>
        /// 全フィルタを列挙し、使用ビュー・対象カテゴリを付与して返す。
        /// </summary>
        public List<FilterRow> EnumerateFilters(Document doc)
        {
            // --- 使用状況マップ（フィルタId -> 使用ビュー一覧）を先に構築 ---
            var usage = BuildUsageMap(doc);

            var rows = new List<FilterRow>();

            // ルールベース（ParameterFilterElement）と選択（SelectionFilterElement）を個別に収集。
            // 抽象基底 FilterElement を OfClass に渡すのは避け、具象型で確実に列挙する。
            var filters = new List<FilterElement>();
            CollectFilters(doc, typeof(ParameterFilterElement), filters);
            CollectFilters(doc, typeof(SelectionFilterElement), filters);

            foreach (var fe in filters)
            {
                if (fe == null) continue;

                var row = new FilterRow
                {
                    Id = fe.Id,
                    Name = fe.Name,
                    OriginalName = fe.Name,
                    Kind = (fe is SelectionFilterElement) ? FilterKind.Selection : FilterKind.Rule,
                    CategoriesText = BuildCategoriesText(doc, fe)
                };

                if (usage.TryGetValue(fe.Id, out var views))
                {
                    // ビュー名でソートして安定表示
                    row.UsedViews = views
                        .OrderByDescending(v => v.IsTemplate)
                        .ThenBy(v => v.ViewName, StringComparer.CurrentCultureIgnoreCase)
                        .ToList();
                }

                rows.Add(row);
            }

            // 未使用を上に、その中は名前昇順
            return rows
                .OrderBy(r => r.IsUsed)
                .ThenBy(r => r.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private static void CollectFilters(Document doc, Type t, List<FilterElement> bucket)
        {
            try
            {
                foreach (var e in new FilteredElementCollector(doc).OfClass(t).ToElements())
                {
                    if (e is FilterElement fe) bucket.Add(fe);
                }
            }
            catch { }
        }

        /// <summary>
        /// 全ビュー（通常＋テンプレート）を走査して フィルタId -> 使用ビュー一覧 を作る。
        /// </summary>
        private static Dictionary<ElementId, List<ViewUsage>> BuildUsageMap(Document doc)
        {
            var map = new Dictionary<ElementId, List<ViewUsage>>();

            IEnumerable<View> views;
            try
            {
                views = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>();
            }
            catch
            {
                return map;
            }

            foreach (var view in views)
            {
                if (view == null) continue;

                ICollection<ElementId> ids;
                try
                {
                    // グラフィックス上書きに対応しないビュー（集計表等）は例外になるため握りつぶす
                    if (!view.AreGraphicsOverridesAllowed()) continue;
                    ids = view.GetFilters();
                }
                catch
                {
                    continue;
                }
                if (ids == null || ids.Count == 0) continue;

                string vname;
                try { vname = view.Name; } catch { vname = "(?)"; }
                bool isTemplate = view.IsTemplate;

                foreach (var fid in ids)
                {
                    if (fid == null || fid == ElementId.InvalidElementId) continue;
                    if (!map.TryGetValue(fid, out var list))
                    {
                        list = new List<ViewUsage>();
                        map[fid] = list;
                    }
                    list.Add(new ViewUsage { ViewName = vname, IsTemplate = isTemplate });
                }
            }

            return map;
        }

        /// <summary>ルールベースフィルタの対象カテゴリ名を "," 区切りで返す。</summary>
        private static string BuildCategoriesText(Document doc, FilterElement fe)
        {
            if (!(fe is ParameterFilterElement pfe)) return "-";

            ICollection<ElementId> catIds;
            try { catIds = pfe.GetCategories(); }
            catch { return ""; }
            if (catIds == null || catIds.Count == 0) return "";

            var names = new List<string>();
            foreach (var cid in catIds)
            {
                try
                {
                    var c = Category.GetCategory(doc, cid);
                    if (c != null && !string.IsNullOrEmpty(c.Name)) names.Add(c.Name);
                }
                catch { }
            }
            names.Sort(StringComparer.CurrentCulture);
            return string.Join(", ", names);
        }
    }
}

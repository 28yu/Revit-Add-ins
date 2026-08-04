using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.ViewTemplateManagement.Models;
using Tools28.Localization;

namespace Tools28.Commands.ViewTemplateManagement.Services
{
    /// <summary>
    /// プロジェクト内の全ビューテンプレートを列挙し、
    /// 各テンプレートがどのビューに適用されているか（未使用か）を判定するサービス。
    ///
    /// 判定方針:
    ///   - テンプレートは IsTemplate == true の View。
    ///   - 使用状況は非テンプレートの View を走査し View.ViewTemplateId を集計。
    ///     ビュー数は要素数より遥かに少ないため同期処理で十分軽量。
    /// </summary>
    public class TemplateScanner
    {
        /// <summary>
        /// 全ビューテンプレートを列挙し、適用先ビュー・種別を付与して返す。
        /// </summary>
        public List<TemplateRow> EnumerateTemplates(Document doc)
        {
            var templates = new List<View>();
            var usage = new Dictionary<ElementId, List<string>>();

            IEnumerable<View> views;
            try
            {
                views = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>();
            }
            catch
            {
                return new List<TemplateRow>();
            }

            foreach (var view in views)
            {
                if (view == null) continue;

                if (view.IsTemplate)
                {
                    templates.Add(view);
                    continue;
                }

                // 非テンプレートビューが参照しているテンプレートを集計
                ElementId tid;
                try { tid = view.ViewTemplateId; }
                catch { continue; }
                if (tid == null || tid == ElementId.InvalidElementId) continue;

                string vname;
                try { vname = view.Name; } catch { vname = "(?)"; }

                if (!usage.TryGetValue(tid, out var list))
                {
                    list = new List<string>();
                    usage[tid] = list;
                }
                list.Add(vname);
            }

            var rows = new List<TemplateRow>();
            foreach (var t in templates)
            {
                string name;
                try { name = t.Name; } catch { name = "(?)"; }

                var row = new TemplateRow
                {
                    Id = t.Id,
                    Name = name,
                    OriginalName = name,
                    ViewTypeText = ViewTypeLabel(t)
                };

                if (usage.TryGetValue(t.Id, out var vlist))
                {
                    vlist.Sort(StringComparer.CurrentCultureIgnoreCase);
                    row.UsedViews = vlist;
                }

                rows.Add(row);
            }

            // 未使用を上に、その中は名前昇順
            return rows
                .OrderBy(r => r.IsUsed)
                .ThenBy(r => r.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        /// <summary>ビューテンプレートの種別（ビュータイプ）を localized 文字列で返す。</summary>
        private static string ViewTypeLabel(View v)
        {
            ViewType vt;
            try { vt = v.ViewType; }
            catch { return ""; }

            switch (vt)
            {
                case ViewType.FloorPlan: return Loc.S("ViewType.FloorPlan");
                case ViewType.CeilingPlan: return Loc.S("ViewType.CeilingPlan");
                case ViewType.EngineeringPlan: return Loc.S("ViewType.EngineeringPlan");
                case ViewType.AreaPlan: return Loc.S("ViewType.AreaPlan");
                case ViewType.Elevation: return Loc.S("ViewType.Elevation");
                case ViewType.Section: return Loc.S("ViewType.Section");
                case ViewType.Detail: return Loc.S("ViewType.Detail");
                case ViewType.ThreeD: return Loc.S("ViewType.ThreeD");
                case ViewType.DraftingView: return Loc.S("ViewType.DraftingView");
                case ViewType.Legend: return Loc.S("ViewType.Legend");
                case ViewType.Schedule: return Loc.S("ViewType.Schedule");
                case ViewType.Rendering: return Loc.S("ViewType.Rendering");
                case ViewType.Walkthrough: return Loc.S("ViewType.Walkthrough");
                default: return vt.ToString();
            }
        }
    }
}

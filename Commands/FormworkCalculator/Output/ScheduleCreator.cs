using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Engine;

namespace Tools28.Commands.FormworkCalculator.Output
{
    /// <summary>
    /// ViewSchedule を作成して DirectShape に格納した型枠情報を集計表示する。
    /// - 対象カテゴリ: OST_GenericModel
    /// - フィルタ: 28Tools_FormworkMarker == "28Tools_Formwork"
    /// - 列: 部位 / 区分 / 面積
    /// - 部位 → 区分 で階層グループ化 (IsItemized=false で集計行のみ表示)
    /// - 総合計を表示
    /// - 名前は "型枠数量集計" 固定（既存同名はコマンド側で削除）
    /// </summary>
    internal static class ScheduleCreator
    {
        internal const string ScheduleName = "型枠数量集計";

        internal static ElementId CreateSchedule(Document doc)
        {
            // 既存の同名集計表を削除
            DeleteScheduleByName(doc, ScheduleName);

            ViewSchedule schedule;
            try
            {
                schedule = ViewSchedule.CreateSchedule(
                    doc, new ElementId(BuiltInCategory.OST_GenericModel));
            }
            catch
            {
                return null;
            }
            if (schedule == null) return null;

            try { schedule.Name = ScheduleName; } catch { }

            var def = schedule.Definition;
            var schedulable = def.GetSchedulableFields();

            // 列: 部位 / 区分 / 面積
            var partField = AddFieldByName(doc, def, schedulable, FormworkParameterManager.ParamCategory);
            var groupField = AddFieldByName(doc, def, schedulable, FormworkParameterManager.ParamGroupKey);
            var areaField = AddFieldByName(doc, def, schedulable, FormworkParameterManager.ParamArea);

            // マーカー（非表示、フィルタ用）
            var markerField = AddFieldByName(doc, def, schedulable, FormworkParameterManager.ParamMarker);
            if (markerField != null)
            {
                try { markerField.IsHidden = true; } catch { }
                try
                {
                    var filter = new ScheduleFilter(
                        markerField.FieldId,
                        ScheduleFilterType.Equal,
                        FormworkParameterManager.MarkerValue);
                    def.AddFilter(filter);
                }
                catch { }
            }

            // 部位でソート・グループ化（ヘッダーあり・フッターあり）
            if (partField != null)
            {
                try
                {
                    var sortPart = new ScheduleSortGroupField(partField.FieldId)
                    {
                        ShowHeader = true,
                        ShowFooter = true,
                        ShowBlankLine = false,
                    };
                    def.AddSortGroupField(sortPart);
                }
                catch { }
            }

            // 区分でもソート（区分ごとに行を分ける）
            if (groupField != null)
            {
                try
                {
                    var sortGroup = new ScheduleSortGroupField(groupField.FieldId)
                    {
                        ShowHeader = false,
                        ShowFooter = false,
                        ShowBlankLine = false,
                    };
                    def.AddSortGroupField(sortGroup);
                }
                catch { }
            }

            // 集計表示: アイテム単位ではなくグループ単位で面積を合計表示
            try { def.IsItemized = false; } catch { }
            try { def.ShowGrandTotal = true; } catch { }

            return schedule.Id;
        }

        private static void DeleteScheduleByName(Document doc, string name)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(v => !v.IsTemplate && v.Name == name)
                .Select(v => v.Id)
                .ToList();
            foreach (var id in existing)
            {
                try { doc.Delete(id); } catch { }
            }
        }

        private static ScheduleField AddFieldByName(
            Document doc,
            ScheduleDefinition def,
            IList<SchedulableField> schedulable,
            string paramName)
        {
            var sf = schedulable.FirstOrDefault(f =>
            {
                try { return f.GetName(doc) == paramName; }
                catch { return false; }
            });
            if (sf == null) return null;
            try { return def.AddField(sf); } catch { return null; }
        }
    }
}

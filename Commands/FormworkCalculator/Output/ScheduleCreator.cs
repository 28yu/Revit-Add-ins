using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Engine;

namespace Tools28.Commands.FormworkCalculator.Output
{
    /// <summary>
    /// ViewSchedule を作成して DirectShape に格納した型枠情報を集計表示する。
    /// OST_GenericModel カテゴリで、識別用マーカーパラメータによりフィルタ。
    /// </summary>
    internal static class ScheduleCreator
    {
        internal static ElementId CreateSchedule(Document doc)
        {
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

            string viewName = $"型枠数量集計_{DateTime.Now:yyyyMMdd_HHmmss}";
            try { schedule.Name = viewName; } catch { }

            var def = schedule.Definition;
            var schedulable = def.GetSchedulableFields();

            // ① 部位フィールド
            var categoryField = AddFieldByName(doc, def, schedulable, FormworkParameterManager.ParamCategory);
            // ② 区分フィールド
            var groupField = AddFieldByName(doc, def, schedulable, FormworkParameterManager.ParamGroupKey);
            // ③ タイプ名（DirectShapeType 名）
            var typeNameField = AddFieldByBip(def, schedulable, BuiltInParameter.ALL_MODEL_TYPE_NAME);
            // ④ 面積
            var areaField = AddFieldByName(doc, def, schedulable, FormworkParameterManager.ParamArea);
            // ⑤ マーカー（フィルタ用、非表示）
            var markerField = AddFieldByName(doc, def, schedulable, FormworkParameterManager.ParamMarker);
            if (markerField != null)
            {
                try { markerField.IsHidden = true; } catch { }
            }

            // フィルタ: マーカー == "28Tools_Formwork"
            if (markerField != null)
            {
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

            // 部位でグループ化・合計
            if (categoryField != null)
            {
                try
                {
                    var sort = new ScheduleSortGroupField(categoryField.FieldId);
                    sort.ShowHeader = true;
                    sort.ShowFooter = true;
                    sort.ShowBlankLine = false;
                    def.AddSortGroupField(sort);
                }
                catch { }
            }

            // 総合計を有効化
            try { def.ShowGrandTotal = true; } catch { }
            try { def.IsItemized = true; } catch { }

            return schedule.Id;
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

        private static ScheduleField AddFieldByBip(
            ScheduleDefinition def,
            IList<SchedulableField> schedulable,
            BuiltInParameter bip)
        {
            var sf = schedulable.FirstOrDefault(f => f.ParameterId == new ElementId(bip));
            if (sf == null) return null;
            try { return def.AddField(sf); } catch { return null; }
        }
    }
}

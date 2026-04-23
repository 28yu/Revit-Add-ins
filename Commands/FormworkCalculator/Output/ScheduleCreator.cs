using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Engine;

namespace Tools28.Commands.FormworkCalculator.Output
{
    /// <summary>
    /// ViewSchedule を作成して DirectShape に格納した型枠情報を集計表示する。
    ///
    /// レイアウト:
    ///   - 対象カテゴリ: OST_GenericModel
    ///   - フィルタ: 28Tools_FormworkMarker == "28Tools_Formwork"
    ///   - 列: 部位 / 区分 / タイプ名 / 面積
    ///   - 部位でグループ化（ヘッダー・フッター表示 → 部位毎の合計行）
    ///   - 区分で2段目ソート
    ///   - 総合計表示
    ///   - IsItemized = true (各行を表示、フッターで合計)
    ///
    /// フィールド検索は SharedParameterElement の ElementId ベース + 名前ベースの
    /// ダブルチェックで取りこぼしを防ぐ。
    /// </summary>
    internal static class ScheduleCreator
    {
        internal const string ScheduleName = "型枠数量集計";

        internal static ElementId CreateSchedule(Document doc)
        {
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

            // 共有パラメータ Id マップを先に取得
            var paramIds = GetFormworkSharedParamIds(doc);

            // 列: 部位 / 区分 / タイプ名 / 面積
            var partField = AddField(doc, def, schedulable, paramIds,
                FormworkParameterManager.ParamCategory);
            var groupField = AddField(doc, def, schedulable, paramIds,
                FormworkParameterManager.ParamGroupKey);
            var typeNameField = AddFieldByBip(def, schedulable, BuiltInParameter.ALL_MODEL_TYPE_NAME);
            var areaField = AddField(doc, def, schedulable, paramIds,
                FormworkParameterManager.ParamArea);
            // マーカー（非表示、フィルタ用）
            var markerField = AddField(doc, def, schedulable, paramIds,
                FormworkParameterManager.ParamMarker);
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

            // 部位でグループ化（ヘッダー + フッター → 部位毎の合計行）
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

            // 区分で2段目ソート
            if (groupField != null)
            {
                try
                {
                    var sortGroup = new ScheduleSortGroupField(groupField.FieldId);
                    def.AddSortGroupField(sortGroup);
                }
                catch { }
            }

            // 総合計
            try { def.IsItemized = true; } catch { }
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

        /// <summary>
        /// 型枠用共有パラメータの名前 → ElementId マップを取得。
        /// SchedulableField.ParameterId で確実にフィールドを特定するため。
        /// </summary>
        private static Dictionary<string, ElementId> GetFormworkSharedParamIds(Document doc)
        {
            var map = new Dictionary<string, ElementId>();
            var names = new[]
            {
                FormworkParameterManager.ParamMarker,
                FormworkParameterManager.ParamCategory,
                FormworkParameterManager.ParamGroupKey,
                FormworkParameterManager.ParamArea,
            };

            try
            {
                var shared = new FilteredElementCollector(doc)
                    .OfClass(typeof(SharedParameterElement))
                    .Cast<SharedParameterElement>();
                foreach (var sp in shared)
                {
                    foreach (var n in names)
                    {
                        if (sp.Name == n && !map.ContainsKey(n))
                            map[n] = sp.Id;
                    }
                }
            }
            catch { }

            return map;
        }

        /// <summary>
        /// ParameterId 一致を優先し、無ければ名前一致で SchedulableField を追加。
        /// </summary>
        private static ScheduleField AddField(
            Document doc,
            ScheduleDefinition def,
            IList<SchedulableField> schedulable,
            Dictionary<string, ElementId> paramIds,
            string paramName)
        {
            SchedulableField sf = null;

            if (paramIds != null && paramIds.TryGetValue(paramName, out var pid) && pid != null)
            {
                sf = schedulable.FirstOrDefault(f => f.ParameterId == pid);
            }

            if (sf == null)
            {
                sf = schedulable.FirstOrDefault(f =>
                {
                    try { return f.GetName(doc) == paramName; }
                    catch { return false; }
                });
            }

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

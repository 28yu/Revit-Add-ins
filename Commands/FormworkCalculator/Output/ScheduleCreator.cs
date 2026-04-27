using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Engine;

namespace Tools28.Commands.FormworkCalculator.Output
{
    /// <summary>
    /// ViewSchedule を作成して DirectShape に格納した型枠情報を集計表示する。
    ///
    /// レイアウト（インスタンス内訳 + 階層集計）:
    ///   - IsItemized = true → 各 DirectShape を個別に表示
    ///   - 列: 件数 / レベル / 部位 / タイプ名 / 区分 / 面積(合計を計算)
    ///   - 階層グループ化 (ShowHeader + ShowFooter で各階層に合計行):
    ///       1段目: レベル  (参照レベル毎)
    ///       2段目: 部位    (カテゴリ毎)
    ///       3段目: タイプ名 (タイプ毎)
    ///   - 面積フィールドに HasTotals=true を設定 → 各グループおよび総合計で面積合計を表示
    ///   - 総合計行を表示
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
            var paramIds = GetFormworkSharedParamIds(doc);

            // 集計表の基本設定 (フィールド追加前に設定する必要がある場合がある)
            //   IsItemized=true: 各インスタンスを個別行で表示
            //   ShowGrandTotal=true: 末尾に総合計行
            try { def.IsItemized = true; } catch { }
            try { def.ShowGrandTotal = true; } catch { }

            // 列順: 件数 / レベル / 部位 / タイプ名 / 区分 / 面積
            var countField = AddCountField(def, schedulable);
            var levelField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamLevel);
            var partField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamCategory);
            var typeNameField = AddFieldByBip(def, schedulable, BuiltInParameter.ALL_MODEL_TYPE_NAME);
            var groupField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamGroupKey);
            var areaField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamArea);
            // 面積フィールドは「合計を計算」を有効化 (各グループ/総合計で合算表示)
            if (areaField != null)
            {
                try { areaField.HasTotals = true; } catch { }
            }

            // マーカー（非表示、フィルタ用）
            var markerField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamMarker);
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

            // 階層グループ化: レベル → 部位 → タイプ名（各階層にヘッダー・フッター）
            AddGroupField(def, levelField);
            AddGroupField(def, partField);
            AddGroupField(def, typeNameField);

            // 区分はソートのみ（見出し・フッタは表示しない）
            if (groupField != null)
            {
                try
                {
                    var sortGroup = new ScheduleSortGroupField(groupField.FieldId)
                    {
                        ShowHeader = false,
                        ShowFooter = false,
                        ShowBlankLine = false,
                        SortOrder = ScheduleSortOrder.Ascending,
                    };
                    def.AddSortGroupField(sortGroup);
                }
                catch { }
            }

            // SortGroupField 追加後に再度設定を確認 (Revit が上書きする場合の保険)
            // FieldId で再取得して設定する (古い参照では Revit 内部状態と同期していない可能性)
            try { def.IsItemized = true; } catch { }
            try { def.ShowGrandTotal = true; } catch { }
            if (areaField != null)
            {
                try
                {
                    var freshAreaField = def.GetField(areaField.FieldId);
                    if (freshAreaField != null) freshAreaField.HasTotals = true;
                }
                catch { }
            }

            return schedule.Id;
        }

        /// <summary>
        /// グループフィールドとして追加（ヘッダー + フッター付き）。
        /// </summary>
        private static void AddGroupField(ScheduleDefinition def, ScheduleField field)
        {
            if (field == null) return;
            try
            {
                var sort = new ScheduleSortGroupField(field.FieldId)
                {
                    ShowHeader = true,
                    ShowFooter = true,
                    ShowBlankLine = false,
                    SortOrder = ScheduleSortOrder.Ascending,
                };
                def.AddSortGroupField(sort);
            }
            catch { }
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

        private static Dictionary<string, ElementId> GetFormworkSharedParamIds(Document doc)
        {
            var map = new Dictionary<string, ElementId>();
            var names = new[]
            {
                FormworkParameterManager.ParamMarker,
                FormworkParameterManager.ParamCategory,
                FormworkParameterManager.ParamLevel,
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

        private static ScheduleField AddCountField(
            ScheduleDefinition def,
            IList<SchedulableField> schedulable)
        {
            var sf = schedulable.FirstOrDefault(f => f.FieldType == ScheduleFieldType.Count);
            if (sf == null) return null;
            try { return def.AddField(sf); } catch { return null; }
        }
    }
}

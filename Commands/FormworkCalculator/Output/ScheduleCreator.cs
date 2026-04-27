using System;
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

            FormworkDebugLog.Section("Schedule Creation");
            LogDef(def, "after-create");

            // 集計表の基本設定 (フィールド追加前に設定する必要がある場合がある)
            //   IsItemized=true: 各インスタンスを個別行で表示
            //   ShowGrandTotal=true: 末尾に総合計行
            try { def.IsItemized = true; } catch (Exception ex) { LogEx("IsItemized=true (before)", ex); }
            try { def.ShowGrandTotal = true; } catch (Exception ex) { LogEx("ShowGrandTotal=true (before)", ex); }
            LogDef(def, "after-set-itemized-before-fields");

            // 列順: 件数 / レベル / 部位 / タイプ名 / 区分 / 面積
            var countField = AddCountField(def, schedulable);
            var levelField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamLevel);
            var partField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamCategory);
            var typeNameField = AddFieldByBip(def, schedulable, BuiltInParameter.ALL_MODEL_TYPE_NAME);
            var groupField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamGroupKey);
            var areaField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamArea);
            LogDef(def, "after-add-fields");
            LogField(areaField, "areaField after-add");

            // 面積フィールドは「合計を計算」を有効化 (各グループ/総合計で合算表示)
            //   ScheduleField.HasTotals は Revit 2024 以降の API。
            //   Revit 2022/2023 にはこのプロパティが存在しないため、リフレクションで動的に設定する。
            //   存在しないバージョンではログに記録するのみ (ユーザーが手動でチェックを入れる必要がある)。
            if (areaField != null)
            {
                // Revit 2022 で「合計を計算」を有効化する API を診断するため、
                // 1度目のみ全プロパティ・メソッドを列挙してログに出力する
                if (FormworkDebugLog.Enabled)
                    DiagScheduleFieldMembersOnce(areaField);

                bool setOk = TrySetHasTotals(areaField, true);
                if (FormworkDebugLog.Enabled)
                    FormworkDebugLog.Log($"  [Sched] HasTotals reflection set: success={setOk}");

                // HasTotals 以外の totals 関連プロパティ/メソッドを試す
                if (!setOk) TryAlternativeTotalsApi(areaField);

                LogField(areaField, "areaField after-set-hastotals-1");
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
            LogDef(def, "after-group-fields");
            LogField(areaField, "areaField after-group");

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
            LogDef(def, "after-sort-group");

            // SortGroupField 追加後に再度設定を確認 (Revit が上書きする場合の保険)
            // FieldId で再取得して設定する (古い参照では Revit 内部状態と同期していない可能性)
            try { def.IsItemized = true; } catch (Exception ex) { LogEx("IsItemized=true (after)", ex); }
            try { def.ShowGrandTotal = true; } catch { }
            if (areaField != null)
            {
                try
                {
                    var freshAreaField = def.GetField(areaField.FieldId);
                    if (freshAreaField != null)
                    {
                        TrySetHasTotals(freshAreaField, true);
                        LogField(freshAreaField, "freshAreaField after-set-hastotals-2");
                    }
                }
                catch (Exception ex) { LogEx("freshAreaField.HasTotals=true", ex); }
            }
            LogDef(def, "FINAL");

            return schedule.Id;
        }

        private static void LogDef(ScheduleDefinition def, string stage)
        {
            if (!FormworkDebugLog.Enabled) return;
            bool isItemized = false;
            bool grandTotal = false;
            int fieldCount = 0;
            int sortCount = 0;
            try { isItemized = def.IsItemized; } catch { }
            try { grandTotal = def.ShowGrandTotal; } catch { }
            try { fieldCount = def.GetFieldCount(); } catch { }
            try { sortCount = def.GetSortGroupFieldCount(); } catch { }
            FormworkDebugLog.Log(
                $"  [Sched:{stage}] IsItemized={isItemized} ShowGrandTotal={grandTotal} " +
                $"fields={fieldCount} sortGroups={sortCount}");
        }

        private static void LogField(ScheduleField field, string label)
        {
            if (!FormworkDebugLog.Enabled || field == null) return;
            string hasTotalsStr = "n/a";
            string fieldType = "?";
            string colHeading = "?";
            bool isCalculated = false;
            bool isHidden = false;
            try
            {
                var prop = typeof(ScheduleField).GetProperty(
                    "HasTotals",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                if (prop != null && prop.CanRead)
                    hasTotalsStr = ((bool)prop.GetValue(field, null)).ToString();
            }
            catch { }
            try { fieldType = field.FieldType.ToString(); } catch { }
            try { colHeading = field.ColumnHeading; } catch { }
            try { isCalculated = field.IsCalculatedField; } catch { }
            try { isHidden = field.IsHidden; } catch { }
            FormworkDebugLog.Log(
                $"  [Sched:{label}] heading='{colHeading}' fieldType={fieldType} " +
                $"HasTotals={hasTotalsStr} IsCalculated={isCalculated} IsHidden={isHidden}");
        }

        private static void LogEx(string action, Exception ex)
        {
            if (!FormworkDebugLog.Enabled) return;
            FormworkDebugLog.Log($"  [Sched:EX] {action}: {ex.GetType().Name}: {ex.Message}");
        }

        /// <summary>
        /// ScheduleField.HasTotals プロパティが存在する場合 (Revit 2024+) のみ設定する。
        /// 存在しない Revit 2021-2023 ではコンパイル参照を避けるためリフレクションで設定。
        /// 戻り値: 設定に成功 (= プロパティが存在し書き込みできた) なら true。
        /// </summary>
        private static bool TrySetHasTotals(ScheduleField field, bool value)
        {
            if (field == null) return false;
            try
            {
                var prop = typeof(ScheduleField).GetProperty(
                    "HasTotals",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                if (prop == null || !prop.CanWrite) return false;
                prop.SetValue(field, value, null);
                return true;
            }
            catch (Exception ex)
            {
                LogEx("TrySetHasTotals", ex);
                return false;
            }
        }

        private static bool _diagDone;

        /// <summary>
        /// 診断用: ScheduleField の全プロパティ・メソッドを1度だけログに出力する。
        /// Revit 2022 で「合計を計算」を制御する API を特定するため。
        /// </summary>
        private static void DiagScheduleFieldMembersOnce(ScheduleField field)
        {
            if (_diagDone) return;
            _diagDone = true;
            if (field == null) return;
            FormworkDebugLog.Section("ScheduleField API Diagnostic");
            try
            {
                var t = typeof(ScheduleField);
                FormworkDebugLog.Log($"  Type: {t.AssemblyQualifiedName}");

                // プロパティ列挙
                FormworkDebugLog.Log("  -- Properties --");
                foreach (var p in t.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
                {
                    string val = "?";
                    try { val = p.CanRead ? (p.GetValue(field, null)?.ToString() ?? "null") : "(no-read)"; }
                    catch (Exception ex) { val = "EX:" + ex.GetType().Name; }
                    FormworkDebugLog.Log(
                        $"    {p.Name} : {p.PropertyType.Name} " +
                        $"[get={p.CanRead} set={p.CanWrite}] = {val}");
                }

                // メソッド列挙 (パラメータなしまたは少ないものに絞る)
                FormworkDebugLog.Log("  -- Methods (Set*/Get*/Is*/Has*/Calc*/Total*) --");
                foreach (var m in t.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
                {
                    string n = m.Name;
                    if (n.StartsWith("get_") || n.StartsWith("set_")) continue; // properties accessor
                    bool relevant = n.StartsWith("Set") || n.StartsWith("Get") ||
                                    n.StartsWith("Is") || n.StartsWith("Has") ||
                                    n.IndexOf("Total", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    n.IndexOf("Calc", StringComparison.OrdinalIgnoreCase) >= 0;
                    if (!relevant) continue;
                    var ps = m.GetParameters();
                    var sig = string.Join(",", ps.Select(p => p.ParameterType.Name).ToArray());
                    FormworkDebugLog.Log($"    {m.Name}({sig}) -> {m.ReturnType.Name}");
                }
            }
            catch (Exception ex)
            {
                LogEx("Diag", ex);
            }
        }

        /// <summary>
        /// HasTotals 以外の代替 API で「合計を計算」を試行する。
        /// Revit 2022 で利用可能な可能性のあるプロパティ名を順次試す。
        /// </summary>
        private static void TryAlternativeTotalsApi(ScheduleField field)
        {
            if (field == null) return;
            // 試行候補のプロパティ名 (Revit のバージョン違いを考慮)
            string[] candidates = new[]
            {
                "ShowTotals",
                "CalculateTotals",
                "HasCalculatedTotal",
                "IsCalculateTotalsCheck",
                "TotalsType",
            };
            var t = typeof(ScheduleField);
            foreach (var name in candidates)
            {
                var prop = t.GetProperty(name,
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                if (prop == null) continue;
                FormworkDebugLog.Log($"  [Sched] Found alternative property: {name} type={prop.PropertyType.Name} canWrite={prop.CanWrite}");
                if (!prop.CanWrite) continue;
                try
                {
                    if (prop.PropertyType == typeof(bool))
                    {
                        prop.SetValue(field, true, null);
                        FormworkDebugLog.Log($"  [Sched] Set {name}=true (bool)");
                    }
                }
                catch (Exception ex)
                {
                    LogEx($"set {name}", ex);
                }
            }

            // 試行候補のメソッド名
            string[] methodCandidates = new[]
            {
                "SetCalculateTotals",
                "SetHasTotals",
                "EnableTotals",
                "SetTotalsType",
            };
            foreach (var name in methodCandidates)
            {
                var methods = t.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance)
                    .Where(m => m.Name == name).ToList();
                foreach (var m in methods)
                {
                    var ps = m.GetParameters();
                    FormworkDebugLog.Log($"  [Sched] Found alternative method: {name}({string.Join(",", ps.Select(p => p.ParameterType.Name).ToArray())})");
                    // bool 1 引数なら呼んでみる
                    if (ps.Length == 1 && ps[0].ParameterType == typeof(bool))
                    {
                        try { m.Invoke(field, new object[] { true }); FormworkDebugLog.Log($"  [Sched] Invoked {name}(true)"); }
                        catch (Exception ex) { LogEx($"invoke {name}", ex); }
                    }
                }
            }
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

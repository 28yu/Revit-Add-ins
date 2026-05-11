using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Engine;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Output
{
    /// <summary>
    /// ViewSchedule を作成して DirectShape に格納した型枠情報を集計表示する。
    ///
    /// レイアウト（集計表示・タイプ名なし）:
    ///   - IsItemized = false → 各 DirectShape をソートキー単位で集約
    ///   - 列: 件数 / レベル / 部位 / 区分 / 面積(合計を計算)
    ///   - 並べ替え/グループ化:
    ///       1段目: レベル  (見出しON / フッタOFF)
    ///       2段目: 部位    (見出しON / フッタOFF)
    ///       3段目: 区分    (見出しOFF / フッタOFF)
    ///   - 面積フィールドの DisplayType=Totals → 各グループおよび総合計で面積合計を表示
    ///   - 総合計行は「合計のみ」モード (ShowGrandTotalTitle=false / ShowGrandTotalCount=false)
    /// </summary>
    internal static class ScheduleCreator
    {
        internal const string ScheduleName = "型枠数量集計";
        internal const string SummaryScheduleName = "型枠数量集計_合計";

        internal static ElementId CreateSchedule(Document doc, FormworkResult result = null)
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
            //   IsItemized=false: ソートキー単位で集約 (各インスタンスの内訳を非表示)
            //   ShowGrandTotal=true: 末尾に総合計行 (合計のみモード)
            try { def.IsItemized = false; } catch (Exception ex) { LogEx("IsItemized=false (before)", ex); }
            try { def.ShowGrandTotal = true; } catch (Exception ex) { LogEx("ShowGrandTotal=true (before)", ex); }
            LogDef(def, "after-set-itemized-before-fields");

            // 列順: 件数 / レベル / 部位 / 区分 / 面積
            var countField = AddCountField(def, schedulable);
            var levelField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamLevel);
            var partField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamCategory);
            var groupField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamGroupKey);
            var areaField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamArea);
            LogDef(def, "after-add-fields");
            LogField(areaField, "areaField after-add");

            // 面積フィールドは「合計を計算」を有効化（最終的にソート・グループ追加後に再設定）
            //   ScheduleField.DisplayType に ScheduleFieldDisplayType.Totals を設定すると
            //   集計表プロパティ「書式」タブのドロップダウンが「合計を計算」になる。
            if (areaField != null)
            {
                try { areaField.DisplayType = ScheduleFieldDisplayType.Totals; }
                catch (Exception ex) { LogEx("areaField.DisplayType=Totals (1st)", ex); }
                LogField(areaField, "areaField after-set-displaytype-1");
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

            // 階層グループ化: レベル → 部位（見出しON / フッタOFF）
            AddGroupField(def, levelField);
            AddGroupField(def, partField);
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

            // SortGroupField 追加後に最終設定（Revit が上書きする場合の保険）。
            // FieldId で再取得して設定（古い参照では Revit 内部状態と同期していない可能性）。
            try { def.IsItemized = false; } catch (Exception ex) { LogEx("IsItemized=false (after)", ex); }
            try { def.ShowGrandTotal = true; } catch { }

            // 総合計を「合計のみ」モードに設定: タイトル文字列と件数欄を非表示。
            // 合計タイトル文字列自体は設定するが、ShowGrandTotalTitle=false により非表示。
            // (ユーザーがモードを変更したときのためにタイトル文字列は保持)
            SetGrandTotalTitle(def, "型枠面積(合計)", showTitle: false, showCount: false);
            if (areaField != null)
            {
                try
                {
                    var freshAreaField = def.GetField(areaField.FieldId);
                    if (freshAreaField != null)
                    {
                        freshAreaField.DisplayType = ScheduleFieldDisplayType.Totals;
                        LogField(freshAreaField, "freshAreaField after-set-displaytype-2");
                    }
                }
                catch (Exception ex) { LogEx("freshAreaField.DisplayType=Totals", ex); }
            }
            LogDef(def, "FINAL");

            return schedule.Id;
        }

        /// <summary>
        /// 「型枠数量集計_合計」サマリ集計表を作成する。
        /// IsItemized=false により、フィルタ条件に合致する全 DirectShape が
        /// 1 行に集約され、面積の合計値が動的に表示される
        /// (DirectShape を手動削除すると Revit が自動再計算)。
        /// Header にラベル「型枠面積(合計)」を太字・赤字・薄黄背景で配置。
        /// </summary>
        internal static ElementId CreateSummarySchedule(Document doc)
        {
            DeleteScheduleByName(doc, SummaryScheduleName);

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

            try { schedule.Name = SummaryScheduleName; } catch { }

            var def = schedule.Definition;
            var schedulable = def.GetSchedulableFields();
            var paramIds = GetFormworkSharedParamIds(doc);

            FormworkDebugLog.Section("Summary Schedule Creation");

            // 件数 + 面積の 2 フィールド (面積は合計を計算)
            var countField = AddCountField(def, schedulable);
            var areaField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamArea);

            // 列ヘッダーに styled なラベルを設定 (Revit API は Body 行 0 のスタイル変更を許可)
            // 型枠オブジェクト削除に追従する動的値は、すぐ下のデータ行 (Body 行 1) で表示される
            if (countField != null)
            {
                try { countField.ColumnHeading = "件数(合計)"; }
                catch (Exception ex) { LogEx("countField.ColumnHeading", ex); }
            }
            if (areaField != null)
            {
                try { areaField.ColumnHeading = "型枠面積(合計)"; }
                catch (Exception ex) { LogEx("areaField.ColumnHeading", ex); }
                try { areaField.DisplayType = ScheduleFieldDisplayType.Totals; }
                catch (Exception ex) { LogEx("summary areaField.DisplayType=Totals", ex); }
            }

            // マーカーフィルタ (鉄骨除外を集計に含めない)
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
                catch (Exception ex) { LogEx("summary marker filter", ex); }
            }

            // 集計モード: アイテム別表示無効 + 総合計行は不要 (1 行で全件集計するため)
            try { def.IsItemized = false; }
            catch (Exception ex) { LogEx("summary IsItemized=false", ex); }
            try { def.ShowGrandTotal = false; } catch { }

            // フィールド追加後に再設定 (Revit が状態を上書きする場合の保険)
            if (areaField != null)
            {
                try
                {
                    var fresh = def.GetField(areaField.FieldId);
                    if (fresh != null) fresh.DisplayType = ScheduleFieldDisplayType.Totals;
                }
                catch { }
            }

            // ヘッダーにラベル行を追加するのではなく、Body の列ヘッダー (行 0) を styled に。
            // これにより列ヘッダー = 動的値の直上 という視覚的関連を保ちつつ、
            // データ行 (行 1) の値は Revit が DirectShape の追加・削除に応じて自動更新する。
            StyleColumnHeaders(schedule);

            FormworkDebugLog.Log($"  [Sched:Summary] created: '{schedule.Name}'");
            return schedule.Id;
        }

        /// <summary>
        /// 集計表の列ヘッダー (Body 行 0) を太字・赤字・薄黄背景でスタイル設定し、
        /// 列幅を広く確保してタイトル「&lt;型枠数量集計_合計&gt;」が改行しないようにする。
        /// データ行のフォントサイズも可能な限り拡大する。
        /// </summary>
        private static void StyleColumnHeaders(ViewSchedule schedule)
        {
            if (schedule == null) return;
            try
            {
                var tableData = schedule.GetTableData();
                if (tableData == null) return;
                var body = tableData.GetSectionData(SectionType.Body);
                if (body == null) return;

                int colCount = body.NumberOfColumns;
                int rowCount = body.NumberOfRows;
                FormworkDebugLog.Log(
                    $"  [Sched:Summary] body rows={rowCount} cols={colCount}");
                if (colCount <= 0 || rowCount <= 0) return;

                // 列幅を設定 (タイトル「<型枠数量集計_合計>」が改行しない総幅を確保しつつ
                // 余白過多にならないバランス)。Revit 内部単位は feet。約 50mm = 0.167 ft。
                for (int c = 0; c < colCount; c++)
                {
                    try
                    {
                        body.SetColumnWidth(c, 0.167);
                        FormworkDebugLog.Log($"  [Sched:Summary] col {c} width=0.167 ft (~50mm)");
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log(
                            $"  [Sched:Summary] SetColumnWidth col={c} EX: {ex.Message}");
                    }
                }

                // 列ヘッダー (Body 行 0) のスタイル: 太字・赤字・薄黄背景・大きめフォント
                // Revit 内部単位 (feet): デフォルト ≈ 0.0104 ft。0.0208 ft ≈ 倍 (16pt 相当)。
                var headerStyle = new TableCellStyle
                {
                    BackgroundColor = new Color(255, 240, 200), // 薄い黄色
                    TextColor = new Color(192, 0, 0),           // 赤
                    IsFontBold = true,
                    FontHorizontalAlignment = HorizontalAlignmentStyle.Center,
                };
                // FontSize はバージョンによっては存在しない可能性があるため try で
                TrySetTableCellStyleFontSize(headerStyle, 0.0208);

                var ov = headerStyle.GetCellStyleOverrideOptions();
                ov.BackgroundColor = true;
                ov.FontColor = true;
                ov.Bold = true;
                TrySetOverrideOption(ov, "FontSize", true);
                TrySetOverrideOption(ov, "HorizontalAlignment", true);
                headerStyle.SetCellStyleOverrideOptions(ov);

                for (int c = 0; c < colCount; c++)
                {
                    try { body.SetCellStyle(0, c, headerStyle); }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log(
                            $"  [Sched:Summary] SetCellStyle col={c} EX: {ex.Message}");
                    }
                }
                FormworkDebugLog.Log(
                    "  [Sched:Summary] column headers styled (bold/red/yellow, large font)");

                // データ行のフォントも拡大 (リフレクションで複数候補を試す)
                TryEnlargeDataFont(schedule, 0.0156);

                // Revit 2022 では TableCellStyle.FontSize が存在しないため、
                // Schedule View の「Body text / Header text」パラメータ (TextNoteType 参照)
                // を経由して大きめのテキストタイプを割り当てる。
                EnlargeViewTextStyles(schedule);
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log(
                    $"  [Sched:Summary] StyleColumnHeaders EX: {ex.GetType().Name}: {ex.Message}");
            }
        }

        /// <summary>
        /// データ行 (Body 行 1+) のフォントサイズを拡大する。
        /// ScheduleField の CellStyle / SetStyle をリフレクションで試行。
        /// </summary>
        private static void TryEnlargeDataFont(ViewSchedule schedule, double fontSize)
        {
            if (schedule == null) return;
            try
            {
                var def = schedule.Definition;
                int fc = def.GetFieldCount();
                for (int i = 0; i < fc; i++)
                {
                    ScheduleField field = null;
                    try { field = def.GetField(i); } catch { }
                    if (field == null) continue;
                    try { if (field.IsHidden) continue; } catch { }

                    var dataStyle = new TableCellStyle();
                    TrySetTableCellStyleFontSize(dataStyle, fontSize);
                    var ov = dataStyle.GetCellStyleOverrideOptions();
                    TrySetOverrideOption(ov, "FontSize", true);
                    dataStyle.SetCellStyleOverrideOptions(ov);

                    var t = field.GetType();
                    bool applied = false;

                    // 候補 1: プロパティ "CellStyle" の setter
                    try
                    {
                        var prop = t.GetProperty("CellStyle");
                        if (prop != null && prop.CanWrite &&
                            prop.PropertyType == typeof(TableCellStyle))
                        {
                            prop.SetValue(field, dataStyle);
                            applied = true;
                            FormworkDebugLog.Log($"  [Sched:Summary] field[{i}] CellStyle set");
                        }
                    }
                    catch (Exception ex) { LogReflEx($"field[{i}].CellStyle", ex); }

                    // 候補 2: メソッド "SetStyle(TableCellStyle)"
                    if (!applied)
                    {
                        try
                        {
                            var m = t.GetMethod("SetStyle", new[] { typeof(TableCellStyle) });
                            if (m != null)
                            {
                                m.Invoke(field, new object[] { dataStyle });
                                applied = true;
                                FormworkDebugLog.Log($"  [Sched:Summary] field[{i}] SetStyle invoked");
                            }
                        }
                        catch (Exception ex) { LogReflEx($"field[{i}].SetStyle", ex); }
                    }

                    if (!applied)
                        FormworkDebugLog.Log($"  [Sched:Summary] field[{i}] no font API found");
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log(
                    $"  [Sched:Summary] TryEnlargeDataFont EX: {ex.GetType().Name}: {ex.Message}");
            }
        }

        /// <summary>
        /// Schedule View の「Body text / Header text / Title text」パラメータに、
        /// 既存テキストタイプの中で最も大きいフォントサイズのものを割り当てる。
        /// これにより、TableCellStyle.FontSize が存在しない Revit 2022 環境でも
        /// 文字を大きく表示できる。
        /// </summary>
        private static void EnlargeViewTextStyles(ViewSchedule schedule)
        {
            if (schedule == null) return;
            try
            {
                var doc = schedule.Document;

                // プロジェクト内の TextNoteType を全て取得し、最大フォントサイズを選択
                ElementId largestId = ElementId.InvalidElementId;
                double largestSize = 0;
                foreach (TextNoteType tt in new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>())
                {
                    try
                    {
                        var p = tt.get_Parameter(BuiltInParameter.TEXT_SIZE);
                        if (p == null || !p.HasValue) continue;
                        double size = p.AsDouble();
                        if (size > largestSize)
                        {
                            largestSize = size;
                            largestId = tt.Id;
                        }
                    }
                    catch { }
                }

                if (largestId == ElementId.InvalidElementId)
                {
                    FormworkDebugLog.Log("  [Sched:Summary] no TextNoteType found for enlargement");
                    return;
                }

                FormworkDebugLog.Log(
                    $"  [Sched:Summary] largest text type id={largestId.IntegerValue} size={largestSize:F4} ft");

                // 候補となるパラメータ名 (英語版 Revit / 日本語版 Revit / 簡体中文版)
                string[] candidateNames = new[]
                {
                    "Body text", "Header text", "Title text",
                    "本体テキスト", "ヘッダー テキスト", "タイトル テキスト",
                    "ボディ テキスト", "ヘッダーテキスト", "タイトルテキスト",
                    "正文文字", "标题栏文字", "标题文字",
                    "タイトル", "ヘッダー", "本体",
                };

                // 1. Schedule View 自身に対して試す
                int assignedView = TrySetTextStyleParams(schedule, candidateNames, largestId, "view");

                // 2. Schedule View の Type 要素に対して試す
                //    (ログから「タイプ」パラメータでタイプ要素が参照されていることが判明)
                int assignedType = 0;
                ElementId typeId = ElementId.InvalidElementId;
                try { typeId = schedule.GetTypeId(); } catch { }
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    var typeElem = doc.GetElement(typeId);
                    if (typeElem != null)
                    {
                        FormworkDebugLog.Log($"  [Sched:Summary] schedule type id={typeId.IntegerValue}");
                        assignedType = TrySetTextStyleParams(typeElem, candidateNames, largestId, "type");

                        // Diagnostic: type 要素の settable ElementId パラメータを全列挙
                        FormworkDebugLog.Log("  [Sched:Summary] type element ElementId params:");
                        foreach (Parameter p in typeElem.Parameters)
                        {
                            try
                            {
                                if (p.StorageType == StorageType.ElementId && !p.IsReadOnly)
                                    FormworkDebugLog.Log(
                                        $"    typeParam='{p.Definition?.Name}' value={p.AsElementId()?.IntegerValue}");
                            }
                            catch { }
                        }
                    }
                }

                int totalAssigned = assignedView + assignedType;
                if (totalAssigned == 0)
                {
                    FormworkDebugLog.Log("  [Sched:Summary] no text-style params matched on view or type");
                }
                else
                {
                    FormworkDebugLog.Log(
                        $"  [Sched:Summary] text style assigned: view={assignedView} type={assignedType}");
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log(
                    $"  [Sched:Summary] EnlargeViewTextStyles EX: {ex.GetType().Name}: {ex.Message}");
            }
        }

        /// <summary>
        /// 候補となるパラメータ名のいずれかに合致した ElementId 型パラメータに、
        /// 指定の ElementId 値 (TextNoteType の Id) を設定する。
        /// </summary>
        private static int TrySetTextStyleParams(Element elem, string[] candidates, ElementId valueId, string label)
        {
            if (elem == null) return 0;
            int count = 0;
            foreach (var name in candidates)
            {
                try
                {
                    var p = elem.LookupParameter(name);
                    if (p != null && !p.IsReadOnly && p.StorageType == StorageType.ElementId)
                    {
                        p.Set(valueId);
                        FormworkDebugLog.Log($"  [Sched:Summary] {label}.'{name}' = {valueId.IntegerValue}");
                        count++;
                    }
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Sched:Summary] {label}.'{name}' set EX: {ex.Message}");
                }
            }
            return count;
        }

        /// <summary>
        /// TableCellStyle.FontSize プロパティを安全に設定する (バージョン差吸収)。
        /// </summary>
        private static void TrySetTableCellStyleFontSize(TableCellStyle style, double sizeFeet)
        {
            try
            {
                var p = style.GetType().GetProperty("FontSize");
                if (p != null && p.CanWrite)
                    p.SetValue(style, sizeFeet);
            }
            catch (Exception ex) { LogReflEx("TableCellStyle.FontSize", ex); }
        }

        /// <summary>
        /// TableCellStyleOverrideOptions の bool プロパティを安全に設定 (バージョン差吸収)。
        /// </summary>
        private static void TrySetOverrideOption(TableCellStyleOverrideOptions ov, string name, bool value)
        {
            try
            {
                var p = ov.GetType().GetProperty(name);
                if (p != null && p.CanWrite && p.PropertyType == typeof(bool))
                    p.SetValue(ov, value);
            }
            catch (Exception ex) { LogReflEx($"override.{name}", ex); }
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
            string displayType = "?";
            string fieldType = "?";
            string colHeading = "?";
            bool isCalculated = false;
            bool isHidden = false;
            try { displayType = field.DisplayType.ToString(); } catch { }
            try { fieldType = field.FieldType.ToString(); } catch { }
            try { colHeading = field.ColumnHeading; } catch { }
            try { isCalculated = field.IsCalculatedField; } catch { }
            try { isHidden = field.IsHidden; } catch { }
            FormworkDebugLog.Log(
                $"  [Sched:{label}] heading='{colHeading}' fieldType={fieldType} " +
                $"DisplayType={displayType} IsCalculated={isCalculated} IsHidden={isHidden}");
        }

        private static void LogEx(string action, Exception ex)
        {
            if (!FormworkDebugLog.Enabled) return;
            FormworkDebugLog.Log($"  [Sched:EX] {action}: {ex.GetType().Name}: {ex.Message}");
        }

        /// <summary>
        /// ScheduleDefinition の総合計タイトルとモードを設定する。
        /// showTitle/showCount で UI の「合計」ドロップダウン相当を制御:
        ///   - (true, true)  → タイトル、件数、合計
        ///   - (true, false) → タイトルと合計
        ///   - (false, true) → 件数と合計
        ///   - (false, false) → 合計のみ
        /// Revit のバージョンによってプロパティ名が異なる可能性があるため、
        /// リフレクションで複数候補を試す。失敗時は内部例外まで詳細ログに記録する。
        /// </summary>
        private static void SetGrandTotalTitle(
            ScheduleDefinition def, string title, bool showTitle, bool showCount)
        {
            if (def == null) return;
            var t = def.GetType();

            // タイトル文字列の設定は ShowGrandTotalTitle=true が前提となるバージョンがあるため、
            // 一旦 true に切り替えてから文字列を設定し、最後に希望の値へ戻す。
            TrySetBool(def, t, "ShowGrandTotalTitle", true);

            if (!string.IsNullOrEmpty(title))
            {
                // 1. プロパティ "GrandTotalTitle" を試す
                if (!TrySetString(def, t, "GrandTotalTitle", title))
                {
                    // 2. メソッド "SetGrandTotalTitle(string)" を試す
                    try
                    {
                        var method = t.GetMethod("SetGrandTotalTitle", new[] { typeof(string) });
                        if (method != null)
                        {
                            method.Invoke(def, new object[] { title });
                            FormworkDebugLog.Log($"  [Sched] GrandTotalTitle set via method: '{title}'");
                        }
                    }
                    catch (Exception ex) { LogReflEx("SetGrandTotalTitle method", ex); }
                }
            }

            // タイトル設定後に最終的な表示モードを反映
            TrySetBool(def, t, "ShowGrandTotalTitle", showTitle);
            TrySetBool(def, t, "ShowGrandTotalCount", showCount);
        }

        private static bool TrySetString(object obj, Type t, string propName, string value)
        {
            try
            {
                var p = t.GetProperty(propName);
                if (p != null && p.CanWrite && p.PropertyType == typeof(string))
                {
                    p.SetValue(obj, value);
                    FormworkDebugLog.Log($"  [Sched] {propName} set: '{value}'");
                    return true;
                }
            }
            catch (Exception ex) { LogReflEx($"{propName} (string)", ex); }
            return false;
        }

        private static void TrySetBool(object obj, Type t, string propName, bool value)
        {
            try
            {
                var p = t.GetProperty(propName);
                if (p != null && p.CanWrite && p.PropertyType == typeof(bool))
                {
                    p.SetValue(obj, value);
                    FormworkDebugLog.Log($"  [Sched] {propName} = {value}");
                }
            }
            catch (Exception ex) { LogReflEx($"{propName} (bool)", ex); }
        }

        /// <summary>
        /// リフレクション経由の TargetInvocationException から内部例外を取り出してログする。
        /// </summary>
        private static void LogReflEx(string action, Exception ex)
        {
            if (!FormworkDebugLog.Enabled) return;
            var inner = (ex is System.Reflection.TargetInvocationException tie && tie.InnerException != null)
                ? tie.InnerException : ex;
            FormworkDebugLog.Log(
                $"  [Sched:EX] {action}: {inner.GetType().FullName}: {inner.Message}");
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

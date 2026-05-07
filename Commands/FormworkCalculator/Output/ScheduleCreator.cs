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
    /// レイアウト（インスタンス内訳 + 階層集計）:
    ///   - IsItemized = true → 各 DirectShape を個別に表示
    ///   - 列: 件数 / レベル / 部位 / タイプ名 / 区分 / 面積(合計を計算)
    ///   - 階層グループ化 (ShowHeader + ShowFooter で各階層に合計行):
    ///       1段目: レベル  (参照レベル毎)
    ///       2段目: 部位    (カテゴリ毎)
    ///       3段目: タイプ名 (タイプ毎)
    ///   - 面積フィールドの DisplayType=Totals → 各グループおよび総合計で面積合計を表示
    ///   - 総合計行を表示
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

            // SortGroupField 追加後に最終設定（Revit が上書きする場合の保険）。
            // FieldId で再取得して設定（古い参照では Revit 内部状態と同期していない可能性）。
            try { def.IsItemized = true; } catch (Exception ex) { LogEx("IsItemized=true (after)", ex); }
            try { def.ShowGrandTotal = true; } catch { }

            // 総合計タイトルを「型枠面積(合計)」に設定し、Excel 総括表との対応を明確化。
            // Revit バージョンによりプロパティ名が異なる可能性があるためリフレクションで安全に設定。
            SetGrandTotalTitle(def, "型枠面積(合計)");
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

            // Body の総合計行は API でスタイル変更できない (Revit の制約)。
            // 代わりに Header セクションに強調表示用の「型枠面積(合計): X.XX ㎡」行を追加し、
            // 集計表名にも合計値を含めて Project Browser での識別性も向上させる。
            EmphasizeGrandTotal(schedule, result);

            return schedule.Id;
        }

        /// <summary>
        /// 集計表の総合計を視覚的に強調する。
        /// Revit API は Body の総合計行へのスタイル変更を許可しないため、
        /// Header セクションに「型枠面積(合計)」のラベルを styled な独立行として追加する。
        /// 値は Body 側の総合計行 (Revit が自動再計算) に任せ、ここではラベルのみとする
        /// （Header テキストは静的なので値を含めると DirectShape 削除時に追従しないため）。
        /// </summary>
        private static void EmphasizeGrandTotal(ViewSchedule schedule, FormworkResult result)
        {
            if (schedule == null) return;
            // ラベルのみ。値は body の総合計行 (Revit auto-updated) で見せる。
            string labelText = "型枠面積(合計) ↓";

            bool added = TryAddStyledHeaderRow(schedule, labelText);
            if (!added)
                FormworkDebugLog.Log("  [Sched:Style] Header row addition failed");
        }

        /// <summary>
        /// Header セクションに「型枠面積(合計): X.XX ㎡」の強調表示行を追加する。
        /// 列が不足していたら追加し、既存ヘッダー行の下に挿入してから
        /// セル結合・スタイル設定を行う。
        /// </summary>
        private static bool TryAddStyledHeaderRow(ViewSchedule schedule, string text)
        {
            try
            {
                var tableData = schedule.GetTableData();
                if (tableData == null) return false;
                var header = tableData.GetSectionData(SectionType.Header);
                if (header == null) return false;

                int initRows = header.NumberOfRows;
                int initCols = header.NumberOfColumns;
                FormworkDebugLog.Log(
                    $"  [Sched:Style] Header initial rows={initRows} cols={initCols}");

                // 列が無ければ 1 列追加
                if (initCols == 0)
                {
                    try { header.InsertColumn(0); }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"  [Sched:Style] Header InsertColumn EX: {ex.Message}");
                        return false;
                    }
                }

                // 既存ヘッダー行の最後に新しい行を追加 (一番下 = body の真上)
                int insertAt = header.NumberOfRows;
                try { header.InsertRow(insertAt); }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Sched:Style] Header InsertRow at {insertAt} EX: {ex.Message}");
                    return false;
                }

                int newRow = insertAt;
                int colCount = header.NumberOfColumns;

                // セル結合 (新規行を全列にまたいで 1 セルに)
                if (colCount > 1)
                {
                    try
                    {
                        var merge = new TableMergedCell(newRow, 0, newRow, colCount - 1);
                        header.MergeCells(merge);
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"  [Sched:Style] Header MergeCells EX: {ex.Message}");
                    }
                }

                // テキスト設定
                try { header.SetCellText(newRow, 0, text); }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Sched:Style] Header SetCellText EX: {ex.Message}");
                    return false;
                }

                // スタイル設定 (太字・赤字・薄黄背景)
                var style = new TableCellStyle
                {
                    BackgroundColor = new Color(255, 240, 200), // 薄い黄色
                    TextColor = new Color(192, 0, 0),           // 赤
                    IsFontBold = true,
                };
                var ov = style.GetCellStyleOverrideOptions();
                ov.BackgroundColor = true;
                ov.FontColor = true;
                ov.Bold = true;
                style.SetCellStyleOverrideOptions(ov);

                try
                {
                    header.SetCellStyle(newRow, 0, style);
                    FormworkDebugLog.Log(
                        $"  [Sched:Style] Header row {newRow} added & styled: '{text}'");
                    return true;
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Sched:Style] Header SetCellStyle EX: {ex.Message}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log(
                    $"  [Sched:Style] EmphasizeGrandTotal Header EX: {ex.GetType().Name}: {ex.Message}");
                return false;
            }
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
                };

                int assigned = 0;
                foreach (var name in candidateNames)
                {
                    try
                    {
                        var p = schedule.LookupParameter(name);
                        if (p != null && !p.IsReadOnly && p.StorageType == StorageType.ElementId)
                        {
                            p.Set(largestId);
                            FormworkDebugLog.Log($"  [Sched:Summary] '{name}' = {largestId.IntegerValue}");
                            assigned++;
                        }
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"  [Sched:Summary] '{name}' set EX: {ex.Message}");
                    }
                }

                // 名前で見つからない場合は全 ElementId 型パラメータをログ (デバッグ用)
                if (assigned == 0)
                {
                    FormworkDebugLog.Log("  [Sched:Summary] no text-style params matched. listing schedule view ElementId params:");
                    foreach (Parameter p in schedule.Parameters)
                    {
                        try
                        {
                            if (p.StorageType == StorageType.ElementId && !p.IsReadOnly)
                                FormworkDebugLog.Log(
                                    $"    paramName='{p.Definition?.Name}' value={p.AsElementId()?.IntegerValue}");
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log(
                    $"  [Sched:Summary] EnlargeViewTextStyles EX: {ex.GetType().Name}: {ex.Message}");
            }
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
        /// ScheduleDefinition の総合計タイトルを設定する。Revit のバージョンによって
        /// プロパティ名が異なる可能性があるため、リフレクションで複数候補を試す。
        /// 失敗時は内部例外まで詳細ログに記録する。
        /// </summary>
        private static void SetGrandTotalTitle(ScheduleDefinition def, string title)
        {
            if (def == null || string.IsNullOrEmpty(title)) return;
            var t = def.GetType();

            // 関連 bool プロパティが存在する場合、先に true にしておく
            // (一部の Revit バージョンでは ShowGrandTotalTitle = true が前提条件のため)
            TrySetBool(def, t, "ShowGrandTotalTitle", true);
            TrySetBool(def, t, "ShowGrandTotalCount", true);

            // 1. プロパティ "GrandTotalTitle" を試す
            if (TrySetString(def, t, "GrandTotalTitle", title)) return;

            // 2. メソッド "SetGrandTotalTitle(string)" を試す
            try
            {
                var method = t.GetMethod("SetGrandTotalTitle", new[] { typeof(string) });
                if (method != null)
                {
                    method.Invoke(def, new object[] { title });
                    FormworkDebugLog.Log($"  [Sched] GrandTotalTitle set via method: '{title}'");
                    return;
                }
            }
            catch (Exception ex) { LogReflEx("SetGrandTotalTitle method", ex); }

            // 3. 診断: ScheduleDefinition の全 settable プロパティを列挙してログに残す
            try
            {
                FormworkDebugLog.Log("  [Sched] -- ScheduleDefinition settable string properties --");
                foreach (var p in t.GetProperties())
                {
                    if (p.CanWrite && p.PropertyType == typeof(string))
                        FormworkDebugLog.Log($"    propString: {p.Name}");
                    else if (p.CanWrite && p.PropertyType == typeof(bool))
                        FormworkDebugLog.Log($"    propBool:   {p.Name}");
                }
                foreach (var m in t.GetMethods())
                {
                    if (m.Name.StartsWith("Set") &&
                        (m.Name.Contains("Total") || m.Name.Contains("Title")))
                        FormworkDebugLog.Log($"    method:     {m.Name}({string.Join(",", m.GetParameters().Select(p => p.ParameterType.Name))})");
                }
            }
            catch { }

            FormworkDebugLog.Log("  [Sched] GrandTotalTitle could not be set in this Revit version");
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

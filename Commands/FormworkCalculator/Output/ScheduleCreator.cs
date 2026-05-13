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

        /// <summary>
        /// 型枠数量集計表を作成する。
        /// sourceFilter が null の場合: 全ソース対象 (旧来の動作)
        /// sourceFilter が指定された場合: その SourceName のみフィルタした集計表を作成
        ///   (ホスト用 / 各リンク用に個別の集計表を作るための引数)
        /// </summary>
        internal static ElementId CreateSchedule(Document doc, FormworkResult result = null,
            string sourceFilter = null)
        {
            string scheduleName = string.IsNullOrEmpty(sourceFilter)
                ? ScheduleName
                : $"{ScheduleName} - {sourceFilter}";
            // 同名の既存集計表に加え、過去実行で作成された "型枠数量集計 - xxx" 形式の
            // 集計表もまとめて削除する (初回の "ホスト" 集計表生成時のみ実行)。
            if (string.IsNullOrEmpty(sourceFilter)
                || sourceFilter == ElementSourceRegistry.HostSourceName)
            {
                DeleteAllFormworkSchedules(doc);
            }
            else
            {
                DeleteScheduleByName(doc, scheduleName);
            }

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

            try { schedule.Name = scheduleName; } catch { }

            var def = schedule.Definition;
            var schedulable = def.GetSchedulableFields();
            var paramIds = GetFormworkSharedParamIds(doc);

            FormworkDebugLog.Section($"Schedule Creation: {scheduleName}");
            LogDef(def, "after-create");

            // 集計表の基本設定 (フィールド追加前に設定する必要がある場合がある)
            //   IsItemized=false: ソートキー単位で集約 (各インスタンスの内訳を非表示)
            //   ShowGrandTotal=true: 末尾に総合計行 (合計のみモード)
            try { def.IsItemized = false; } catch (Exception ex) { LogEx("IsItemized=false (before)", ex); }
            try { def.ShowGrandTotal = true; } catch (Exception ex) { LogEx("ShowGrandTotal=true (before)", ex); }
            LogDef(def, "after-set-itemized-before-fields");

            // 列順: 件数 / レベル / 部位 / 区分 / 面積
            // ソース列はユーザー要望により非表示 (ホスト/リンク別の集計表として作成するため)
            var countField = AddCountField(def, schedulable);
            var levelField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamLevel);
            var partField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamCategory);
            var groupField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamGroupKey);
            var areaField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamArea);
            LogDef(def, "after-add-fields");
            LogField(areaField, "areaField after-add");

            // ソースフィルタ用フィールド (非表示)
            ScheduleField sourceField = null;
            if (!string.IsNullOrEmpty(sourceFilter))
            {
                sourceField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamSource);
                if (sourceField != null)
                {
                    try { sourceField.IsHidden = true; } catch { }
                }
            }

            // 列ヘッダーを短く設定 (パラメータ名 "28Tools_Formwork_xxx" が長いため
            // セル幅が大きくなりすぎないように短い見出しに置き換える)。
            if (levelField != null)
            {
                try { levelField.ColumnHeading = "レベル"; }
                catch (Exception ex) { LogEx("levelField.ColumnHeading", ex); }
            }
            if (partField != null)
            {
                try { partField.ColumnHeading = "部位"; }
                catch (Exception ex) { LogEx("partField.ColumnHeading", ex); }
            }
            if (groupField != null)
            {
                try { groupField.ColumnHeading = "区分"; }
                catch (Exception ex) { LogEx("groupField.ColumnHeading", ex); }
            }
            if (areaField != null)
            {
                try { areaField.ColumnHeading = "型枠面積"; }
                catch (Exception ex) { LogEx("areaField.ColumnHeading", ex); }
            }

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

            // ソースフィルタ (sourceField が指定されている場合のみ)
            if (sourceField != null && !string.IsNullOrEmpty(sourceFilter))
            {
                try
                {
                    var sourceFlt = new ScheduleFilter(
                        sourceField.FieldId,
                        ScheduleFilterType.Equal,
                        sourceFilter);
                    def.AddFilter(sourceFlt);
                }
                catch (Exception ex) { LogEx("source filter", ex); }
            }
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

            // 外観タブ「データの前に空白行を挿入」を OFF に設定。
            // Revit API のプロパティ名はバージョンによって異なるためリフレクションで複数候補を試す。
            TrySetBlankLineBeforeData(def, false);

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

            // 列幅を内容に応じて自動調整 (文字と罫線の余白を保ちつつ改行を防ぐ)
            try { doc.Regenerate(); } catch { }
            SetMainScheduleColumnWidths(schedule);

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

            // フィールド: 件数 / ソース / 面積。
            // 「ソース」でグループ化することで、ホスト・各リンクモデル毎に 1 行ずつ表示し、
            // 末尾に ShowGrandTotal による全体合計を表示する。
            var countField = AddCountField(def, schedulable);
            var sourceField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamSource);
            var areaField = AddField(doc, def, schedulable, paramIds, FormworkParameterManager.ParamArea);

            // 列ヘッダーに styled なラベルを設定
            if (countField != null)
            {
                try { countField.ColumnHeading = "件数(合計)"; }
                catch (Exception ex) { LogEx("countField.ColumnHeading", ex); }
            }
            if (sourceField != null)
            {
                try { sourceField.ColumnHeading = "ソース"; }
                catch (Exception ex) { LogEx("summary sourceField.ColumnHeading", ex); }
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

            // ソースでソート (ホスト → リンク名順)。グループ見出し/フッターは非表示
            // (IsItemized=false により、ソースごとに 1 行に集約される)
            if (sourceField != null)
            {
                try
                {
                    var sortSource = new ScheduleSortGroupField(sourceField.FieldId)
                    {
                        ShowHeader = false,
                        ShowFooter = false,
                        ShowBlankLine = false,
                        SortOrder = ScheduleSortOrder.Ascending,
                    };
                    def.AddSortGroupField(sortSource);
                }
                catch (Exception ex) { LogEx("summary sourceField sortGroup", ex); }
            }

            // 集計モード: アイテム別表示無効 + 総合計行を表示 (ソース毎の小計 + 全体合計)
            try { def.IsItemized = false; }
            catch (Exception ex) { LogEx("summary IsItemized=false", ex); }
            try { def.ShowGrandTotal = true; } catch { }
            SetGrandTotalTitle(def, "全体合計", showTitle: true, showCount: true);

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
            try { doc.Regenerate(); } catch { }
            StyleColumnHeaders(schedule);

            FormworkDebugLog.Log($"  [Sched:Summary] created: '{schedule.Name}'");
            return schedule.Id;
        }

        /// <summary>
        /// メイン集計表の列幅を、各列の実際のセル内容の最大文字幅に応じて自動調整する。
        /// 全角 (CJK) は半角の2倍幅として概算し、両側に一定の余白を加える。
        /// </summary>
        private static void SetMainScheduleColumnWidths(ViewSchedule schedule)
        {
            if (schedule == null) return;
            try
            {
                var tableData = schedule.GetTableData();
                var body = tableData?.GetSectionData(SectionType.Body);
                if (body == null) return;

                int colCount = body.NumberOfColumns;
                int rowCount = body.NumberOfRows;
                FormworkDebugLog.Log(
                    $"  [Sched] AutoWidth: body rows={rowCount} cols={colCount}");

                // 既定の最小幅 (文字数 0 でもこれ以上は確保)。
                // 列順: 件数 / レベル / 部位 / 区分 / 面積。
                double[] minWidthsFeet = { 0.075, 0.080, 0.075, 0.085, 0.090 };
                // 各列の最大幅 (mm) を制限 (異常に広がるのを抑止)
                const double maxColumnWidthMm = 200.0;

                for (int c = 0; c < colCount; c++)
                {
                    double maxUnits = 0;
                    for (int r = 0; r < rowCount; r++)
                    {
                        string text = string.Empty;
                        try { text = body.GetCellText(r, c) ?? string.Empty; } catch { }
                        double u = MeasureTextUnits(text);
                        if (u > maxUnits) maxUnits = u;
                    }

                    // 単位 → mm 換算: 半角 1 単位 = 約 2.0mm (Revit 既定 2.5mm フォント想定)
                    // 余白: 両側合計 7mm を加える
                    double widthMm = maxUnits * 2.0 + 7.0;
                    double widthFeet = widthMm / 304.8;

                    double minFeet = c < minWidthsFeet.Length ? minWidthsFeet[c] : 0.080;
                    if (widthFeet < minFeet) widthFeet = minFeet;
                    double maxFeet = maxColumnWidthMm / 304.8;
                    if (widthFeet > maxFeet) widthFeet = maxFeet;

                    try
                    {
                        body.SetColumnWidth(c, widthFeet);
                        FormworkDebugLog.Log(
                            $"  [Sched] col {c} maxUnits={maxUnits:F1} → width={widthFeet:F3} ft (~{widthFeet * 304.8:F0}mm)");
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"  [Sched] SetColumnWidth col={c} EX: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Sched] SetMainScheduleColumnWidths EX: {ex.Message}");
            }
        }

        /// <summary>
        /// 文字幅の概算単位を返す。半角 = 1.0、全角 (CJK) = 2.0 として加算。
        /// </summary>
        private static double MeasureTextUnits(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            double w = 0;
            foreach (char ch in s)
            {
                if (IsFullWidth(ch)) w += 2.0;
                else w += 1.0;
            }
            return w;
        }

        /// <summary>
        /// 全角文字 (CJK ひらがな・カタカナ・漢字・全角記号) の判定。
        /// </summary>
        private static bool IsFullWidth(char ch)
        {
            return (ch >= 0x3000 && ch <= 0x30FF)   // CJK記号・ひらがな・カタカナ
                || (ch >= 0x3400 && ch <= 0x9FFF)   // CJK統合漢字 (拡張Aを含む)
                || (ch >= 0xF900 && ch <= 0xFAFF)   // CJK互換漢字
                || (ch >= 0xFF00 && ch <= 0xFF60)   // 全角形式
                || (ch >= 0xFFE0 && ch <= 0xFFE6);  // 全角記号
        }

        /// <summary>
        /// 集計表の列ヘッダー (Body 行 0) を太字・赤字・薄黄背景でスタイル設定し、
        /// 列幅を広く確保してタイトル「&lt;型枠数量集計_合計&gt;」が改行しないようにする。
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

                // 列幅を内容に応じて自動調整 (文字と罫線の余白を保ちつつ改行を防ぐ)
                for (int c = 0; c < colCount; c++)
                {
                    double maxUnits = 0;
                    for (int r = 0; r < rowCount; r++)
                    {
                        string text = string.Empty;
                        try { text = body.GetCellText(r, c) ?? string.Empty; } catch { }
                        double u = MeasureTextUnits(text);
                        if (u > maxUnits) maxUnits = u;
                    }
                    // 余白: 両側合計 7mm。件数(合計)・型枠面積(合計) は最小幅 30mm 確保
                    double widthMm = Math.Max(30.0, maxUnits * 2.0 + 7.0);
                    double widthFeet = widthMm / 304.8;
                    try
                    {
                        body.SetColumnWidth(c, widthFeet);
                        FormworkDebugLog.Log(
                            $"  [Sched:Summary] col {c} maxUnits={maxUnits:F1} → width={widthFeet:F3} ft (~{widthFeet * 304.8:F0}mm)");
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log(
                            $"  [Sched:Summary] SetColumnWidth col={c} EX: {ex.Message}");
                    }
                }

                // 列ヘッダー (Body 行 0) のスタイル: 太字・赤字、列ごとの背景色。
                //   件数列 (col 0)   = 薄黄
                //   面積列 (col 1)   = ピンク
                //   それ以外 (隠し列) = 薄黄
                // Revit 2022 ではデータ行 (Body 行 1+) のスタイル変更は API で許可されないため、
                // 列ヘッダーセルの背景でピンクを表現する。
                var yellow = new Color(255, 240, 200);
                var pink = new Color(255, 200, 220);
                var redText = new Color(192, 0, 0);
                const int areaColumnIndex = 1; // 件数(0) / 面積(1) / [hidden marker(2)]

                for (int c = 0; c < colCount; c++)
                {
                    var style = new TableCellStyle
                    {
                        BackgroundColor = (c == areaColumnIndex) ? pink : yellow,
                        TextColor = redText,
                        IsFontBold = true,
                        FontHorizontalAlignment = HorizontalAlignmentStyle.Center,
                    };
                    var ov = style.GetCellStyleOverrideOptions();
                    ov.BackgroundColor = true;
                    ov.FontColor = true;
                    ov.Bold = true;
                    TrySetOverrideOption(ov, "HorizontalAlignment", true);
                    style.SetCellStyleOverrideOptions(ov);

                    try { body.SetCellStyle(0, c, style); }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log(
                            $"  [Sched:Summary] SetCellStyle col={c} EX: {ex.Message}");
                    }
                }

                // 試行: データ行 (Body 行 1) の面積セルをピンクに。
                // Revit 2022 は body のデータ行スタイル上書きを許可しないため通常は例外。
                // 将来の Revit でサポートされた場合に動作するよう試行のみ行う。
                if (rowCount > 1 && colCount > areaColumnIndex)
                {
                    try
                    {
                        var dataPinkStyle = new TableCellStyle
                        {
                            BackgroundColor = pink,
                        };
                        var ov2 = dataPinkStyle.GetCellStyleOverrideOptions();
                        ov2.BackgroundColor = true;
                        dataPinkStyle.SetCellStyleOverrideOptions(ov2);
                        body.SetCellStyle(1, areaColumnIndex, dataPinkStyle);
                        FormworkDebugLog.Log(
                            "  [Sched:Summary] data row area cell colored pink (Revit allowed it)");
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log(
                            $"  [Sched:Summary] data row pink EX (expected on Revit 2022): {ex.Message}");
                    }
                }

                FormworkDebugLog.Log(
                    "  [Sched:Summary] column headers styled (bold/red, area col pink)");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log(
                    $"  [Sched:Summary] StyleColumnHeaders EX: {ex.GetType().Name}: {ex.Message}");
            }
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
        /// 「データの前に空白行を挿入」(外観タブ) を設定する。
        /// ScheduleDefinition の bool プロパティとして複数の候補名でリフレクション試行する。
        /// </summary>
        private static void TrySetBlankLineBeforeData(ScheduleDefinition def, bool value)
        {
            if (def == null) return;
            var t = def.GetType();
            // 複数バージョンで確認された候補名を優先度順に試す
            var candidates = new[]
            {
                "ShowBlankLines",
                "InsertLineBeforeData",
                "BlankLineBeforeData",
                "ShowLeadingBlankLine",
                "AddBlankLineBeforeData",
                "InsertBlankLineBeforeData",
            };
            bool applied = false;
            foreach (var name in candidates)
            {
                var p = t.GetProperty(name);
                if (p != null && p.CanWrite && p.PropertyType == typeof(bool))
                {
                    try
                    {
                        p.SetValue(def, value);
                        FormworkDebugLog.Log($"  [Sched] BlankLineBeforeData via '{name}' = {value}");
                        applied = true;
                        break;
                    }
                    catch (Exception ex) { LogReflEx($"BlankLineBeforeData.{name}", ex); }
                }
            }
            if (!applied)
                FormworkDebugLog.Log("  [Sched] BlankLineBeforeData: no matching property found (manual setting required)");
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
        /// グループフィールドとして追加（見出しON / フッタOFF）。
        /// </summary>
        private static void AddGroupField(ScheduleDefinition def, ScheduleField field)
        {
            if (field == null) return;
            try
            {
                var sort = new ScheduleSortGroupField(field.FieldId)
                {
                    ShowHeader = true,
                    ShowFooter = false,
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

        /// <summary>
        /// 「型枠数量集計」および「型枠数量集計 - {ソース名}」形式の集計表をまとめて削除する。
        /// 再実行時の前段クリーンアップに使う (リンク構成が変わって不要な集計表が残るのを防ぐ)。
        /// サマリ集計表 (型枠数量集計_合計) は別途 CreateSummarySchedule で削除される。
        /// </summary>
        private static void DeleteAllFormworkSchedules(Document doc)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(v => !v.IsTemplate
                    && (v.Name == ScheduleName
                        || v.Name.StartsWith(ScheduleName + " - ")))
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

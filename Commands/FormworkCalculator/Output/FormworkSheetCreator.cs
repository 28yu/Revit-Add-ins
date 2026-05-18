using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Engine;

namespace Tools28.Commands.FormworkCalculator.Output
{
    /// <summary>
    /// 型枠数量算出の結果（3Dビュー・集計表）を配置するシートを自動作成する。
    ///
    /// 仕様:
    ///   - 既存シートで最も多く使われている図枠 (TitleBlock FamilySymbol) を採用
    ///   - レイアウト:
    ///       上: サマリ集計表 (型枠数量集計_合計)
    ///       中段: ホスト集計表 + 各リンク集計表を左から右へ並べる
    ///             (シート幅を超える場合は次の行に折り返し)
    ///       右下: 3D ビュー
    ///   - シート名: 「型枠数量集計」, シート番号: 自動 (型枠-NNN)
    /// </summary>
    internal static class FormworkSheetCreator
    {
        internal const string SheetName = "型枠数量集計";

        /// <summary>
        /// プロジェクト内の全 型枠分析 3D ビュー (タグ + 新旧名前パターン) を収集する。
        /// </summary>
        internal static List<ElementId> CollectAllAnalysisViewIds(Document doc)
        {
            var list = FormworkOutputFinder.FindAllAnalysisViews(doc)
                .Select(v => v.Id)
                .ToList();
            FormworkDebugLog.Log($"  [Sheet] CollectAllAnalysisViewIds: {list.Count}件");
            return list;
        }

        /// <summary>
        /// プロジェクト内の全 ビュー別 型枠集計表 (「型枠数量集計 - xxx」) を収集する。
        /// サマリ (「型枠数量集計_合計」) は除く。
        /// </summary>
        internal static List<ElementId> CollectAllPerViewScheduleIds(Document doc)
        {
            var list = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(v => !v.IsTemplate
                    && v.Name != ScheduleCreator.SummaryScheduleName
                    && (v.Name == ScheduleCreator.ScheduleName
                        || v.Name.StartsWith(ScheduleCreator.ScheduleName + " - ")))
                .OrderBy(v => v.Name)
                .Select(v => v.Id)
                .ToList();
            FormworkDebugLog.Log($"  [Sheet] CollectAllPerViewScheduleIds: {list.Count}件");
            return list;
        }

        /// <summary>
        /// 既存の型枠数量集計シートがあれば削除して、現在のビュー・集計表で再構築する。
        /// 既存シート名 / 番号は保持して同一の位置に再作成する。
        /// </summary>
        internal static ElementId CreateOrUpdateSheet(
            Document doc,
            IList<ElementId> analysisViewIds,
            IList<ElementId> mainScheduleIds,
            ElementId summaryScheduleId)
        {
            if (doc == null) return null;

            // 既存の型枠シートを検索し、シート名・番号を保存してから削除 (タグ + 名前フォールバック)
            string existingName = null;
            string existingNumber = null;
            try
            {
                var existing = FormworkOutputFinder.FindFormworkSheet(doc);
                if (existing != null)
                {
                    existingName = existing.Name;
                    existingNumber = existing.SheetNumber;
                    doc.Delete(existing.Id);
                    FormworkDebugLog.Log($"  [Sheet] 既存シート削除: '{existingNumber} - {existingName}'");
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Sheet] 既存シート削除 EX: {ex.Message}");
            }

            return CreateSheetMulti(doc, analysisViewIds, mainScheduleIds, summaryScheduleId,
                preferredName: existingName, preferredNumber: existingNumber);
        }

        /// <summary>
        /// 複数の 3D ビュー (型枠分析ビュー) を配置可能な集計シートを作成する。
        /// レイアウト:
        ///   上段: サマリ集計表
        ///   中段: ビュー別集計表を横並び配置 (折り返しあり)
        ///   下段: 各 3D ビューのビューポートをグリッドで配置 (タイトル非表示)
        /// </summary>
        internal static ElementId CreateSheetMulti(
            Document doc,
            IList<ElementId> view3DIds,
            IList<ElementId> mainScheduleIds,
            ElementId summaryScheduleId,
            string preferredName = null,
            string preferredNumber = null)
        {
            if (doc == null) return null;

            ElementId titleBlockTypeId = FindMostUsedTitleBlock(doc);
            ViewSheet sheet;
            try
            {
                sheet = ViewSheet.Create(doc, titleBlockTypeId ?? ElementId.InvalidElementId);
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Sheet] Create EX: {ex.Message}");
                return null;
            }
            if (sheet == null) return null;

            try { sheet.Name = !string.IsNullOrEmpty(preferredName) ? preferredName : UniqueSheetName(doc, SheetName); } catch { }
            try { sheet.SheetNumber = !string.IsNullOrEmpty(preferredNumber) ? preferredNumber : FindNextSheetNumber(doc); } catch { }

            // 出力タグ書き込み (リネーム耐性のための識別子)
            try
            {
                FormworkParameterManager.SetOutputTag(
                    sheet,
                    FormworkParameterManager.OutputKindSheet,
                    string.Empty);
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Sheet] SetOutputTag EX: {ex.Message}");
            }

            var draw = GetDrawableArea(doc, sheet);
            const double margin = 0.082;
            const double gap = 0.05;

            double leftX = draw.Min.U + margin;
            double topY = draw.Max.V - margin;
            double rightX = draw.Max.U - margin;
            double bottomY = draw.Min.V + margin;
            double currentTopY = topY;

            // サマリ集計表
            if (summaryScheduleId != null && summaryScheduleId != ElementId.InvalidElementId)
            {
                var sumRes = PlaceScheduleAt(doc, sheet, summaryScheduleId, leftX, currentTopY);
                if (sumRes.HasValue) currentTopY = sumRes.Value.bottomY - gap;
            }

            // ビュー別集計表を横並び (折り返しあり)
            if (mainScheduleIds != null && mainScheduleIds.Count > 0)
            {
                double rowTopY = currentTopY;
                double rowLeftX = leftX;
                double rowMaxBottomY = rowTopY;

                foreach (var id in mainScheduleIds)
                {
                    if (id == null || id == ElementId.InvalidElementId) continue;
                    var res = PlaceScheduleAt(doc, sheet, id, rowLeftX, rowTopY);
                    if (!res.HasValue) continue;
                    double placedRight = res.Value.rightX;
                    double placedBottom = res.Value.bottomY;
                    if (placedRight > rightX && rowLeftX > leftX)
                    {
                        try { doc.Delete(res.Value.instanceId); } catch { }
                        rowTopY = rowMaxBottomY - gap;
                        rowLeftX = leftX;
                        rowMaxBottomY = rowTopY;
                        var res2 = PlaceScheduleAt(doc, sheet, id, rowLeftX, rowTopY);
                        if (!res2.HasValue) continue;
                        placedRight = res2.Value.rightX;
                        placedBottom = res2.Value.bottomY;
                    }
                    if (placedBottom < rowMaxBottomY) rowMaxBottomY = placedBottom;
                    rowLeftX = placedRight + gap;
                }
                currentTopY = rowMaxBottomY - gap;
            }

            // 3D ビューをグリッド配置 (集計表の下、シート下端まで)
            if (view3DIds != null && view3DIds.Count > 0)
            {
                PlaceViewportGrid(doc, sheet, view3DIds,
                    leftX, rightX, currentTopY, bottomY);
            }

            FormworkDebugLog.Log(
                $"  [Sheet] created: '{sheet.SheetNumber} - {sheet.Name}' " +
                $"views={view3DIds?.Count ?? 0} schedules={mainScheduleIds?.Count ?? 0}");
            return sheet.Id;
        }

        /// <summary>
        /// 3D ビューを横並びで配置し、シート幅を超えたら次の行に折り返す。
        /// 各ビューポートの実サイズ (GetBoxOutline) を使って互いに重ならないよう配置する。
        /// </summary>
        private static void PlaceViewportGrid(
            Document doc, ViewSheet sheet, IList<ElementId> viewIds,
            double leftX, double rightX, double topY, double bottomY)
        {
            int n = viewIds.Count;
            if (n == 0) return;

            const double vpGap = 0.05; // ≈15mm

            double rowTopY = topY;
            double rowLeftX = leftX;
            double rowMaxBottomY = rowTopY;

            for (int i = 0; i < n; i++)
            {
                // 仮配置 (シート外の安全な位置) で実サイズを測定
                double tmpX = leftX + 1000.0;
                double tmpY = topY;
                Viewport vp;
                try
                {
                    vp = Viewport.Create(doc, sheet.Id, viewIds[i], new XYZ(tmpX, tmpY, 0));
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Sheet] Viewport.Create EX (grid {i}): {ex.Message}");
                    continue;
                }
                if (vp == null) continue;
                ApplyNoTitleViewportType(doc, vp);

                double w = 0.7, h = 0.5; // フォールバック
                double curCx = tmpX, curCy = tmpY;
                try
                {
                    var ol = vp.GetBoxOutline();
                    if (ol != null)
                    {
                        w = ol.MaximumPoint.X - ol.MinimumPoint.X;
                        h = ol.MaximumPoint.Y - ol.MinimumPoint.Y;
                        curCx = (ol.MinimumPoint.X + ol.MaximumPoint.X) / 2.0;
                        curCy = (ol.MinimumPoint.Y + ol.MaximumPoint.Y) / 2.0;
                    }
                }
                catch { }

                // 行の右端を超えるなら次の行に折り返す (ただし行の先頭は折り返さない)
                double targetLeft = rowLeftX;
                double targetTop = rowTopY;
                if (targetLeft + w > rightX && rowLeftX > leftX)
                {
                    rowTopY = rowMaxBottomY - vpGap;
                    rowLeftX = leftX;
                    targetLeft = rowLeftX;
                    targetTop = rowTopY;
                }

                double targetCx = targetLeft + w / 2.0;
                double targetCy = targetTop - h / 2.0;
                try
                {
                    ElementTransformUtils.MoveElement(doc, vp.Id,
                        new XYZ(targetCx - curCx, targetCy - curCy, 0));
                }
                catch (Exception ex)
                {
                    FormworkDebugLog.Log($"  [Sheet] Viewport move EX (grid {i}): {ex.Message}");
                }

                double placedBottom = targetTop - h;
                if (placedBottom < rowMaxBottomY) rowMaxBottomY = placedBottom;
                rowLeftX = targetLeft + w + vpGap;
            }
        }

        /// <summary>
        /// 集計シートを作成し、ビュー・集計表を配置する。
        /// </summary>
        /// <param name="mainScheduleIds">
        /// ホスト集計表 (先頭) + 各リンク集計表 (続き) の ID リスト。
        /// 横並びで配置され、幅が足りなければ次行に折り返す。
        /// </param>
        internal static ElementId CreateSheet(
            Document doc,
            ElementId view3DId,
            IList<ElementId> mainScheduleIds,
            ElementId summaryScheduleId)
        {
            if (doc == null) return null;

            ElementId titleBlockTypeId = FindMostUsedTitleBlock(doc);

            ViewSheet sheet;
            try
            {
                sheet = ViewSheet.Create(doc,
                    titleBlockTypeId ?? ElementId.InvalidElementId);
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Sheet] Create EX: {ex.Message}");
                return null;
            }
            if (sheet == null) return null;

            try { sheet.Name = UniqueSheetName(doc, SheetName); } catch { }
            try { sheet.SheetNumber = FindNextSheetNumber(doc); } catch { }

            // 出力タグ書き込み (リネーム耐性のための識別子)
            try
            {
                FormworkParameterManager.SetOutputTag(
                    sheet,
                    FormworkParameterManager.OutputKindSheet,
                    string.Empty);
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Sheet] SetOutputTag EX: {ex.Message}");
            }

            // 配置可能領域 (図枠の内側) を取得
            var draw = GetDrawableArea(doc, sheet);
            FormworkDebugLog.Log(
                $"  [Sheet] drawArea U=[{draw.Min.U:F2}..{draw.Max.U:F2}] V=[{draw.Min.V:F2}..{draw.Max.V:F2}] ft");

            const double margin = 0.082;  // ≈25mm
            const double gap = 0.05;      // ≈15mm (集計表間の隙間)

            // 左上から
            double leftX = draw.Min.U + margin;
            double topY = draw.Max.V - margin;
            double rightX = draw.Max.U - margin;
            double currentTopY = topY;

            // サマリ集計表を先に配置（小さく、識別しやすい合計を上に）
            if (summaryScheduleId != null && summaryScheduleId != ElementId.InvalidElementId)
            {
                var sumRes = PlaceScheduleAt(doc, sheet, summaryScheduleId, leftX, currentTopY);
                if (sumRes.HasValue) currentTopY = sumRes.Value.bottomY - gap;
            }

            // メイン集計表 (ホスト + 各リンク) を中段に横並びで配置
            double schedulesBottomY = currentTopY;
            if (mainScheduleIds != null && mainScheduleIds.Count > 0)
            {
                double rowTopY = currentTopY;
                double rowLeftX = leftX;
                double rowMaxBottomY = rowTopY;

                foreach (var id in mainScheduleIds)
                {
                    if (id == null || id == ElementId.InvalidElementId) continue;

                    var res = PlaceScheduleAt(doc, sheet, id, rowLeftX, rowTopY);
                    if (!res.HasValue) continue;

                    double placedRight = res.Value.rightX;
                    double placedBottom = res.Value.bottomY;

                    // 配置後の右端がシートの右マージンを超えたら、その集計表を次の行に移動
                    if (placedRight > rightX && rowLeftX > leftX)
                    {
                        // 既存の集計表で行を確定し、新しい行を開始
                        try { doc.Delete(res.Value.instanceId); } catch { }
                        rowTopY = rowMaxBottomY - gap;
                        rowLeftX = leftX;
                        rowMaxBottomY = rowTopY;

                        var res2 = PlaceScheduleAt(doc, sheet, id, rowLeftX, rowTopY);
                        if (!res2.HasValue) continue;
                        placedRight = res2.Value.rightX;
                        placedBottom = res2.Value.bottomY;
                    }

                    if (placedBottom < rowMaxBottomY) rowMaxBottomY = placedBottom;
                    rowLeftX = placedRight + gap;
                }
                schedulesBottomY = rowMaxBottomY;
            }

            // 3Dビューを右下に配置
            if (view3DId != null && view3DId != ElementId.InvalidElementId)
            {
                PlaceViewportBottomRight(doc, sheet, view3DId, draw, margin);
            }

            FormworkDebugLog.Log(
                $"  [Sheet] created: '{sheet.SheetNumber} - {sheet.Name}' " +
                $"schedules={mainScheduleIds?.Count ?? 0}");
            return sheet.Id;
        }

        /// <summary>
        /// 既存シートで最も多く使われている TitleBlock の FamilySymbol Id を返す。
        /// 既存シートが無い場合はプロジェクト内の任意の TitleBlock を返し、それも無ければ null。
        /// </summary>
        private static ElementId FindMostUsedTitleBlock(Document doc)
        {
            var sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => !s.IsTemplate && !s.IsPlaceholder)
                .ToList();

            var counts = new Dictionary<ElementId, int>(new ElementIdComparer());
            foreach (var sh in sheets)
            {
                IEnumerable<FamilyInstance> tbInstances;
                try
                {
                    tbInstances = new FilteredElementCollector(doc, sh.Id)
                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>();
                }
                catch { continue; }

                foreach (var tb in tbInstances)
                {
                    var typeId = tb.GetTypeId();
                    if (typeId == null || typeId == ElementId.InvalidElementId) continue;
                    if (counts.ContainsKey(typeId)) counts[typeId]++;
                    else counts[typeId] = 1;
                }
            }

            if (counts.Count > 0)
            {
                var best = counts.OrderByDescending(kv => kv.Value).First();
                FormworkDebugLog.Log(
                    $"  [Sheet] most-used title block id={best.Key.IntValue()} count={best.Value}");
                return best.Key;
            }

            // フォールバック: 任意の TitleBlock タイプ
            var anySym = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsElementType()
                .FirstElementId();
            FormworkDebugLog.Log(
                $"  [Sheet] no existing sheets; fallback title block id={anySym?.IntValue()}");
            return anySym;
        }

        /// <summary>
        /// 同名シートが既存の場合に名前重複しないようサフィックスを付ける。
        /// </summary>
        private static string UniqueSheetName(Document doc, string baseName)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Select(s => s.Name)
                .ToHashSet();
            if (!existing.Contains(baseName)) return baseName;
            for (int i = 2; i < 1000; i++)
            {
                var candidate = $"{baseName} ({i})";
                if (!existing.Contains(candidate)) return candidate;
            }
            return baseName;
        }

        /// <summary>
        /// 「型枠-NNN」形式で未使用のシート番号を返す。
        /// </summary>
        private static string FindNextSheetNumber(Document doc)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Select(s => s.SheetNumber)
                .ToHashSet();
            for (int n = 1; n < 1000; n++)
            {
                var cand = $"型枠-{n:D3}";
                if (!existing.Contains(cand)) return cand;
            }
            return $"型枠-{DateTime.Now:HHmmss}";
        }

        /// <summary>
        /// 図枠の内側 (シート上の TitleBlock インスタンスの BoundingBox) を返す。
        /// 図枠が無い場合は sheet.Outline (描画可能領域全体) にフォールバック。
        /// </summary>
        private static BoundingBoxUV GetDrawableArea(Document doc, ViewSheet sheet)
        {
            try
            {
                var tb = new FilteredElementCollector(doc, sheet.Id)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .WhereElementIsNotElementType()
                    .FirstOrDefault() as FamilyInstance;
                if (tb != null)
                {
                    var bb = tb.get_BoundingBox(sheet);
                    if (bb != null && bb.Min != null && bb.Max != null)
                        return new BoundingBoxUV(bb.Min.X, bb.Min.Y, bb.Max.X, bb.Max.Y);
                }
            }
            catch { }

            try { return sheet.Outline; }
            catch
            {
                // 最終フォールバック: A1 横向き相当 (約 841×594mm)
                return new BoundingBoxUV(0, 0, 2.76, 1.95);
            }
        }

        /// <summary>
        /// 集計表をシート上の指定位置 (top-left 基準) に配置し、
        /// 配置インスタンスの ID と実 BoundingBox の右端 X / 下端 Y を返す。
        /// 失敗時は null。
        /// </summary>
        private static (ElementId instanceId, double rightX, double bottomY)? PlaceScheduleAt(
            Document doc, ViewSheet sheet, ElementId scheduleId,
            double leftX, double topY)
        {
            ScheduleSheetInstance inst;
            try
            {
                inst = ScheduleSheetInstance.Create(
                    doc, sheet.Id, scheduleId, new XYZ(leftX, topY, 0));
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Sheet] PlaceSchedule {scheduleId.IntValue()} EX: {ex.Message}");
                return null;
            }
            if (inst == null) return null;

            double right = leftX + 0.7;   // フォールバック: 約 213mm
            double bottom = topY - 0.328; // フォールバック: 約 100mm
            try
            {
                var bb = inst.get_BoundingBox(sheet);
                if (bb != null && bb.Min != null && bb.Max != null)
                {
                    right = bb.Max.X;
                    bottom = bb.Min.Y;
                }
            }
            catch { }
            FormworkDebugLog.Log(
                $"  [Sheet] schedule {scheduleId.IntValue()} placed at " +
                $"L={leftX:F2}/T={topY:F2} → R={right:F2}/B={bottom:F2}");
            return (inst.Id, right, bottom);
        }

        /// <summary>
        /// 3D ビューのビューポートを右下に配置する。
        /// 一旦仮配置してビューポートの実サイズを取得してから、
        /// 右下のマージン位置に合うよう移動する。
        /// </summary>
        private static void PlaceViewportBottomRight(
            Document doc, ViewSheet sheet, ElementId view3DId,
            BoundingBoxUV draw, double margin)
        {
            // 仮配置位置 (シート中央付近)
            double tmpX = (draw.Min.U + draw.Max.U) / 2.0;
            double tmpY = (draw.Min.V + draw.Max.V) / 2.0;

            Viewport vp;
            try
            {
                vp = Viewport.Create(doc, sheet.Id, view3DId, new XYZ(tmpX, tmpY, 0));
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Sheet] Viewport.Create EX: {ex.Message}");
                return;
            }
            if (vp == null) return;

            // ビューポートタイプを「タイトルなし」に切り替える
            ApplyNoTitleViewportType(doc, vp);

            // ビューポートの実アウトラインから幅・高さを取得
            double w = 0, h = 0;
            try
            {
                var ol = vp.GetBoxOutline();
                if (ol != null)
                {
                    w = ol.MaximumPoint.X - ol.MinimumPoint.X;
                    h = ol.MaximumPoint.Y - ol.MinimumPoint.Y;
                }
            }
            catch { }

            // 右下に合わせて移動 (center 基準)。下端から少し浮かせて配置 (≈30mm)。
            const double upOffset = 0.098;  // ≈30mm (100mm up から 70mm 下げた位置)
            double targetCx = draw.Max.U - margin - w / 2.0;
            double targetCy = draw.Min.V + margin + h / 2.0 + upOffset;
            try
            {
                var delta = new XYZ(targetCx - tmpX, targetCy - tmpY, 0);
                ElementTransformUtils.MoveElement(doc, vp.Id, delta);
                FormworkDebugLog.Log(
                    $"  [Sheet] viewport bottom-right cx={targetCx:F2} cy={targetCy:F2} w={w:F2} h={h:F2} upOffset={upOffset:F3}");
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Sheet] viewport move EX: {ex.Message}");
            }
        }

        /// <summary>
        /// ビューポートに「タイトルなし」タイプを適用する。
        /// 探索順:
        ///   1. cachedId が有効な場合はそれを直接使用 (キャッシュ)
        ///   2. 名前に「タイトルなし」「No Title」を含む既存タイプ
        ///   3. パラメータ VIEWPORT_ATTR_SHOW_LABEL=0 のタイプ
        ///   4. 既定タイプを複製して SHOW_LABEL=0 に設定 (フォールバック)
        ///   5. 同名が既に存在する場合は既存タイプを取得して再設定
        /// </summary>
        /// <returns>適用したタイプの ElementId (次回の cachedId として使用)。失敗時は InvalidElementId。</returns>
        private static ElementId ApplyNoTitleViewportType(Document doc, Viewport vp, ElementId cachedId = null)
        {
            if (doc == null || vp == null) return ElementId.InvalidElementId;
            try
            {
                // キャッシュ済みの ID が有効な場合はそれを直接使用
                if (cachedId != null && cachedId != ElementId.InvalidElementId)
                {
                    try
                    {
                        vp.ChangeTypeId(cachedId);
                        FormworkDebugLog.Log($"  [Sheet] ChangeTypeId (cached) OK → id={cachedId.IntValue()}");
                        return cachedId;
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"  [Sheet] ChangeTypeId (cached) EX: {ex.Message} → 再検索");
                    }
                }

                ElementId noTitleId = FindNoTitleViewportType(doc, out string foundReason);
                FormworkDebugLog.Log($"  [Sheet] FindNoTitleViewportType → id={noTitleId?.IntValue()} reason='{foundReason}'");

                // 既存に見つからなければ、現在タイプを複製して SHOW_LABEL=0 に設定
                if (noTitleId == ElementId.InvalidElementId)
                {
                    ElementId currentTypeId = null;
                    string currentTypeName = null;
                    try
                    {
                        currentTypeId = vp.GetTypeId();
                        var currentType = doc.GetElement(currentTypeId) as ElementType;
                        currentTypeName = currentType?.Name;
                        FormworkDebugLog.Log($"  [Sheet] currentType id={currentTypeId?.IntValue()} name='{currentTypeName}'");

                        if (currentType != null)
                        {
                            ElementType dup = null;
                            try
                            {
                                dup = currentType.Duplicate("Viewport_NoTitle_28Tools") as ElementType;
                                FormworkDebugLog.Log($"  [Sheet] Duplicate OK id={dup?.Id?.IntValue()}");
                            }
                            catch (Exception dupEx)
                            {
                                FormworkDebugLog.Log($"  [Sheet] Duplicate EX: {dupEx.Message} → 既存タイプを検索");
                                // 同名タイプが既存 → 直接検索
                                dup = new FilteredElementCollector(doc)
                                    .OfCategory(BuiltInCategory.OST_Viewports)
                                    .WhereElementIsElementType()
                                    .Cast<ElementType>()
                                    .FirstOrDefault(t => { try { return t.Name == "Viewport_NoTitle_28Tools"; } catch { return false; } });
                                FormworkDebugLog.Log($"  [Sheet] 既存タイプ直接検索 → id={dup?.Id?.IntValue()} name='{dup?.Name}'");
                            }

                            if (dup != null)
                            {
                                noTitleId = dup.Id;
                                // SHOW_LABEL=0 を確実に設定
                                var p = dup.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL);
                                FormworkDebugLog.Log($"  [Sheet] SHOW_LABEL on dup: exists={p != null} isReadOnly={p?.IsReadOnly} storageType={p?.StorageType} value={p?.AsInteger()}");
                                if (p != null && !p.IsReadOnly && p.StorageType == StorageType.Integer)
                                {
                                    p.Set(0);
                                    FormworkDebugLog.Log($"  [Sheet] SHOW_LABEL=0 設定完了 (dup)");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"  [Sheet] duplicate/create viewport type EX: {ex.Message}");
                    }
                }
                else
                {
                    // 既存タイプでも SHOW_LABEL=0 を確認・再設定
                    try
                    {
                        var existingType = doc.GetElement(noTitleId) as ElementType;
                        var p = existingType?.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL);
                        FormworkDebugLog.Log($"  [Sheet] SHOW_LABEL on existing: exists={p != null} isReadOnly={p?.IsReadOnly} value={p?.AsInteger()}");
                        if (p != null && !p.IsReadOnly && p.StorageType == StorageType.Integer && p.AsInteger() != 0)
                        {
                            p.Set(0);
                            FormworkDebugLog.Log($"  [Sheet] SHOW_LABEL=0 再設定完了 (existing)");
                        }
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"  [Sheet] SHOW_LABEL再設定 EX: {ex.Message}");
                    }
                }

                if (noTitleId != ElementId.InvalidElementId)
                {
                    try
                    {
                        vp.ChangeTypeId(noTitleId);
                        FormworkDebugLog.Log($"  [Sheet] ChangeTypeId OK → no-title id={noTitleId.IntValue()}");
                        return noTitleId;
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"  [Sheet] ChangeTypeId EX: {ex.Message}");
                    }
                }
                else
                {
                    FormworkDebugLog.Log($"  [Sheet] タイトルなしタイプ未設定 → ビューポートタイトルが表示される");
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Sheet] ApplyNoTitleViewportType EX: {ex.Message}");
            }
            return ElementId.InvalidElementId;
        }

        private static ElementId FindNoTitleViewportType(Document doc, out string reason)
        {
            reason = "not found";
            // FilteredElementCollector.OfCategory(OST_Viewports).WhereElementIsElementType() は
            // 環境によって 0 件を返すことがある。OfClass(typeof(ElementType)) で全タイプを取得し、
            // VIEWPORT_ATTR_SHOW_LABEL パラメータを持つものだけをビューポートタイプとして扱う。
            var types = new FilteredElementCollector(doc)
                .OfClass(typeof(ElementType))
                .Cast<ElementType>()
                .Where(t =>
                {
                    try { return t.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL) != null; }
                    catch { return false; }
                })
                .ToList();

            FormworkDebugLog.Log($"  [Sheet] viewport types found (OfClass): {types.Count}");
            foreach (var vt in types)
            {
                try
                {
                    var p = vt.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL);
                    FormworkDebugLog.Log($"  [Sheet]   type id={vt.Id.IntValue()} name='{vt.Name}' SHOW_LABEL={p?.AsInteger()}");
                }
                catch { }
            }

            // 1. VIEWPORT_ATTR_SHOW_LABEL = 0 のタイプを最優先
            foreach (var vt in types)
            {
                try
                {
                    var p = vt.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL);
                    if (p != null && p.StorageType == StorageType.Integer && p.AsInteger() == 0)
                    {
                        reason = $"SHOW_LABEL=0: '{vt.Name}'";
                        return vt.Id;
                    }
                }
                catch { }
            }

            // 2. 名前に「タイトルなし」「No Title」を含むタイプ
            foreach (var vt in types)
            {
                string name = null;
                try { name = vt.Name; } catch { }
                if (string.IsNullOrEmpty(name)) continue;
                if (name.Contains("タイトルなし") ||
                    name.IndexOf("No Title", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("NoTitle", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.Contains("無し") ||
                    name.Contains("無タイトル"))
                {
                    reason = $"name match: '{name}'";
                    return vt.Id;
                }
            }

            return ElementId.InvalidElementId;
        }

        private sealed class ElementIdComparer : IEqualityComparer<ElementId>
        {
            public bool Equals(ElementId x, ElementId y) =>
                x != null && y != null && x.IntValue() == y.IntValue();
            public int GetHashCode(ElementId obj) =>
                obj?.IntValue().GetHashCode() ?? 0;
        }
    }
}

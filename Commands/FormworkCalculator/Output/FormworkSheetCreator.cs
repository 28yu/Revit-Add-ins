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
    ///   - レイアウト: 集計表を左上から縦に並べ、3Dビューを右下に配置
    ///   - シート名: 「型枠数量集計」, シート番号: 自動 (型枠-NNN)
    /// </summary>
    internal static class FormworkSheetCreator
    {
        internal const string SheetName = "型枠数量集計";

        /// <summary>
        /// 集計シートを作成し、ビュー・集計表を配置する。
        /// </summary>
        internal static ElementId CreateSheet(
            Document doc,
            ElementId view3DId,
            ElementId mainScheduleId,
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

            // 配置可能領域 (図枠の内側) を取得
            var draw = GetDrawableArea(doc, sheet);
            FormworkDebugLog.Log(
                $"  [Sheet] drawArea U=[{draw.Min.U:F2}..{draw.Max.U:F2}] V=[{draw.Min.V:F2}..{draw.Max.V:F2}] ft");

            const double margin = 0.082;  // ≈25mm

            // 左上から集計表を縦に並べる
            double leftX = draw.Min.U + margin;
            double topY = draw.Max.V - margin;
            double currentTopY = topY;

            // サマリ集計表を先に配置（小さく、識別しやすい合計を上に）
            if (summaryScheduleId != null && summaryScheduleId != ElementId.InvalidElementId)
            {
                currentTopY = PlaceSchedule(doc, sheet, summaryScheduleId,
                    leftX, currentTopY, margin);
            }

            // メイン集計表を下に配置
            if (mainScheduleId != null && mainScheduleId != ElementId.InvalidElementId)
            {
                currentTopY = PlaceSchedule(doc, sheet, mainScheduleId,
                    leftX, currentTopY, margin);
            }

            // 3Dビューを右下に配置
            if (view3DId != null && view3DId != ElementId.InvalidElementId)
            {
                PlaceViewportBottomRight(doc, sheet, view3DId, draw, margin);
            }

            FormworkDebugLog.Log($"  [Sheet] created: '{sheet.SheetNumber} - {sheet.Name}'");
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
        /// 集計表をシート左上基準で配置し、配置後の下端 Y 座標を返す
        /// (次の集計表を下に積むため)。
        /// </summary>
        private static double PlaceSchedule(
            Document doc, ViewSheet sheet, ElementId scheduleId,
            double leftX, double topY, double gap)
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
                return topY;
            }
            if (inst == null) return topY;

            // 配置後の bb から下端を取得して次の Y を返す
            try
            {
                var bb = inst.get_BoundingBox(sheet);
                if (bb != null && bb.Min != null)
                {
                    double bottom = bb.Min.Y;
                    FormworkDebugLog.Log(
                        $"  [Sheet] schedule {scheduleId.IntValue()} placed top={topY:F2} bottom={bottom:F2}");
                    return bottom - gap;
                }
            }
            catch { }

            // bb 取得失敗時は概算で下げる (約 100mm)
            return topY - 0.328;
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
        ///   1. 名前に「タイトルなし」「No Title」を含む既存タイプ
        ///   2. パラメータ VIEWPORT_ATTR_SHOW_LABEL=0 のタイプ
        ///   3. 既定タイプを複製して SHOW_LABEL=0 に設定 (フォールバック)
        /// </summary>
        private static void ApplyNoTitleViewportType(Document doc, Viewport vp)
        {
            if (doc == null || vp == null) return;
            try
            {
                ElementId noTitleId = FindNoTitleViewportType(doc);

                // 既存に見つからなければ、現在タイプを複製して SHOW_LABEL=0 に設定
                if (noTitleId == ElementId.InvalidElementId)
                {
                    try
                    {
                        var currentTypeId = vp.GetTypeId();
                        var currentType = doc.GetElement(currentTypeId) as ElementType;
                        if (currentType != null)
                        {
                            var dup = currentType.Duplicate("Viewport_NoTitle_28Tools");
                            var p = dup.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL);
                            if (p != null && !p.IsReadOnly && p.StorageType == StorageType.Integer)
                            {
                                p.Set(0);
                                noTitleId = dup.Id;
                                FormworkDebugLog.Log(
                                    $"  [Sheet] duplicated viewport type → no-title id={noTitleId.IntValue()}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"  [Sheet] duplicate viewport type EX: {ex.Message}");
                    }
                }

                if (noTitleId != ElementId.InvalidElementId)
                {
                    try
                    {
                        vp.ChangeTypeId(noTitleId);
                        FormworkDebugLog.Log($"  [Sheet] viewport type → no-title id={noTitleId.IntValue()}");
                    }
                    catch (Exception ex)
                    {
                        FormworkDebugLog.Log($"  [Sheet] ChangeTypeId EX: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [Sheet] ApplyNoTitleViewportType EX: {ex.Message}");
            }
        }

        private static ElementId FindNoTitleViewportType(Document doc)
        {
            var types = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Viewports)
                .WhereElementIsElementType()
                .Cast<ElementType>()
                .ToList();

            // 1. 名前で「タイトルなし」「No Title」を含むタイプ
            foreach (var vt in types)
            {
                string name = null;
                try { name = vt.Name; } catch { }
                if (string.IsNullOrEmpty(name)) continue;
                if (name.Contains("タイトルなし") ||
                    name.IndexOf("No Title", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.Contains("無し") ||
                    name.Contains("無タイトル"))
                {
                    return vt.Id;
                }
            }

            // 2. VIEWPORT_ATTR_SHOW_LABEL = 0 のタイプ
            foreach (var vt in types)
            {
                try
                {
                    var p = vt.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL);
                    if (p != null && p.StorageType == StorageType.Integer && p.AsInteger() == 0)
                        return vt.Id;
                }
                catch { }
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

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Output
{
    /// <summary>
    /// 集計結果を Revit のテキスト TextNote ベースの表として作成する。
    /// ViewSchedule の自動生成は要素カテゴリに依存するため、本実装では製図ビュー + TextNote で表示する簡易版。
    /// </summary>
    internal static class ScheduleCreator
    {
        internal static ElementId CreateSummaryDraftingView(
            Document doc,
            FormworkResult result,
            FormworkSettings settings,
            ElementId textNoteTypeId = null)
        {
            var vftype = FindDraftingViewType(doc);
            if (vftype == null) return null;

            ViewDrafting view;
            try { view = ViewDrafting.Create(doc, vftype.Id); }
            catch { return null; }

            string viewName = $"型枠数量集計_{DateTime.Now:yyyyMMdd_HHmmss}";
            try { view.Name = viewName; } catch { }

            if (textNoteTypeId == null || textNoteTypeId == ElementId.InvalidElementId)
            {
                var tnt = new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .FirstOrDefault();
                if (tnt != null) textNoteTypeId = tnt.Id;
            }
            if (textNoteTypeId == null) return view.Id;

            var lines = BuildLines(result, settings);
            string content = string.Join("\n", lines);
            var opts = new TextNoteOptions(textNoteTypeId)
            {
                HorizontalAlignment = HorizontalTextAlignment.Left,
            };
            try
            {
                TextNote.Create(doc, view.Id, new XYZ(0, 0, 0), content, opts);
            }
            catch { }

            return view.Id;
        }

        private static ViewFamilyType FindDraftingViewType(Document doc)
        {
            foreach (var t in new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>())
            {
                if (t.ViewFamily == ViewFamily.Drafting) return t;
            }
            return null;
        }

        private static IEnumerable<string> BuildLines(FormworkResult r, FormworkSettings s)
        {
            yield return "型枠数量集計";
            yield return $"作成日時: {DateTime.Now:yyyy/MM/dd HH:mm:ss}";
            yield return $"対象要素数: {r.ProcessedElementCount}";
            yield return "";
            yield return "■ 総括";
            yield return $"  型枠面積合計: {r.TotalFormworkArea:F2} ㎡";
            yield return $"  控除面積合計: {r.TotalDeductedArea:F2} ㎡";
            yield return $"  傾斜面（計算対象外）: {r.InclinedFaceArea:F2} ㎡";
            yield return "";

            yield return "■ 部位別";
            yield return "  部位      要素数     型枠面積(㎡)   控除面積(㎡)";
            foreach (var c in r.CategoryResults)
            {
                yield return $"  {CategoryLabel(c.Category),-8}  {c.ElementCount,5}    {c.FormworkArea,10:F2}    {c.DeductedArea,10:F2}";
            }

            if (s.GroupByZone && r.ZoneResults.Count > 0)
            {
                yield return "";
                yield return "■ 工区別";
                foreach (var z in r.ZoneResults)
                    yield return $"  {z.Zone}: {z.FormworkArea:F2} ㎡ ({z.ElementCount}要素)";
            }

            if (s.GroupByFormworkType && r.TypeResults.Count > 0)
            {
                yield return "";
                yield return "■ 型枠種別";
                foreach (var t in r.TypeResults)
                    yield return $"  {t.FormworkType}: {t.FormworkArea:F2} ㎡ ({t.ElementCount}要素)";
            }

            if (r.Errors.Count > 0)
            {
                yield return "";
                yield return $"■ エラー・注記: {r.Errors.Count} 件";
            }
        }

        private static string CategoryLabel(CategoryGroup cg)
        {
            switch (cg)
            {
                case CategoryGroup.Column: return "柱";
                case CategoryGroup.Beam: return "梁";
                case CategoryGroup.Wall: return "壁";
                case CategoryGroup.Slab: return "スラブ";
                case CategoryGroup.Foundation: return "基礎";
                case CategoryGroup.Stairs: return "階段";
                default: return "その他";
            }
        }
    }
}

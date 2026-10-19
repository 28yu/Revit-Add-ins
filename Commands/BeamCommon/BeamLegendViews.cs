using System;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamCommon
{
    /// <summary>
    /// 梁天端・梁下端色分けの凡例を描くビューの取得・作成結果
    /// </summary>
    public class BeamLegendResult
    {
        public View View { get; set; }
        /// <summary>同名の凡例があり、上書きしない設定だったため作らなかった</summary>
        public bool Skipped { get; set; }
        /// <summary>モデルに凡例ビューが無いため、製図ビューで作った</summary>
        public bool UsedDraftingFallback { get; set; }
    }

    /// <summary>
    /// 凡例を描くビューを用意する。
    ///
    /// Revit API には凡例ビューを新規作成する手段が無いため、モデル内の既存の凡例ビューを
    /// 複製して作る。凡例ビューが1つも無いモデルでは、従来どおり製図ビューで作る。
    /// </summary>
    public static class BeamLegendViews
    {
        /// <summary>
        /// 凡例（凡例ビュー／旧版の製図ビュー）を名前の完全一致で探す。凡例ビューを優先する
        /// </summary>
        public static View Find(Document doc, string viewName)
        {
            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && v.Name == viewName &&
                    (v.ViewType == ViewType.Legend || v.ViewType == ViewType.DraftingView))
                .ToList();
            return views.FirstOrDefault(v => v.ViewType == ViewType.Legend) ?? views.FirstOrDefault();
        }

        /// <summary>
        /// 凡例を描くビューを用意する（トランザクション内で呼ぶ）。既存なら中身を消して再利用する
        /// </summary>
        public static BeamLegendResult Prepare(Document doc, string viewName, bool overwriteExisting)
        {
            View existing = Find(doc, viewName);
            if (existing != null && !overwriteExisting)
                return new BeamLegendResult { Skipped = true };

            View legendSource = FindLegendToDuplicate(doc);

            // 既存は中身を消して使い回す（シートに配置した凡例を残すため）。
            // 旧版の製図ビューは、凡例ビューを作れないモデルに限り使い回す
            if (existing != null &&
                (existing.ViewType == ViewType.Legend || legendSource == null))
            {
                ClearViewContents(doc, existing);
                return new BeamLegendResult
                {
                    View = existing,
                    UsedDraftingFallback = existing.ViewType != ViewType.Legend
                };
            }

            if (existing != null)
            {
                // 旧版の製図ビューの凡例。シートに置かれている場合があるので消さずに名前だけ変える
                try { existing.Name = viewName + "（旧）"; } catch (Exception) { }
            }

            View view = null;
            bool fallback = false;
            if (legendSource != null)
            {
                ElementId newId = legendSource.Duplicate(ViewDuplicateOption.Duplicate);
                view = doc.GetElement(newId) as View;
                if (view != null)
                    ClearViewContents(doc, view);
            }
            if (view == null)
            {
                ElementId draftingTypeId = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Drafting)?.Id;
                if (draftingTypeId == null)
                    return new BeamLegendResult();
                view = ViewDrafting.Create(doc, draftingTypeId);
                fallback = true;
            }

            try
            {
                view.Name = viewName;
            }
            catch (Exception)
            {
                view.Name = viewName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            }

            return new BeamLegendResult { View = view, UsedDraftingFallback = fallback };
        }

        private static View FindLegendToDuplicate(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .FirstOrDefault(v => !v.IsTemplate && v.ViewType == ViewType.Legend &&
                    v.CanViewBeDuplicated(ViewDuplicateOption.Duplicate));
        }

        /// <summary>
        /// ビューの中身（文字・塗潰し領域・凡例コンポーネントなど）だけを削除する
        /// </summary>
        private static void ClearViewContents(Document doc, View view)
        {
            var ids = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Id != view.Id && e.OwnerViewId == view.Id)
                .Select(e => e.Id)
                .ToList();
            foreach (var id in ids)
            {
                try { doc.Delete(id); } catch (Exception) { }
            }
        }
    }
}

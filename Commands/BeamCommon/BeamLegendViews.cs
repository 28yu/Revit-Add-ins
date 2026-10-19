using System;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Localization;

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
        /// <summary>同名の製図ビュー（旧版で作った凡例）を上書きした</summary>
        public bool ReusedDraftingView { get; set; }
        /// <summary>作成に失敗したときの理由</summary>
        public string Error { get; set; }
    }

    /// <summary>
    /// 凡例を描くビューを用意する。
    ///
    /// Revit API には凡例ビューを新規作成する手段が無いため、モデル内の既存の凡例ビューを
    /// 複製して作る。凡例ビューが1つも無いモデルでは、従来どおり製図ビューで作る。
    /// 同名の凡例があり上書きする場合は、そのビュー自体を描き直す（種類は変えない）。
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
        /// 凡例を描くビューを用意する（トランザクション内で呼ぶ）。
        /// 同名のビューがあれば、種類（凡例ビュー／製図ビュー）に関係なくそのビューの中身を消して描き直す
        /// （シートに配置した凡例をそのまま残すため）。無ければ凡例ビューを新しく作る
        /// </summary>
        public static BeamLegendResult Prepare(Document doc, string viewName, bool overwriteExisting)
        {
            View existing = Find(doc, viewName);
            if (existing != null)
            {
                if (!overwriteExisting)
                    return new BeamLegendResult { Skipped = true };

                ClearViewContents(doc, existing);
                if (!existing.IsValidObject)
                    return new BeamLegendResult { Error = Loc.S("Beam.LegendViewLost") };
                return new BeamLegendResult
                {
                    View = existing,
                    ReusedDraftingView = existing.ViewType == ViewType.DraftingView
                };
            }

            View view = null;
            bool fallback = false;
            View legendSource = FindLegendToDuplicate(doc);
            if (legendSource != null)
            {
                // 「複製」は詳細（文字・塗潰し領域など）を写さないので、空の凡例ビューができる
                ElementId newId = legendSource.Duplicate(ViewDuplicateOption.Duplicate);
                view = doc.GetElement(newId) as View;
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
        /// ビューの中身（文字・塗潰し領域・線・凡例コンポーネントなど描いたもの）だけを削除する。
        /// ⚠ ビューに属する要素をすべて消すと、ビューと一体の内部要素まで消えて
        ///   ビュー自体が削除されるため、描画要素の種類に絞って消す
        /// </summary>
        private static void ClearViewContents(Document doc, View view)
        {
            var ids = new FilteredElementCollector(doc)
                .WherePasses(new ElementOwnerViewFilter(view.Id))
                .WhereElementIsNotElementType()
                .Where(IsDrawnElement)
                .Select(e => e.Id)
                .ToList();
            foreach (var id in ids)
            {
                try { doc.Delete(id); } catch (Exception) { }
            }
        }

        private static bool IsDrawnElement(Element e)
        {
            if (e is TextNote || e is FilledRegion || e is CurveElement ||
                e is Dimension || e is IndependentTag || e is FamilyInstance)
                return true;
            Category cat = e.Category;
            return cat != null && cat.Id.Equals(new ElementId(BuiltInCategory.OST_LegendComponents));
        }
    }
}

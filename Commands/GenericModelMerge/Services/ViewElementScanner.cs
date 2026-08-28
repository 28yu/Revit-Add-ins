using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.GenericModelMerge.Models;

namespace Tools28.Commands.GenericModelMerge.Services
{
    /// <summary>
    /// アクティブ 3D ビューに実際に表示されている要素を走査し、カテゴリ別に集計する。
    ///
    /// FilteredElementCollector にビュー Id を渡すと、セクションボックス・ビューフィルタ・
    /// 要素の非表示設定が効いた「そのビューで見えている要素」だけが返る。
    /// そのため「ビュー内の全要素」の絞り込みはビュー側の設定にそのまま従う。
    /// </summary>
    internal static class ViewElementScanner
    {
        /// <summary>
        /// ビュー内の可視要素をモデルカテゴリ単位に集計する。要素数の多い順に並べて返す。
        /// </summary>
        public static List<MergeCategoryRow> ScanCategories(Document doc, View3D view)
        {
            var map = new Dictionary<int, MergeCategoryRow>();

            var collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType();

            foreach (Element e in collector)
            {
                if (!IsMergeCandidate(e)) continue;

                int key = e.Category.Id.IntValue();
                if (!map.TryGetValue(key, out var row))
                {
                    row = new MergeCategoryRow
                    {
                        CategoryId = e.Category.Id,
                        Name = e.Category.Name,
                    };
                    map[key] = row;
                }
                row.ElementIds.Add(e.Id);
            }

            return map.Values
                .OrderByDescending(r => r.ElementCount)
                .ThenBy(r => r.Name, StringComparer.CurrentCulture)
                .ToList();
        }

        /// <summary>
        /// 形状を読み取る対象になり得る要素か。
        /// 注釈・ビュー固有要素・カテゴリなし要素は形状(Solid)を持たないので除外する。
        /// </summary>
        private static bool IsMergeCandidate(Element e)
        {
            if (e == null) return false;
            if (e is ElementType) return false;
            if (e.ViewSpecific) return false;

            // リンクモデルは対象外。リンク内要素の形状は取得方法が異なり
            // (GetTotalTransform によるホスト座標への変換が必要)、本機能では扱わない。
            if (e is RevitLinkInstance) return false;

            Category cat = null;
            try { cat = e.Category; } catch { }
            if (cat == null) return false;
            if (cat.CategoryType != CategoryType.Model) return false;

            // 通り芯・レベル・参照面などは「モデル」カテゴリだが立体形状を持たない。
            // ここでは弾かず、形状取得時に Solid が 0 個なら自然に除外される。
            return true;
        }

        /// <summary>プロジェクト内の材質を名前順で列挙する。</summary>
        public static List<MaterialRow> ListMaterials(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Material))
                .Cast<Material>()
                .Select(m => new MaterialRow { Id = m.Id, Name = m.Name })
                .OrderBy(m => m.Name, StringComparer.CurrentCulture)
                .ToList();
        }
    }
}

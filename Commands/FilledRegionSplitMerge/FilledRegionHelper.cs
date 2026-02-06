using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FilledRegionSplitMerge
{
    /// <summary>
    /// 塗り潰し領域の分割/統合処理を行うヘルパークラス
    /// </summary>
    public class FilledRegionHelper
    {
        /// <summary>
        /// 選択された領域の分析結果
        /// </summary>
        public class SelectionAnalysis
        {
            public bool CanSplit { get; set; }
            public bool CanMerge { get; set; }
            public List<FilledRegion> FilledRegions { get; set; }
            public List<FilledRegionType> AvailableTypes { get; set; }
            public FilledRegionType DefaultType { get; set; }
        }

        /// <summary>
        /// 選択された領域を分析
        /// </summary>
        public static SelectionAnalysis AnalyzeSelection(Document doc, ICollection<ElementId> selectedIds)
        {
            var analysis = new SelectionAnalysis
            {
                CanSplit = false,
                CanMerge = false,
                FilledRegions = new List<FilledRegion>(),
                AvailableTypes = new List<FilledRegionType>()
            };

            if (selectedIds == null || selectedIds.Count == 0)
                return analysis;

            // FilledRegionのみをフィルタリング
            foreach (var id in selectedIds)
            {
                var element = doc.GetElement(id);
                if (element is FilledRegion fr)
                {
                    analysis.FilledRegions.Add(fr);
                }
            }

            if (analysis.FilledRegions.Count == 0)
                return analysis;

            // 分割可能かチェック（複数エリアを持つ領域があるか）
            foreach (var fr in analysis.FilledRegions)
            {
                var boundaries = fr.GetBoundaries();
                if (boundaries.Count > 1)
                {
                    analysis.CanSplit = true;
                    break;
                }
            }

            // 統合可能かチェック（複数の領域が選択されているか）
            if (analysis.FilledRegions.Count > 1)
            {
                analysis.CanMerge = true;
            }

            // 利用可能なパターンを取得
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType));

            foreach (FilledRegionType frt in collector)
            {
                analysis.AvailableTypes.Add(frt);
            }

            // デフォルトパターンを決定（選択された領域が同じパターンの場合、それを優先）
            if (analysis.FilledRegions.Count > 0)
            {
                var firstTypeId = analysis.FilledRegions[0].GetTypeId();
                if (analysis.FilledRegions.All(fr => fr.GetTypeId() == firstTypeId))
                {
                    analysis.DefaultType = doc.GetElement(firstTypeId) as FilledRegionType;
                }
                else if (analysis.AvailableTypes.Count > 0)
                {
                    analysis.DefaultType = analysis.AvailableTypes[0];
                }
            }

            return analysis;
        }

        /// <summary>
        /// 塗り潰し領域を分割
        /// </summary>
        /// <returns>作成された領域数</returns>
        public static int SplitFilledRegion(Document doc, FilledRegion filledRegion)
        {
            var boundaries = filledRegion.GetBoundaries();
            if (boundaries.Count <= 1)
            {
                return 0; // 分割不可
            }

            var typeId = filledRegion.GetTypeId();
            var view = doc.GetElement(filledRegion.OwnerViewId) as View;

            if (view == null)
            {
                throw new InvalidOperationException("ビューが見つかりません");
            }

            int createdCount = 0;

            // 各境界線に対して新しい領域を作成
            foreach (var curveLoop in boundaries)
            {
                var loops = new List<CurveLoop> { curveLoop };
                // TODO: Revit API調査が必要 - NewFilledRegionメソッドの正しい使用方法を確認
                // エラー: CS1061 'Document' に 'NewFilledRegion' の定義が含まれていない
                doc.Create.NewFilledRegion(view, typeId, loops);
                createdCount++;
            }

            // 元の領域を削除
            doc.Delete(filledRegion.Id);

            return createdCount;
        }

        /// <summary>
        /// 複数の塗り潰し領域を統合
        /// </summary>
        /// <returns>統合された領域数</returns>
        public static int MergeFilledRegions(Document doc, List<FilledRegion> filledRegions, ElementId newTypeId)
        {
            if (filledRegions == null || filledRegions.Count < 2)
            {
                return 0; // 統合不可
            }

            // 全ての境界線を取得
            var allBoundaries = new List<CurveLoop>();
            View view = null;

            foreach (var fr in filledRegions)
            {
                var boundaries = fr.GetBoundaries();
                allBoundaries.AddRange(boundaries);

                if (view == null)
                {
                    view = doc.GetElement(fr.OwnerViewId) as View;
                }
            }

            if (view == null)
            {
                throw new InvalidOperationException("ビューが見つかりません");
            }

            // 新しい統合領域を作成
            // TODO: Revit API調査が必要 - NewFilledRegionメソッドの正しい使用方法を確認
            doc.Create.NewFilledRegion(view, newTypeId, allBoundaries);

            // 元の領域を全て削除
            var idsToDelete = filledRegions.Select(fr => fr.Id).ToList();
            doc.Delete(idsToDelete);

            return filledRegions.Count;
        }

        /// <summary>
        /// FilledRegionTypeの名前を取得
        /// </summary>
        public static string GetTypeName(FilledRegionType type)
        {
            if (type == null)
                return string.Empty;

            return type.Name;
        }
    }
}

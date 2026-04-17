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
                FilledRegion.Create(doc, typeId, view.Id, loops);
                createdCount++;
            }

            // 元の領域を削除
            doc.Delete(filledRegion.Id);

            return createdCount;
        }

        /// <summary>
        /// 複数の塗り潰し領域を統合（重なり合う領域にも対応）
        /// </summary>
        /// <returns>統合された領域数</returns>
        public static int MergeFilledRegions(Document doc, List<FilledRegion> filledRegions, ElementId newTypeId)
        {
            if (filledRegions == null || filledRegions.Count < 2)
            {
                return 0; // 統合不可
            }

            // 各領域の境界ループを収集
            var regionLoops = new List<IList<CurveLoop>>();
            View view = null;

            foreach (var fr in filledRegions)
            {
                var boundaries = fr.GetBoundaries();
                if (boundaries != null && boundaries.Count > 0)
                    regionLoops.Add(boundaries);

                if (view == null)
                    view = doc.GetElement(fr.OwnerViewId) as View;
            }

            if (view == null)
            {
                throw new InvalidOperationException("ビューが見つかりません");
            }

            if (regionLoops.Count == 0)
                return 0;

            // 重なり合う領域を考慮してブーリアン和で統合
            var mergedLoops = UnionRegionLoops(regionLoops);

            // Union に失敗した場合はフォールバック（単純連結）
            if (mergedLoops == null || mergedLoops.Count == 0)
            {
                mergedLoops = new List<CurveLoop>();
                foreach (var loops in regionLoops)
                    mergedLoops.AddRange(loops);
            }

            FilledRegion.Create(doc, newTypeId, view.Id, mergedLoops);

            // 元の領域を全て削除
            var idsToDelete = filledRegions.Select(fr => fr.Id).ToList();
            doc.Delete(idsToDelete);

            return filledRegions.Count;
        }

        /// <summary>
        /// 複数の CurveLoop 群をブーリアン和で統合する
        /// （重なり合う領域を1つの輪郭に統合、離れている領域は複数ループとして残る）
        /// </summary>
        private static List<CurveLoop> UnionRegionLoops(List<IList<CurveLoop>> regionLoops)
        {
            // 共通の基準平面（最初の外形ループから法線を決定）
            XYZ normal = null;
            foreach (var loops in regionLoops)
            {
                if (loops.Count == 0) continue;
                try
                {
                    var plane = loops[0].GetPlane();
                    if (plane != null)
                    {
                        normal = plane.Normal.Normalize();
                        break;
                    }
                }
                catch { }
            }

            if (normal == null)
                return null;

            const double thickness = 1.0; // 1 feet の薄板として押し出し

            Solid unionSolid = null;

            foreach (var loops in regionLoops)
            {
                Solid solid;
                try
                {
                    solid = GeometryCreationUtilities.CreateExtrusionGeometry(loops, normal, thickness);
                }
                catch
                {
                    continue;
                }

                if (solid == null || solid.Volume < 1e-9)
                    continue;

                if (unionSolid == null)
                {
                    unionSolid = solid;
                    continue;
                }

                try
                {
                    var combined = BooleanOperationsUtils.ExecuteBooleanOperation(
                        unionSolid, solid, BooleanOperationsType.Union);
                    if (combined != null && combined.Volume > 1e-9)
                        unionSolid = combined;
                }
                catch
                {
                    // 和集合に失敗した場合はその領域を無視して続行
                }
            }

            if (unionSolid == null)
                return null;

            // 押し出し方向の"上面"となる平面（法線が +normal と一致する PlanarFace）を取得
            PlanarFace topFace = null;
            foreach (Face face in unionSolid.Faces)
            {
                if (face is PlanarFace pf
                    && pf.FaceNormal.IsAlmostEqualTo(normal))
                {
                    topFace = pf;
                    break;
                }
            }

            if (topFace == null)
                return null;

            try
            {
                var rawLoops = topFace.GetEdgesAsCurveLoops();
                if (rawLoops == null || rawLoops.Count == 0)
                    return null;

                // 上面は押し出し方向へ thickness ぶんオフセットしているので元平面へ戻す
                var back = Transform.CreateTranslation(-thickness * normal);
                var result = new List<CurveLoop>();
                foreach (var loop in rawLoops)
                {
                    var moved = new CurveLoop();
                    foreach (Curve c in loop)
                        moved.Append(c.CreateTransformed(back));
                    result.Add(moved);
                }
                return result;
            }
            catch
            {
                return null;
            }
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

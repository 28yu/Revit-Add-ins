using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Localization;

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

            // 分割可能かチェック（離れた複数のエリアを持つ領域があるか）
            // 穴のループは独立したエリアではないため、境界ループ数ではなくエリア数で判定する
            foreach (var fr in analysis.FilledRegions)
            {
                var boundaries = fr.GetBoundaries();
                if (boundaries.Count > 1)
                {
                    var islands = GetIslands(boundaries);
                    if (islands != null && islands.Count > 1)
                    {
                        analysis.CanSplit = true;
                        break;
                    }
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
                throw new InvalidOperationException(Loc.S("Common.ViewNotFound"));
            }

            // 外形ループとその穴をまとめて1エリアとする（穴を別の領域にしない）
            var islands = GetIslands(boundaries);
            if (islands == null || islands.Count <= 1)
            {
                return 0; // 分割不可（穴があるだけの1エリア、または形状を解析できない）
            }

            int createdCount = 0;

            foreach (var islandLoops in islands)
            {
                FilledRegion.Create(doc, typeId, view.Id, islandLoops);
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
                throw new InvalidOperationException(Loc.S("Common.ViewNotFound"));
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

        // 領域の形状を薄い板として押し出して扱うときの厚さ（feet）
        private const double ExtrusionThickness = 1.0;

        /// <summary>
        /// 複数の CurveLoop 群をブーリアン和で統合する
        /// （重なり合う領域を1つの輪郭に統合、離れている領域は複数ループとして残る）
        /// 1つでも押し出し・和集合に失敗した領域があれば null を返す
        /// （その領域を黙って捨てると、元の領域を削除したときに消えてしまうため）
        /// </summary>
        private static List<CurveLoop> UnionRegionLoops(List<IList<CurveLoop>> regionLoops)
        {
            XYZ normal = GetNormal(regionLoops);
            if (normal == null)
                return null;

            Solid unionSolid = null;

            foreach (var loops in regionLoops)
            {
                Solid solid = Extrude(loops, normal);
                if (solid == null)
                    return null;

                if (unionSolid == null)
                {
                    unionSolid = solid;
                    continue;
                }

                try
                {
                    var combined = BooleanOperationsUtils.ExecuteBooleanOperation(
                        unionSolid, solid, BooleanOperationsType.Union);
                    // 和集合は元のどちらよりも小さくならない。小さければ失敗とみなす
                    double minVolume = Math.Max(unionSolid.Volume, solid.Volume) * 0.999;
                    if (combined == null || combined.Volume < minVolume)
                        return null;
                    unionSolid = combined;
                }
                catch
                {
                    return null;
                }
            }

            return GetTopFaceLoops(unionSolid, normal);
        }

        /// <summary>
        /// 境界ループを、離れたエリアごと（外形ループ＋その穴）にまとめる。
        /// 形状を解析できない場合は null を返す
        /// </summary>
        private static List<List<CurveLoop>> GetIslands(IList<CurveLoop> boundaries)
        {
            XYZ normal = GetNormal(new List<IList<CurveLoop>> { boundaries });
            if (normal == null)
                return null;

            Solid solid = Extrude(boundaries, normal);
            if (solid == null)
                return null;

            try
            {
                var islands = new List<List<CurveLoop>>();
                foreach (Solid part in SolidUtils.SplitVolumes(solid))
                {
                    var loops = GetTopFaceLoops(part, normal);
                    if (loops == null)
                        return null;
                    islands.Add(loops);
                }

                // 取りこぼしが無いこと（ループ数の合計が元と一致）を確認する
                if (islands.Sum(i => i.Count) != boundaries.Count)
                    return null;

                return islands;
            }
            catch
            {
                return null;
            }
        }

        // 共通の基準平面の法線（最初に取得できたループの平面から決定）
        private static XYZ GetNormal(List<IList<CurveLoop>> regionLoops)
        {
            foreach (var loops in regionLoops)
            {
                if (loops.Count == 0) continue;
                try
                {
                    var plane = loops[0].GetPlane();
                    if (plane != null)
                        return plane.Normal.Normalize();
                }
                catch { }
            }
            return null;
        }

        private static Solid Extrude(IList<CurveLoop> loops, XYZ normal)
        {
            try
            {
                var solid = GeometryCreationUtilities.CreateExtrusionGeometry(loops, normal, ExtrusionThickness);
                if (solid == null || solid.Volume < 1e-9)
                    return null;
                return solid;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 押し出したソリッドの"上面"（法線が +normal の PlanarFace）すべてのループを、
        /// 元の平面に戻して返す。離れたエリアは上面が複数になるため、すべて集める。
        /// 上面の合計面積が体積/厚さと一致しない（取りこぼしがある）場合は null を返す
        /// </summary>
        private static List<CurveLoop> GetTopFaceLoops(Solid solid, XYZ normal)
        {
            if (solid == null)
                return null;

            try
            {
                var back = Transform.CreateTranslation(-ExtrusionThickness * normal);
                var result = new List<CurveLoop>();
                double topArea = 0;

                foreach (Face face in solid.Faces)
                {
                    if (!(face is PlanarFace pf) || !pf.FaceNormal.IsAlmostEqualTo(normal))
                        continue;

                    topArea += pf.Area;
                    foreach (var loop in pf.GetEdgesAsCurveLoops())
                    {
                        var moved = new CurveLoop();
                        foreach (Curve c in loop)
                            moved.Append(c.CreateTransformed(back));
                        result.Add(moved);
                    }
                }

                if (result.Count == 0)
                    return null;

                double expectedArea = solid.Volume / ExtrusionThickness;
                if (Math.Abs(topArea - expectedArea) > expectedArea * 0.001)
                    return null;

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

using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Engine;
using Tools28.Commands.GenericModelMerge.Models;

namespace Tools28.Commands.GenericModelMerge.Services
{
    /// <summary>形状読み取り＋結合の結果。</summary>
    internal class MergeResult
    {
        /// <summary>生成する一般モデルに入れる形状。</summary>
        public List<Solid> Solids { get; } = new List<Solid>();

        /// <summary>形状が取得できた元要素の Id（非表示処理の対象）。</summary>
        public List<ElementId> SourceElementIds { get; } = new List<ElementId>();

        /// <summary>形状(Solid)を持たず対象外になった要素数（通り芯・レベル・線分など）。</summary>
        public int NoGeometryCount { get; set; }

        /// <summary>結合 (Boolean) に失敗し、単独の形状として残した個数。</summary>
        public int UnionFailedCount { get; set; }

        /// <summary>形状がすべて切断ボックスの外側にあり、対象外にした要素数。</summary>
        public int OutsideSectionBoxCount { get; set; }

        /// <summary>切断ボックスの外側を切り取った形状の数。</summary>
        public int ClippedCount { get; set; }

        /// <summary>切断ボックスで切り取れず、元の形のまま入れた形状の数。</summary>
        public int ClipFailedCount { get; set; }
    }

    /// <summary>
    /// 複数要素から Solid を読み取り、指定された方法で「一つのかたまり」にまとめる。
    ///
    /// Solid の取得は FormworkCalculator の <see cref="SolidUnionProcessor"/> を再利用する
    /// （GeometryInstance の展開・微小ソリッドの除外まで実績のある実装のため）。
    /// </summary>
    internal static class SolidMerger
    {
        private const double MinSolidVolume = 1e-6;

        /// <summary>接触判定の許容値（フィート）。約 0.3 mm。</summary>
        private const double TouchTolerance = 0.001;

        /// <param name="clipper">切断ボックスが有効なビューなら、その外側を切り取る。null なら切らない。</param>
        public static MergeResult Build(Document doc, IList<ElementId> elementIds, MergeCombineMode mode,
            SectionBoxClipper clipper)
        {
            var result = new MergeResult();

            // 要素ごとの Solid 群と外形範囲を先に集める
            var perElement = new List<List<Solid>>();
            var boxes = new List<BoundingBoxXYZ>();

            foreach (var id in elementIds)
            {
                Element e = doc.GetElement(id);
                if (e == null) continue;

                List<Solid> solids;
                try { solids = SolidUnionProcessor.GetSolids(e); }
                catch { solids = new List<Solid>(); }

                if (solids.Count == 0)
                {
                    result.NoGeometryCount++;
                    continue;
                }

                BoundingBoxXYZ bb = null;
                if (clipper != null)
                {
                    solids = clipper.Clip(solids);
                    if (solids.Count == 0)
                    {
                        // 切断ボックスの面に触れているだけで、中身はすべて外側
                        result.OutsideSectionBoxCount++;
                        continue;
                    }
                    // 切り取った後の形で「くっついているか」を判定する
                    bb = BoundsOf(solids);
                }
                else
                {
                    try { bb = e.get_BoundingBox(null); } catch { }
                }

                perElement.Add(solids);
                boxes.Add(bb);
                result.SourceElementIds.Add(id);
            }

            if (clipper != null)
            {
                result.ClippedCount = clipper.ClippedCount;
                result.ClipFailedCount = clipper.FailedCount;
            }

            switch (mode)
            {
                case MergeCombineMode.UnionAll:
                    {
                        var all = new List<Solid>();
                        foreach (var s in perElement) all.AddRange(s);
                        AddUnioned(all, result);
                        break;
                    }
                case MergeCombineMode.UnionTouching:
                    {
                        foreach (var group in GroupByTouching(perElement, boxes))
                            AddUnioned(group, result);
                        break;
                    }
                default: // KeepShapes: 形状はそのまま、要素としてだけ 1 つにまとめる
                    {
                        foreach (var s in perElement) result.Solids.AddRange(s);
                        break;
                    }
            }

            return result;
        }

        /// <summary>
        /// 1 グループ分の Solid を Union して結果へ追加する。
        /// Boolean が失敗した Solid は捨てずに単独の形状として残す
        /// （黙って消えると「形状が欠けた一般モデル」ができてしまうため）。
        /// </summary>
        private static void AddUnioned(List<Solid> solids, MergeResult result)
        {
            Solid acc = null;
            var leftovers = new List<Solid>();

            foreach (var s in solids)
            {
                if (s == null || s.Volume <= MinSolidVolume) continue;
                if (acc == null) { acc = s; continue; }

                Solid merged = null;
                try
                {
                    merged = BooleanOperationsUtils.ExecuteBooleanOperation(
                        acc, s, BooleanOperationsType.Union);
                }
                catch { merged = null; }

                if (merged != null && merged.Volume > MinSolidVolume)
                {
                    acc = merged;
                }
                else
                {
                    leftovers.Add(s);
                    result.UnionFailedCount++;
                }
            }

            if (acc != null) result.Solids.Add(acc);
            result.Solids.AddRange(leftovers);
        }

        /// <summary>
        /// 外形範囲が接触・交差している要素同士を Union-Find で 1 グループにまとめる。
        /// A-C が接し C-B が接するなら A・B・C は同じグループになる（推移的に連結する）。
        /// </summary>
        private static List<List<Solid>> GroupByTouching(
            List<List<Solid>> perElement, List<BoundingBoxXYZ> boxes)
        {
            int n = perElement.Count;
            var parent = new int[n];
            for (int i = 0; i < n; i++) parent[i] = i;

            for (int i = 0; i < n; i++)
            {
                for (int j = i + 1; j < n; j++)
                {
                    if (BoxesOverlap(boxes[i], boxes[j]))
                        Unite(parent, i, j);
                }
            }

            var groups = new Dictionary<int, List<Solid>>();
            for (int i = 0; i < n; i++)
            {
                int root = Find(parent, i);
                if (!groups.TryGetValue(root, out var list))
                {
                    list = new List<Solid>();
                    groups[root] = list;
                }
                list.AddRange(perElement[i]);
            }
            return new List<List<Solid>>(groups.Values);
        }

        /// <summary>複数の Solid をまとめて囲む外形範囲（ワールド座標）。</summary>
        private static BoundingBoxXYZ BoundsOf(List<Solid> solids)
        {
            XYZ min = null, max = null;
            foreach (var s in solids)
            {
                BoundingBoxXYZ bb;
                try { bb = s.GetBoundingBox(); } catch { continue; }
                if (bb == null) continue;
                Transform t = bb.Transform ?? Transform.Identity;
                foreach (double x in new[] { bb.Min.X, bb.Max.X })
                foreach (double y in new[] { bb.Min.Y, bb.Max.Y })
                foreach (double z in new[] { bb.Min.Z, bb.Max.Z })
                {
                    XYZ p = t.OfPoint(new XYZ(x, y, z));
                    min = min == null ? p : new XYZ(Math.Min(min.X, p.X), Math.Min(min.Y, p.Y), Math.Min(min.Z, p.Z));
                    max = max == null ? p : new XYZ(Math.Max(max.X, p.X), Math.Max(max.Y, p.Y), Math.Max(max.Z, p.Z));
                }
            }
            if (min == null) return null;
            return new BoundingBoxXYZ { Min = min, Max = max };
        }

        private static int Find(int[] parent, int i)
        {
            while (parent[i] != i)
            {
                parent[i] = parent[parent[i]];
                i = parent[i];
            }
            return i;
        }

        private static void Unite(int[] parent, int a, int b)
        {
            int ra = Find(parent, a);
            int rb = Find(parent, b);
            if (ra != rb) parent[rb] = ra;
        }

        private static bool BoxesOverlap(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            // 外形が取れなかった要素は、どことも接していない扱いにする
            if (a == null || b == null) return false;
            const double t = TouchTolerance;
            return a.Min.X - t <= b.Max.X && b.Min.X - t <= a.Max.X
                && a.Min.Y - t <= b.Max.Y && b.Min.Y - t <= a.Max.Y
                && a.Min.Z - t <= b.Max.Z && b.Min.Z - t <= a.Max.Z;
        }
    }
}

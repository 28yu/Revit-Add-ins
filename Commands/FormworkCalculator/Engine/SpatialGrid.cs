using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 要素の BoundingBox を 3D 格子セルに登録し、接触候補ペアを効率的に列挙する。
    ///
    /// 【目的】
    ///   大規模モデル(500+ 要素) で O(N²) の全ペア走査を避ける。要素が存在する
    ///   セルと隣接セル (3×3×3=27個) にある要素だけを候補として返す。
    ///
    /// 【セルサイズの自動算出】
    ///   全要素 BB の対角線長の中央値 × 2 を採用する。
    ///   - 典型要素 2個分のサイズなら「ほぼ接触している要素」は隣接セル内に収まる
    ///   - 大きすぎる/小さすぎる要素の外れ値の影響を中央値で除去する
    /// </summary>
    internal class SpatialGrid
    {
        private readonly Dictionary<(int, int, int), List<int>> _cells = new Dictionary<(int, int, int), List<int>>();
        private readonly Dictionary<int, BoundingBoxXYZ> _bboxes = new Dictionary<int, BoundingBoxXYZ>();
        private readonly double _cellSize;
        private readonly double _tolFeet;

        internal double CellSize => _cellSize;

        private SpatialGrid(double cellSize, double tolFeet)
        {
            _cellSize = cellSize;
            _tolFeet = tolFeet;
        }

        internal static SpatialGrid Build(
            IEnumerable<ContactFaceDetector.ElementFacesContext> contexts, double tolFeet)
        {
            double cellSize = ComputeCellSize(contexts);
            var grid = new SpatialGrid(cellSize, tolFeet);

            foreach (var ctx in contexts)
            {
                if (ctx.BB == null) continue;
                grid._bboxes[ctx.ElementId] = ctx.BB;

                var (imin, jmin, kmin) = grid.CellCoord(ctx.BB.Min);
                var (imax, jmax, kmax) = grid.CellCoord(ctx.BB.Max);

                for (int i = imin; i <= imax; i++)
                    for (int j = jmin; j <= jmax; j++)
                        for (int k = kmin; k <= kmax; k++)
                        {
                            var key = (i, j, k);
                            if (!grid._cells.TryGetValue(key, out var list))
                            {
                                list = new List<int>();
                                grid._cells[key] = list;
                            }
                            list.Add(ctx.ElementId);
                        }
            }
            return grid;
        }

        /// <summary>
        /// 指定要素と BBox が重なる候補要素 ID を列挙する。自身は含まない。
        /// </summary>
        internal IEnumerable<int> GetCandidates(int elementId)
        {
            if (!_bboxes.TryGetValue(elementId, out var bb)) yield break;

            var (imin, jmin, kmin) = CellCoord(bb.Min);
            var (imax, jmax, kmax) = CellCoord(bb.Max);

            var seen = new HashSet<int>();
            for (int i = imin; i <= imax; i++)
                for (int j = jmin; j <= jmax; j++)
                    for (int k = kmin; k <= kmax; k++)
                    {
                        if (!_cells.TryGetValue((i, j, k), out var list)) continue;
                        foreach (var id in list)
                        {
                            if (id == elementId) continue;
                            if (!seen.Add(id)) continue;
                            if (!_bboxes.TryGetValue(id, out var other)) continue;
                            if (BBoxesOverlap(bb, other, _tolFeet))
                                yield return id;
                        }
                    }
        }

        private (int, int, int) CellCoord(XYZ p)
        {
            return (
                (int)Math.Floor(p.X / _cellSize),
                (int)Math.Floor(p.Y / _cellSize),
                (int)Math.Floor(p.Z / _cellSize)
            );
        }

        private static bool BBoxesOverlap(BoundingBoxXYZ a, BoundingBoxXYZ b, double tol)
        {
            return a.Min.X - tol <= b.Max.X && b.Min.X - tol <= a.Max.X
                && a.Min.Y - tol <= b.Max.Y && b.Min.Y - tol <= a.Max.Y
                && a.Min.Z - tol <= b.Max.Z && b.Min.Z - tol <= a.Max.Z;
        }

        private static double ComputeCellSize(IEnumerable<ContactFaceDetector.ElementFacesContext> contexts)
        {
            var diagonals = new List<double>();
            foreach (var ctx in contexts)
            {
                if (ctx.BB == null) continue;
                var d = ctx.BB.Max - ctx.BB.Min;
                double len = Math.Sqrt(d.X * d.X + d.Y * d.Y + d.Z * d.Z);
                if (len > 1e-6) diagonals.Add(len);
            }
            if (diagonals.Count == 0) return 10.0;   // 妥当なデフォルト (約3m)
            diagonals.Sort();
            double median = diagonals[diagonals.Count / 2];
            double cell = median * 2.0;
            // 極端に小さい/大きいセルサイズを防ぐ
            if (cell < 1.0) cell = 1.0;         // 最小 1ft (≈ 0.3m)
            if (cell > 200.0) cell = 200.0;     // 最大 200ft (≈ 60m)
            return cell;
        }
    }
}

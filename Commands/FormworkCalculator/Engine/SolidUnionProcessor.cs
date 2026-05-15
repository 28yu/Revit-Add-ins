using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    internal static class SolidUnionProcessor
    {
        private const double MinSolidVolume = 1e-6;

        internal static List<Solid> GetSolids(Element elem)
        {
            return GetSolids(elem, null);
        }

        /// <summary>
        /// 要素の Solid を取得し、transform が指定されていれば各 Solid に適用する。
        /// リンク要素の場合は GetTotalTransform を渡してホスト座標系に変換する。
        ///
        /// WallSweep フォールバック: 通常の Options で solid が 0 個だった場合、
        /// IncludeNonVisibleObjects=true で再試行する。リンクドキュメント内の
        /// WallSweep は「非表示要素」扱いされて geometry が空で返ることが多いため
        /// (Revit API の既知の挙動)。WallSweep 以外には適用しない。
        ///
        /// skipWallSweepRetry=true: BIM360ワークシェアリングのリンク要素では
        /// IncludeNonVisibleObjects=true の再取得が非常に遅い場合があるためスキップ。
        /// </summary>
        internal static List<Solid> GetSolids(Element elem, Transform transform,
            bool skipWallSweepRetry = false)
        {
            var result = new List<Solid>();
            if (elem == null) return result;

            Options opt = new Options
            {
                ComputeReferences = false,
                IncludeNonVisibleObjects = false,
                DetailLevel = ViewDetailLevel.Fine,
            };

            GeometryElement geom;
            try { geom = elem.get_Geometry(opt); }
            catch { geom = null; }
            if (geom != null) CollectSolidsRecursive(geom, result);

            // WallSweep フォールバック: 空だったら IncludeNonVisibleObjects=true で再取得
            // リンク要素の場合はスキップ (BIM360ワークシェアリングで非常に遅くなる)
            if (!skipWallSweepRetry && result.Count == 0 && elem is WallSweep)
            {
                Options opt2 = new Options
                {
                    ComputeReferences = false,
                    IncludeNonVisibleObjects = true,
                    DetailLevel = ViewDetailLevel.Fine,
                };
                GeometryElement geom2;
                try { geom2 = elem.get_Geometry(opt2); }
                catch { geom2 = null; }
                if (geom2 != null) CollectSolidsRecursive(geom2, result);
                if (result.Count > 0 && FormworkDebugLog.Enabled)
                {
                    FormworkDebugLog.Log(
                        $"  [LinkSweepDiag] WallSweep fallback E{elem.Id.IntValue()} " +
                        $"IncludeNonVisibleObjects=true → recovered {result.Count} solids");
                }
            }

            // transform が Identity でなければ各 Solid を変換
            if (transform != null && !transform.IsIdentity)
            {
                var transformed = new List<Solid>();
                foreach (var s in result)
                {
                    if (s == null || s.Volume <= MinSolidVolume) continue;
                    try
                    {
                        var ts = SolidUtils.CreateTransformed(s, transform);
                        if (ts != null && ts.Volume > MinSolidVolume) transformed.Add(ts);
                    }
                    catch { }
                }
                result = transformed;
            }
            return result;
        }

        private static void CollectSolidsRecursive(GeometryElement geom, List<Solid> output)
        {
            foreach (GeometryObject obj in geom)
            {
                if (obj is Solid s)
                {
                    if (s.Volume > MinSolidVolume)
                        output.Add(s);
                }
                else if (obj is GeometryInstance gi)
                {
                    var inst = gi.GetInstanceGeometry();
                    if (inst != null)
                        CollectSolidsRecursive(inst, output);
                }
            }
        }

        internal static Solid Union(IEnumerable<Solid> solids)
        {
            Solid acc = null;
            foreach (var s in solids)
            {
                if (s == null || s.Volume <= MinSolidVolume) continue;
                if (acc == null)
                {
                    acc = s;
                    continue;
                }
                try
                {
                    var r = BooleanOperationsUtils.ExecuteBooleanOperation(
                        acc, s, BooleanOperationsType.Union);
                    if (r != null && r.Volume > MinSolidVolume)
                        acc = r;
                }
                catch { }
            }
            return acc;
        }

        internal static List<Solid> UnionByProximity(List<Element> elements, double tol = 0.1)
        {
            var groups = new List<List<Solid>>();
            var groupBbs = new List<BoundingBoxXYZ>();

            foreach (var e in elements)
            {
                var solids = GetSolids(e);
                if (solids.Count == 0) continue;

                BoundingBoxXYZ eb = null;
                try { eb = e.get_BoundingBox(null); } catch { }
                if (eb == null) continue;

                int idx = -1;
                for (int i = 0; i < groupBbs.Count; i++)
                {
                    if (BoxesOverlap(groupBbs[i], eb, tol))
                    {
                        idx = i;
                        break;
                    }
                }
                if (idx < 0)
                {
                    groups.Add(new List<Solid>(solids));
                    groupBbs.Add(CloneBox(eb));
                }
                else
                {
                    groups[idx].AddRange(solids);
                    groupBbs[idx] = MergeBox(groupBbs[idx], eb);
                }
            }

            var result = new List<Solid>();
            foreach (var g in groups)
            {
                var u = Union(g);
                if (u != null) result.Add(u);
                else result.AddRange(g);
            }
            return result;
        }

        private static bool BoxesOverlap(BoundingBoxXYZ a, BoundingBoxXYZ b, double tol)
        {
            if (a == null || b == null) return false;
            return a.Min.X - tol <= b.Max.X && b.Min.X - tol <= a.Max.X
                && a.Min.Y - tol <= b.Max.Y && b.Min.Y - tol <= a.Max.Y
                && a.Min.Z - tol <= b.Max.Z && b.Min.Z - tol <= a.Max.Z;
        }

        private static BoundingBoxXYZ CloneBox(BoundingBoxXYZ b)
        {
            return new BoundingBoxXYZ { Min = b.Min, Max = b.Max };
        }

        private static BoundingBoxXYZ MergeBox(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            return new BoundingBoxXYZ
            {
                Min = new XYZ(Math.Min(a.Min.X, b.Min.X), Math.Min(a.Min.Y, b.Min.Y), Math.Min(a.Min.Z, b.Min.Z)),
                Max = new XYZ(Math.Max(a.Max.X, b.Max.X), Math.Max(a.Max.Y, b.Max.Y), Math.Max(a.Max.Z, b.Max.Z)),
            };
        }
    }
}

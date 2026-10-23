using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.GenericModelMerge.Services
{
    /// <summary>
    /// 3D ビューの切断ボックス（セクションボックス）の外側の形状を切り取る。
    ///
    /// ビュー Id 付きの FilteredElementCollector は「切断ボックスに一部でも入っている要素」を返すが、
    /// 要素の形状は切断ボックスと関係なく丸ごと取れる。そのままだと、はみ出した部分まで
    /// 一般モデルに入ってしまうため、切断ボックスと同じ大きさの箱との共通部分（Intersect）を取る。
    /// </summary>
    internal class SectionBoxClipper
    {
        private const double MinSolidVolume = 1e-6;

        /// <summary>箱の内側とみなす許容値（フィート）。約 0.3 mm。</summary>
        private const double InsideTolerance = 0.001;

        private readonly Solid _box;
        private readonly Transform _toLocal;
        private readonly XYZ _min;
        private readonly XYZ _max;

        /// <summary>はみ出した部分を切り取った形状の数。</summary>
        public int ClippedCount { get; private set; }

        /// <summary>切り取りに失敗し、元の形のまま残した形状の数。</summary>
        public int FailedCount { get; private set; }

        private SectionBoxClipper(Solid box, Transform toLocal, XYZ min, XYZ max)
        {
            _box = box;
            _toLocal = toLocal;
            _min = min;
            _max = max;
        }

        /// <summary>切断ボックスが有効でなければ null を返す。</summary>
        public static SectionBoxClipper FromView(View3D view)
        {
            if (view == null || !view.IsSectionBoxActive) return null;

            BoundingBoxXYZ sb;
            try { sb = view.GetSectionBox(); }
            catch { return null; }
            if (sb == null) return null;

            XYZ min = sb.Min, max = sb.Max;
            double height = max.Z - min.Z;
            if (max.X - min.X < InsideTolerance || max.Y - min.Y < InsideTolerance || height < InsideTolerance)
                return null;

            // GetSectionBox の Min/Max は箱の座標系の値。Transform で実際の位置（回転込み）に戻す
            var p0 = new XYZ(min.X, min.Y, min.Z);
            var p1 = new XYZ(max.X, min.Y, min.Z);
            var p2 = new XYZ(max.X, max.Y, min.Z);
            var p3 = new XYZ(min.X, max.Y, min.Z);
            var loop = new CurveLoop();
            loop.Append(Line.CreateBound(p0, p1));
            loop.Append(Line.CreateBound(p1, p2));
            loop.Append(Line.CreateBound(p2, p3));
            loop.Append(Line.CreateBound(p3, p0));

            Solid local = GeometryCreationUtilities.CreateExtrusionGeometry(
                new List<CurveLoop> { loop }, XYZ.BasisZ, height);

            Transform t = sb.Transform ?? Transform.Identity;
            Solid world = t.IsIdentity ? local : SolidUtils.CreateTransformed(local, t);

            return new SectionBoxClipper(world, t.Inverse, min, max);
        }

        /// <summary>
        /// 形状のうち切断ボックスの内側だけを返す。完全に外側の形状は返さない。
        /// 切り取りに失敗した形状は、黙って消さず元の形のまま返す（FailedCount で通知）。
        /// </summary>
        public List<Solid> Clip(IEnumerable<Solid> solids)
        {
            var result = new List<Solid>();
            foreach (var s in solids)
            {
                if (s == null || s.Volume <= MinSolidVolume) continue;

                // 箱の中に完全に収まっているものは切らない（Boolean は重く、失敗の元にもなる）
                if (IsInside(s)) { result.Add(s); continue; }

                Solid cut = null;
                try
                {
                    cut = BooleanOperationsUtils.ExecuteBooleanOperation(
                        s, _box, BooleanOperationsType.Intersect);
                }
                catch (Exception)
                {
                    FailedCount++;
                    result.Add(s);
                    continue;
                }

                // 共通部分が無い＝箱の外側にある形状
                if (cut == null || cut.Volume <= MinSolidVolume) continue;

                ClippedCount++;
                result.Add(cut);
            }
            return result;
        }

        private bool IsInside(Solid s)
        {
            BoundingBoxXYZ bb;
            try { bb = s.GetBoundingBox(); }
            catch { return false; }
            if (bb == null) return false;

            Transform toWorld = bb.Transform ?? Transform.Identity;
            const double t = InsideTolerance;
            foreach (double x in new[] { bb.Min.X, bb.Max.X })
            foreach (double y in new[] { bb.Min.Y, bb.Max.Y })
            foreach (double z in new[] { bb.Min.Z, bb.Max.Z })
            {
                XYZ p = _toLocal.OfPoint(toWorld.OfPoint(new XYZ(x, y, z)));
                if (p.X < _min.X - t || p.X > _max.X + t ||
                    p.Y < _min.Y - t || p.Y > _max.Y + t ||
                    p.Z < _min.Z - t || p.Z > _max.Z + t)
                    return false;
            }
            return true;
        }
    }
}

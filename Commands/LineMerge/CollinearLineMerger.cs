using System;
using System.Collections.Generic;

namespace Tools28.Commands.LineMerge
{
    /// <summary>3次元の点・方向。Revit API に依存させず、計算部分だけを単体で試せるようにしている。</summary>
    internal struct Vec3
    {
        public readonly double X, Y, Z;
        public Vec3(double x, double y, double z) { X = x; Y = y; Z = z; }
        public static Vec3 operator +(Vec3 a, Vec3 b) => new Vec3(a.X + b.X, a.Y + b.Y, a.Z + b.Z);
        public static Vec3 operator -(Vec3 a, Vec3 b) => new Vec3(a.X - b.X, a.Y - b.Y, a.Z - b.Z);
        public static Vec3 operator *(Vec3 a, double k) => new Vec3(a.X * k, a.Y * k, a.Z * k);
        public double Dot(Vec3 b) => X * b.X + Y * b.Y + Z * b.Z;
        public Vec3 Cross(Vec3 b) => new Vec3(Y * b.Z - Z * b.Y, Z * b.X - X * b.Z, X * b.Y - Y * b.X);
        public double Length => Math.Sqrt(Dot(this));
    }

    /// <summary>まとめる対象の1本の直線。Tag には呼び出し側の要素（ElementId など）を入れる。</summary>
    internal class MergeSegment<T>
    {
        public T Tag;
        public Vec3 Start;
        public Vec3 End;
        public double Length => (End - Start).Length;
    }

    /// <summary>まとめた結果。Keep を新しい端点に伸ばし、Remove を削除する。</summary>
    internal class MergeResult<T>
    {
        public MergeSegment<T> Keep;
        public List<MergeSegment<T>> Remove = new List<MergeSegment<T>>();
        public Vec3 NewStart;
        public Vec3 NewEnd;
    }

    /// <summary>
    /// 同一直線上にある直線をまとめる計算。
    /// 向きが平行で、両端が同じ直線の上にある線を1組にし、組ごとに端から端までの1本にする
    /// （運営者の判断で、離れている線どうしもすき間を埋めてつなぐ）。角度のある線はつながない。
    /// </summary>
    internal static class CollinearLineMerger
    {
        /// <summary>同じ直線とみなす距離の許容値（Revit の内部単位＝フィート）。約 0.5mm。</summary>
        public const double DistanceTolerance = 0.5 / 304.8;

        public static List<MergeResult<T>> Merge<T>(IList<MergeSegment<T>> segments)
        {
            var results = new List<MergeResult<T>>();

            // 長い線から順に「組の基準」にする（基準からの距離で判定するので、少しずつずれて連鎖することがない）
            var order = new List<MergeSegment<T>>(segments);
            order.RemoveAll(s => s.Length < 1e-9);
            order.Sort((a, b) => b.Length.CompareTo(a.Length));

            var used = new bool[order.Count];
            for (int i = 0; i < order.Count; i++)
            {
                if (used[i]) continue;
                used[i] = true;

                var basis = order[i];
                Vec3 origin = basis.Start;
                Vec3 dir = (basis.End - basis.Start) * (1.0 / basis.Length);
                var group = new List<MergeSegment<T>> { basis };

                for (int j = i + 1; j < order.Count; j++)
                {
                    if (used[j]) continue;
                    var s = order[j];
                    if (DistanceToLine(s.Start, origin, dir) > DistanceTolerance) continue;
                    if (DistanceToLine(s.End, origin, dir) > DistanceTolerance) continue;
                    used[j] = true;
                    group.Add(s);
                }

                if (group.Count < 2) continue;

                // 基準線の向きに沿った位置で、いちばん端の2点を求める
                double tMin = double.MaxValue, tMax = double.MinValue;
                foreach (var s in group)
                {
                    foreach (var p in new[] { s.Start, s.End })
                    {
                        double t = (p - origin).Dot(dir);
                        if (t < tMin) tMin = t;
                        if (t > tMax) tMax = t;
                    }
                }

                var r = new MergeResult<T>
                {
                    Keep = basis,
                    NewStart = origin + dir * tMin,
                    NewEnd = origin + dir * tMax,
                };
                for (int k = 1; k < group.Count; k++) r.Remove.Add(group[k]);
                results.Add(r);
            }

            return results;
        }

        /// <summary>点 p から、origin を通り向き dir（長さ1）の無限直線までの距離。</summary>
        private static double DistanceToLine(Vec3 p, Vec3 origin, Vec3 dir)
        {
            return (p - origin).Cross(dir).Length;
        }
    }
}

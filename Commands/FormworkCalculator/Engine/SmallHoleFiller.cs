using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 平面 Face の内側ループのうち「直径 ≤ MaxHoleDimMm の小さな穴」を検出して
    /// 「穴を塞ぐ」処理を行うヘルパー。
    ///
    /// 【ユーザー要件】
    /// 円形開口が 300φ 以下のような小さな穴は、現場で型枠に穴をあけずに
    /// そのまま貼って施工する。集計面積もその穴分を「型枠あり」として加算する。
    ///
    /// 【適用】 床 / 壁 / 梁 / 基礎 / 屋根 / その他。柱 (Column) は対象外。
    /// </summary>
    internal static class SmallHoleFiller
    {
        // 「小さな穴」と判定する境界寸法 (loop の BoundingBox 最大辺)
        public const double MaxHoleDimMm = 300.0;

        /// <summary>
        /// face の CurveLoops を取得し、内側ループ (穴) のうち BBox 最大寸法が
        /// MaxHoleDimMm 以下のものを除外する。
        /// </summary>
        /// <returns>
        ///   filteredLoops: 小さな穴を除いた CurveLoop リスト (押し出しに使う)
        ///   filledAreaFt2: 除外した小穴の面積合計 (ft²、数量加算に使う)
        ///   filledCount:   除外した小穴の数 (ログ用)
        /// </returns>
        public static (IList<CurveLoop> filteredLoops, double filledAreaFt2, int filledCount)
            Process(Face face)
        {
            if (face == null)
                return (null, 0.0, 0);

            IList<CurveLoop> loops;
            try { loops = face.GetEdgesAsCurveLoops(); }
            catch { return (null, 0.0, 0); }

            if (loops == null || loops.Count == 0)
                return (loops, 0.0, 0);

            // 内側ループが無い (穴無し) の場合は早期 return
            if (loops.Count == 1)
                return (loops, 0.0, 0);

            XYZ normal;
            try { normal = FaceClassifier.GetFaceNormal(face); }
            catch { normal = null; }
            if (normal == null)
                return (loops, 0.0, 0);

            double thresholdFt = UnitUtils.ConvertToInternalUnits(
                MaxHoleDimMm, UnitTypeId.Millimeters);

            var filtered = new List<CurveLoop>();
            double filledArea = 0.0;
            int filledCount = 0;

            // loops[0] は外周ループ (Revit API のドキュメント保証)、常に残す
            filtered.Add(loops[0]);

            for (int i = 1; i < loops.Count; i++)
            {
                var loop = loops[i];
                double maxDim = ComputeLoopMaxDim(loop);
                if (maxDim > 0 && maxDim <= thresholdFt)
                {
                    // 小さな穴 → 除外して面積を加算
                    double a = Math.Abs(ComputeLoopArea2D(loop, normal));
                    filledArea += a;
                    filledCount++;
                }
                else
                {
                    filtered.Add(loop);
                }
            }

            return (filtered, filledArea, filledCount);
        }

        /// <summary>
        /// CurveLoop の 3D バウンディングボックスの最大寸法を返す (ft 単位)。
        /// Arc の中点も評価して曲線が端点を超える場合に対応する。
        /// </summary>
        private static double ComputeLoopMaxDim(CurveLoop loop)
        {
            if (loop == null) return 0;
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            bool any = false;

            void Sample(XYZ p)
            {
                if (p == null) return;
                if (p.X < minX) minX = p.X;
                if (p.Y < minY) minY = p.Y;
                if (p.Z < minZ) minZ = p.Z;
                if (p.X > maxX) maxX = p.X;
                if (p.Y > maxY) maxY = p.Y;
                if (p.Z > maxZ) maxZ = p.Z;
                any = true;
            }

            foreach (Curve c in loop)
            {
                if (c == null) continue;
                try { Sample(c.GetEndPoint(0)); } catch { }
                try { Sample(c.GetEndPoint(1)); } catch { }
                try
                {
                    if (c.IsBound) Sample(c.Evaluate(0.5, true));
                }
                catch { }
            }

            if (!any) return 0;
            double dx = maxX - minX;
            double dy = maxY - minY;
            double dz = maxZ - minZ;
            return Math.Max(Math.Max(dx, dy), dz);
        }

        /// <summary>
        /// CurveLoop の面積を 2D 平面で計算する (Shoelace 公式)。
        /// Face の法線に垂直な平面に投影してから計算する。
        /// </summary>
        private static double ComputeLoopArea2D(CurveLoop loop, XYZ normal)
        {
            if (loop == null || normal == null) return 0;

            // 法線に垂直な 2 軸 (u, v) を構築
            XYZ u;
            if (Math.Abs(normal.Z) > 0.99)
                u = XYZ.BasisX;
            else
                u = XYZ.BasisZ.CrossProduct(normal).Normalize();
            XYZ v = normal.CrossProduct(u).Normalize();

            var pts = new List<(double x, double y)>();
            foreach (Curve c in loop)
            {
                if (c == null) continue;
                IList<XYZ> tess;
                try { tess = c.Tessellate(); }
                catch { continue; }
                if (tess == null) continue;
                // 末尾点は次曲線の先頭点と重複するため最後の要素を除外
                for (int i = 0; i < tess.Count - 1; i++)
                {
                    var p = tess[i];
                    pts.Add((p.DotProduct(u), p.DotProduct(v)));
                }
            }
            if (pts.Count < 3) return 0;

            // Shoelace
            double area2 = 0;
            int n = pts.Count;
            for (int i = 0; i < n; i++)
            {
                var (x1, y1) = pts[i];
                var (x2, y2) = pts[(i + 1) % n];
                area2 += x1 * y2 - x2 * y1;
            }
            return Math.Abs(area2) * 0.5;
        }
    }
}

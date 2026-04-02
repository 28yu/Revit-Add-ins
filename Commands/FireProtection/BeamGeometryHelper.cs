using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FireProtection
{
    /// <summary>
    /// 梁・柱の平面輪郭取得とオフセット・統合処理
    /// </summary>
    public static class BeamGeometryHelper
    {
        /// <summary>
        /// 要素のオフセット輪郭を取得
        /// 梁: LocationCurve（方向+幅が正確）、柱: BoundingBox
        /// </summary>
        public static CurveLoop GetElementOffsetOutline(
            Element element, View view, double offsetFeet,
            double startExt = -1, double endExt = -1)
        {
            if (view.ViewType == ViewType.Section)
                return GetOutlineForSectionView(element, view, offsetFeet);

            var fi = element as FamilyInstance;
            if (fi == null) return null;

            if (element.Category.Id.IntegerValue ==
                (int)BuiltInCategory.OST_StructuralFraming)
            {
                var outline = GetBeamOutlineFromCurve(fi, offsetFeet);
                if (outline != null) return outline;
            }

            return GetOutlineFromBoundingBox(element, view, offsetFeet);
        }

        /// <summary>
        /// 断面ビュー用: BoundingBoxをビュー座標系に投影して矩形を生成
        /// </summary>
        private static CurveLoop GetOutlineForSectionView(
            Element element, View view, double offsetFeet)
        {
            BoundingBoxXYZ bb = element.get_BoundingBox(view);
            if (bb == null)
                bb = element.get_BoundingBox(null);
            if (bb == null) return null;

            Transform viewTransform;
            try
            {
                viewTransform = view.CropBox.Transform;
            }
            catch
            {
                return GetOutlineFromBoundingBox(element, view, offsetFeet);
            }

            Transform inverse = viewTransform.Inverse;
            XYZ origin = viewTransform.Origin;
            XYZ rightDir = viewTransform.BasisX;
            XYZ upDir = viewTransform.BasisY;

            // BoundingBox全8頂点をビュー座標に変換し2D範囲を取得
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;

            for (int ix = 0; ix <= 1; ix++)
            for (int iy = 0; iy <= 1; iy++)
            for (int iz = 0; iz <= 1; iz++)
            {
                XYZ corner = new XYZ(
                    ix == 0 ? bb.Min.X : bb.Max.X,
                    iy == 0 ? bb.Min.Y : bb.Max.Y,
                    iz == 0 ? bb.Min.Z : bb.Max.Z);
                XYZ vp = inverse.OfPoint(corner);
                if (vp.X < minX) minX = vp.X;
                if (vp.Y < minY) minY = vp.Y;
                if (vp.X > maxX) maxX = vp.X;
                if (vp.Y > maxY) maxY = vp.Y;
            }

            minX -= offsetFeet;
            minY -= offsetFeet;
            maxX += offsetFeet;
            maxY += offsetFeet;

            if (maxX - minX < 0.001 || maxY - minY < 0.001)
                return null;

            // ビュー座標→モデル座標（断面平面上の3D点）
            XYZ p0 = origin + minX * rightDir + minY * upDir;
            XYZ p1 = origin + maxX * rightDir + minY * upDir;
            XYZ p2 = origin + maxX * rightDir + maxY * upDir;
            XYZ p3 = origin + minX * rightDir + maxY * upDir;

            CurveLoop loop = new CurveLoop();
            loop.Append(Line.CreateBound(p0, p1));
            loop.Append(Line.CreateBound(p1, p2));
            loop.Append(Line.CreateBound(p2, p3));
            loop.Append(Line.CreateBound(p3, p0));

            return loop;
        }

        /// <summary>
        /// 梁のロケーションカーブから平面輪郭を生成
        /// </summary>
        private static CurveLoop GetBeamOutlineFromCurve(
            FamilyInstance beam, double offsetFeet)
        {
            LocationCurve locCurve = beam.Location as LocationCurve;
            if (locCurve == null) return null;

            Line line = locCurve.Curve as Line;
            if (line == null) return null;

            XYZ start = new XYZ(line.GetEndPoint(0).X, line.GetEndPoint(0).Y, 0);
            XYZ end = new XYZ(line.GetEndPoint(1).X, line.GetEndPoint(1).Y, 0);

            if (start.DistanceTo(end) < 0.001)
                return null;

            XYZ direction = (end - start).Normalize();
            XYZ perp = new XYZ(-direction.Y, direction.X, 0);

            double beamWidth = GetBeamWidth(beam);
            double halfExtent = beamWidth / 2.0 + offsetFeet;

            // 端部延長: 梁幅/2 + offset（接合部ギャップをカバーしつつoffsetに比例）
            double endExt = beamWidth / 2.0 + offsetFeet;
            XYZ extStart = start - direction * endExt;
            XYZ extEnd = end + direction * endExt;

            XYZ p0 = extStart - perp * halfExtent;
            XYZ p1 = extEnd - perp * halfExtent;
            XYZ p2 = extEnd + perp * halfExtent;
            XYZ p3 = extStart + perp * halfExtent;

            CurveLoop loop = new CurveLoop();
            loop.Append(Line.CreateBound(p0, p1));
            loop.Append(Line.CreateBound(p1, p2));
            loop.Append(Line.CreateBound(p2, p3));
            loop.Append(Line.CreateBound(p3, p0));

            return loop;
        }

        /// <summary>
        /// BoundingBoxからオフセット矩形を生成
        /// </summary>
        private static CurveLoop GetOutlineFromBoundingBox(
            Element element, View view, double offsetFeet)
        {
            BoundingBoxXYZ bb = element.get_BoundingBox(view);
            if (bb == null)
                bb = element.get_BoundingBox(null);
            if (bb == null) return null;

            double minX = bb.Min.X - offsetFeet;
            double minY = bb.Min.Y - offsetFeet;
            double maxX = bb.Max.X + offsetFeet;
            double maxY = bb.Max.Y + offsetFeet;

            if (maxX - minX < 0.001 || maxY - minY < 0.001)
                return null;

            XYZ p0 = new XYZ(minX, minY, 0);
            XYZ p1 = new XYZ(maxX, minY, 0);
            XYZ p2 = new XYZ(maxX, maxY, 0);
            XYZ p3 = new XYZ(minX, maxY, 0);

            CurveLoop loop = new CurveLoop();
            loop.Append(Line.CreateBound(p0, p1));
            loop.Append(Line.CreateBound(p1, p2));
            loop.Append(Line.CreateBound(p2, p3));
            loop.Append(Line.CreateBound(p3, p0));

            return loop;
        }

        /// <summary>
        /// 複数の輪郭をソリッドブーリアン演算で統合し、
        /// 外周ループと統合失敗ループを分離して返す
        /// </summary>
        public static MergeResult MergeOutlines(List<CurveLoop> outlines)
        {
            var result = new MergeResult();

            if (outlines.Count == 0) return result;
            if (outlines.Count == 1)
            {
                result.MergedLoops.Add(outlines[0]);
                return result;
            }

            Solid unionSolid = null;

            foreach (var outline in outlines)
            {
                try
                {
                    Solid extrusion = GeometryCreationUtilities.CreateExtrusionGeometry(
                        new List<CurveLoop> { outline }, XYZ.BasisZ, 1.0);

                    if (unionSolid == null)
                    {
                        unionSolid = extrusion;
                    }
                    else
                    {
                        try
                        {
                            unionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(
                                unionSolid, extrusion, BooleanOperationsType.Union);
                        }
                        catch
                        {
                            result.UnmergedLoops.Add(outline);
                        }
                    }
                }
                catch
                {
                    result.UnmergedLoops.Add(outline);
                }
            }

            if (unionSolid != null)
            {
                ExtractOuterLoopsFromSolid(unionSolid, result);

                if (result.MergedLoops.Count == 0)
                {
                    result.UnmergedLoops.Clear();
                    result.UnmergedLoops.AddRange(outlines);
                }
            }
            else
            {
                result.UnmergedLoops.AddRange(outlines);
            }

            return result;
        }

        /// <summary>
        /// ソリッドの上面から外周ループのみ抽出（穴ループは除外）
        /// </summary>
        private static void ExtractOuterLoopsFromSolid(
            Solid solid, MergeResult result)
        {
            foreach (Face face in solid.Faces)
            {
                PlanarFace pf = face as PlanarFace;
                if (pf == null || Math.Abs(pf.FaceNormal.Z - 1.0) > 0.01)
                    continue;

                foreach (var edgeLoop in pf.GetEdgesAsCurveLoops())
                {
                    CurveLoop flatLoop = FlattenLoop(edgeLoop);
                    if (flatLoop == null || flatLoop.Count() == 0)
                        continue;

                    // 外周ループ（反時計回り）と穴ループ（時計回り）の
                    // 両方を保持。梁で囲まれた空間は穴として正しく表現
                    result.MergedLoops.Add(flatLoop);
                }
            }
        }

        /// <summary>
        /// CurveLoopをZ=0に平坦化
        /// </summary>
        private static CurveLoop FlattenLoop(IEnumerable<Curve> edgeLoop)
        {
            var flatLoop = new CurveLoop();

            foreach (Curve c in edgeLoop)
            {
                XYZ s = c.GetEndPoint(0);
                XYZ e = c.GetEndPoint(1);
                XYZ fs = new XYZ(s.X, s.Y, 0);
                XYZ fe = new XYZ(e.X, e.Y, 0);

                if (fs.DistanceTo(fe) < 0.0001)
                    continue;

                try
                {
                    flatLoop.Append(Line.CreateBound(fs, fe));
                }
                catch
                {
                    return null;
                }
            }

            return flatLoop;
        }

        /// <summary>
        /// CurveLoopが反時計回り（外周）かどうかをShoelace公式で判定
        /// Revit 2022にはCurveLoop.IsCounterClockwiseがないため手動実装
        /// </summary>
        private static bool IsLoopCounterClockwise(CurveLoop loop)
        {
            double signedArea = 0;
            foreach (Curve c in loop)
            {
                XYZ p0 = c.GetEndPoint(0);
                XYZ p1 = c.GetEndPoint(1);
                signedArea += (p1.X - p0.X) * (p1.Y + p0.Y);
            }
            // signedArea < 0 → 反時計回り（XY平面、Y上向き）
            return signedArea < 0;
        }

        /// <summary>
        /// 梁幅を取得
        /// </summary>
        public static double GetBeamWidth(FamilyInstance beam)
        {
            var paramNames = new[] { "b", "B", "幅", "梁幅", "W", "w", "Width", "width" };

            foreach (string name in paramNames)
            {
                Parameter p = beam.LookupParameter(name);
                if (p != null && p.StorageType == StorageType.Double)
                {
                    double val = p.AsDouble();
                    if (val > 0) return val;
                }
            }

            ElementType beamType = beam.Document.GetElement(beam.GetTypeId()) as ElementType;
            if (beamType != null)
            {
                foreach (string name in paramNames)
                {
                    Parameter p = beamType.LookupParameter(name);
                    if (p != null && p.StorageType == StorageType.Double)
                    {
                        double val = p.AsDouble();
                        if (val > 0) return val;
                    }
                }
            }

            BoundingBoxXYZ bb = beam.get_BoundingBox(null);
            if (bb != null)
            {
                LocationCurve lc = beam.Location as LocationCurve;
                if (lc != null && lc.Curve is Line line)
                {
                    XYZ dir = line.Direction.Normalize();
                    XYZ perp = new XYZ(-dir.Y, dir.X, 0).Normalize();
                    double width = Math.Abs((bb.Max - bb.Min).DotProduct(perp));
                    if (width > 0.01) return width;
                }
            }

            return 200.0 / 304.8;
        }

        /// <summary>
        /// 要素群から「耐火被覆」を含むパラメータを検出
        /// </summary>
        public static List<FireProtectionParameterInfo> DetectFireProtectionParameters(
            IEnumerable<Element> elements)
        {
            var paramValues = new Dictionary<string, HashSet<string>>();
            var paramCounts = new Dictionary<string, int>();

            foreach (var elem in elements)
            {
                foreach (Parameter p in elem.Parameters)
                {
                    if (p.Definition == null) continue;
                    string name = p.Definition.Name;
                    if (!name.Contains("耐火被覆")) continue;

                    if (!paramValues.ContainsKey(name))
                    {
                        paramValues[name] = new HashSet<string>();
                        paramCounts[name] = 0;
                    }

                    string value = null;
                    if (p.StorageType == StorageType.String)
                        value = p.AsString();
                    else if (p.StorageType == StorageType.Integer)
                        value = p.AsValueString();
                    else if (p.StorageType == StorageType.Double)
                        value = p.AsValueString();

                    if (!string.IsNullOrEmpty(value) && value.Trim().Length > 0)
                    {
                        paramValues[name].Add(value.Trim());
                        paramCounts[name]++;
                    }
                }
            }

            return paramValues
                .Select(kv => new FireProtectionParameterInfo
                {
                    ParameterName = kv.Key,
                    DetectedCount = paramCounts[kv.Key],
                    UniqueValues = kv.Value.OrderBy(v => v).ToList()
                })
                .OrderByDescending(p => p.DetectedCount)
                .ToList();
        }
    }
}

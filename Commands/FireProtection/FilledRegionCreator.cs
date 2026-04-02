using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FireProtection
{
    /// <summary>
    /// 耐火被覆塗潰領域の作成・管理
    /// </summary>
    public static class FilledRegionCreator
    {
        internal const string TypePrefix = "耐火被覆_";

        /// <summary>
        /// 耐火被覆種類ごとに塗潰領域を作成
        /// </summary>
        /// <param name="allStructuralElements">T字接合判定用の全構造要素（梁+柱、カテゴリ選択に関わらず）</param>
        public static int CreateFilledRegions(
            Document doc,
            View activeView,
            Dictionary<string, List<Element>> elementsByType,
            Dictionary<string, double> offsetByType,
            ElementId fillPatternId,
            ElementId lineStyleId,
            List<FireProtectionTypeEntry> orderedTypes,
            bool overwriteExisting)
        {
            if (overwriteExisting)
            {
                CleanupExistingRegions(doc, activeView);
            }

            int totalCreated = 0;

            foreach (var typeEntry in orderedTypes)
            {
                string typeName = typeEntry.Name;
                if (!elementsByType.ContainsKey(typeName))
                    continue;

                var elements = elementsByType[typeName];
                double offsetFeet = offsetByType.ContainsKey(typeName)
                    ? offsetByType[typeName] : 0;
                Color color = new Color(typeEntry.ColorR, typeEntry.ColorG, typeEntry.ColorB);

                string regionTypeName = TypePrefix + typeName;
                ElementId regionTypeId = GetOrCreateFilledRegionType(
                    doc, regionTypeName, fillPatternId, color);
                if (regionTypeId == null) continue;

                // LocationCurve+offsetで各要素のアウトラインを生成
                var outlines = new List<CurveLoop>();
                var beamSegments = new List<int>(); // outlines内のindex→elementsのindex
                var beamStarts = new List<XYZ>();
                var beamEnds = new List<XYZ>();

                for (int ei = 0; ei < elements.Count; ei++)
                {
                    var outline = BeamGeometryHelper.GetElementOffsetOutline(
                        elements[ei], activeView, offsetFeet);
                    if (outline == null) continue;
                    outlines.Add(outline);

                    // 梁の端点情報を記録（T字パッチ用）
                    var fi = elements[ei] as FamilyInstance;
                    if (fi != null && elements[ei].Category.Id.IntegerValue ==
                        (int)BuiltInCategory.OST_StructuralFraming)
                    {
                        LocationCurve lc = fi.Location as LocationCurve;
                        if (lc?.Curve != null)
                        {
                            var s = lc.Curve.GetEndPoint(0);
                            var e = lc.Curve.GetEndPoint(1);
                            beamSegments.Add(ei);
                            beamStarts.Add(new XYZ(s.X, s.Y, 0));
                            beamEnds.Add(new XYZ(e.X, e.Y, 0));
                        }
                    }
                }

                // T字接合パッチ: 処理対象梁同士のT字接合点にパッチ矩形を追加
                // （L字コーナー/自由端/非処理対象にはパッチなし）
                double patchTol = 500.0 / 304.8;
                var patchKeys = new HashSet<string>();

                for (int bi = 0; bi < beamSegments.Count; bi++)
                {
                    XYZ[] eps = { beamStarts[bi], beamEnds[bi] };
                    foreach (var ep in eps)
                    {
                        for (int bj = 0; bj < beamSegments.Count; bj++)
                        {
                            if (bi == bj) continue;

                            // T字判定: paramが0.05〜0.95なら梁本体上
                            double segLen = beamStarts[bj].DistanceTo(beamEnds[bj]);
                            if (segLen < 0.02) continue;

                            XYZ v = beamEnds[bj] - beamStarts[bj];
                            XYZ w = ep - beamStarts[bj];
                            double param = (w.X * v.X + w.Y * v.Y) / (v.X * v.X + v.Y * v.Y);
                            if (param <= 0.05 || param >= 0.95) continue;

                            // 最近点距離
                            XYZ closest = beamStarts[bj] + param * v;
                            if (ep.DistanceTo(closest) > patchTol) continue;

                            // 重複防止
                            string key = $"{Math.Round(ep.X, 3)},{Math.Round(ep.Y, 3)}";
                            if (patchKeys.Contains(key)) continue;
                            patchKeys.Add(key);

                            // パッチ: 接続先梁の半幅+offset の正方形
                            var cfi = elements[beamSegments[bj]] as FamilyInstance;
                            double cHalfW = cfi != null
                                ? BeamGeometryHelper.GetBeamWidth(cfi) / 2.0 + offsetFeet
                                : offsetFeet * 3;

                            XYZ p0 = new XYZ(ep.X - cHalfW, ep.Y - cHalfW, 0);
                            XYZ p1 = new XYZ(ep.X + cHalfW, ep.Y - cHalfW, 0);
                            XYZ p2 = new XYZ(ep.X + cHalfW, ep.Y + cHalfW, 0);
                            XYZ p3 = new XYZ(ep.X - cHalfW, ep.Y + cHalfW, 0);

                            var patch = new CurveLoop();
                            patch.Append(Line.CreateBound(p0, p1));
                            patch.Append(Line.CreateBound(p1, p2));
                            patch.Append(Line.CreateBound(p2, p3));
                            patch.Append(Line.CreateBound(p3, p0));
                            outlines.Add(patch);
                            break;
                        }
                    }
                }

                // デバッグログ
                try
                {
                    System.IO.Directory.CreateDirectory(@"C:\temp");
                    System.IO.File.AppendAllText(@"C:\temp\FireProtection_debug.txt",
                        $"\n[{System.DateTime.Now:HH:mm:ss}] {typeName} offset={offsetFeet * 304.8:F0}mm" +
                        $" beams={beamSegments.Count} patches={patchKeys.Count} outlines={outlines.Count}\n");
                }
                catch { }

                if (outlines.Count == 0) continue;

                // 断面ビュー: ビュー座標→XY平面に変換してunion→戻す
                List<CurveLoop> mergeInput = outlines;
                Transform sectionTransform = null;
                Transform sectionInverse = null;

                if (activeView.ViewType == ViewType.Section)
                {
                    try
                    {
                        sectionTransform = activeView.CropBox.Transform;
                        sectionInverse = sectionTransform.Inverse;
                        mergeInput = outlines
                            .Select(loop => TransformLoop(loop, sectionInverse))
                            .Where(loop => loop != null)
                            .ToList();
                    }
                    catch
                    {
                        sectionTransform = null;
                        mergeInput = outlines;
                    }
                }

                var mergeResult = BeamGeometryHelper.MergeOutlines(mergeInput);

                // 断面ビューの場合、マージ結果をモデル座標に戻す
                if (sectionTransform != null)
                {
                    mergeResult.MergedLoops = mergeResult.MergedLoops
                        .Select(loop => TransformLoop(loop, sectionTransform))
                        .Where(loop => loop != null)
                        .ToList();
                    mergeResult.UnmergedLoops = mergeResult.UnmergedLoops
                        .Select(loop => TransformLoop(loop, sectionTransform))
                        .Where(loop => loop != null)
                        .ToList();
                }

                // 統合済みループから塗潰領域を作成
                if (mergeResult.MergedLoops.Count > 0)
                {
                    try
                    {
                        var region = FilledRegion.Create(
                            doc, regionTypeId, activeView.Id, mergeResult.MergedLoops);
                        ApplyLineStyle(region, lineStyleId);
                        totalCreated++;
                    }
                    catch
                    {
                        foreach (var loop in mergeResult.MergedLoops)
                        {
                            try
                            {
                                var region = FilledRegion.Create(
                                    doc, regionTypeId, activeView.Id,
                                    new List<CurveLoop> { loop });
                                ApplyLineStyle(region, lineStyleId);
                                totalCreated++;
                            }
                            catch { }
                        }
                    }
                }

                foreach (var loop in mergeResult.UnmergedLoops)
                {
                    try
                    {
                        var region = FilledRegion.Create(
                            doc, regionTypeId, activeView.Id,
                            new List<CurveLoop> { loop });
                        ApplyLineStyle(region, lineStyleId);
                        totalCreated++;
                    }
                    catch { }
                }
            }

            return totalCreated;
        }

        private static CurveLoop TransformLoop(CurveLoop loop, Transform transform)
        {
            try
            {
                var newLoop = new CurveLoop();
                foreach (Curve c in loop)
                {
                    XYZ s = transform.OfPoint(c.GetEndPoint(0));
                    XYZ e = transform.OfPoint(c.GetEndPoint(1));
                    if (s.DistanceTo(e) < 0.0001) continue;
                    newLoop.Append(Line.CreateBound(s, e));
                }
                return newLoop.Count() > 0 ? newLoop : null;
            }
            catch { return null; }
        }

        private static void ApplyLineStyle(FilledRegion region, ElementId lineStyleId)
        {
            if (lineStyleId != null && lineStyleId != ElementId.InvalidElementId)
            {
                region.SetLineStyleId(lineStyleId);
            }
        }

        private static ElementId GetOrCreateFilledRegionType(
            Document doc, string typeName, ElementId fillPatternId, Color color)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault(t => t.Name == typeName);

            if (existing != null)
                doc.Delete(existing.Id);

            var baseType = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault();

            if (baseType == null) return null;

            var newType = baseType.Duplicate(typeName) as FilledRegionType;
            if (newType == null) return null;

            newType.ForegroundPatternColor = color;
            if (fillPatternId != null && fillPatternId != ElementId.InvalidElementId)
                newType.ForegroundPatternId = fillPatternId;
            newType.BackgroundPatternId = ElementId.InvalidElementId;
            newType.IsMasking = false;

            return newType.Id;
        }

        private static void CleanupExistingRegions(Document doc, View view)
        {
            var existingRegions = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(FilledRegion))
                .Cast<FilledRegion>()
                .Where(fr =>
                {
                    var frType = doc.GetElement(fr.GetTypeId()) as FilledRegionType;
                    return frType != null && frType.Name.StartsWith(TypePrefix);
                })
                .Select(fr => fr.Id)
                .ToList();

            foreach (var id in existingRegions)
            {
                try { doc.Delete(id); } catch { }
            }

            var existingTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .Where(t => t.Name.StartsWith(TypePrefix))
                .Select(t => t.Id)
                .ToList();

            foreach (var id in existingTypes)
            {
                try { doc.Delete(id); } catch { }
            }
        }

        internal static List<Color> GenerateColors(int count)
        {
            var baseColors = new[]
            {
                new { R = 144, G = 175, B = 197 },
                new { R = 165, G = 196, B = 152 },
                new { R = 210, G = 165, B = 120 },
                new { R = 180, G = 150, B = 180 },
                new { R = 120, G = 180, B = 190 },
                new { R = 200, G = 160, B = 150 },
                new { R = 160, G = 185, B = 130 },
                new { R = 185, G = 170, B = 140 },
                new { R = 150, G = 170, B = 200 },
                new { R = 200, G = 185, B = 150 },
            };

            var colors = new List<Color>();
            for (int i = 0; i < count; i++)
            {
                var bc = baseColors[i % baseColors.Length];
                int r = bc.R, g = bc.G, b = bc.B;

                if (i >= baseColors.Length)
                {
                    float brightness = 0.8f + (float)(i % baseColors.Length)
                        / baseColors.Length * 0.4f;
                    r = Math.Min((int)(r * brightness), 255);
                    g = Math.Min((int)(g * brightness), 255);
                    b = Math.Min((int)(b * brightness), 255);
                }

                colors.Add(new Color((byte)r, (byte)g, (byte)b));
            }

            return colors;
        }
    }
}

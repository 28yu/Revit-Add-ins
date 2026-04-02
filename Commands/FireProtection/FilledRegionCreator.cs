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
            bool overwriteExisting,
            List<Element> allStructuralElements = null,
            HashSet<int> processingElementIds = null)
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

                // T字接合判定用: 全構造要素（梁+柱）の線分/位置を収集
                // カテゴリ選択に関わらず、接続先の幅を考慮するため
                var refStarts = new List<XYZ>();
                var refEnds = new List<XYZ>();
                var refWidths = new List<double>();
                var refElemIds = new List<int>(); // 参照要素のElementId

                foreach (var re in (allStructuralElements ?? new List<Element>()))
                {
                    var rfi = re as FamilyInstance;
                    if (rfi == null) continue;
                    int reId = re.Id.IntegerValue;

                    // 梁: LocationCurve
                    LocationCurve rlc = rfi.Location as LocationCurve;
                    if (rlc?.Curve != null)
                    {
                        var rs = rlc.Curve.GetEndPoint(0);
                        var re2 = rlc.Curve.GetEndPoint(1);
                        refStarts.Add(new XYZ(rs.X, rs.Y, 0));
                        refEnds.Add(new XYZ(re2.X, re2.Y, 0));
                        refElemIds.Add(reId);

                        if (re.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                            refWidths.Add(BeamGeometryHelper.GetBeamWidth(rfi));
                        else
                            refWidths.Add(GetColumnWidth(rfi));
                        continue;
                    }

                    // 柱: LocationPointからBoundingBoxで幅推定
                    LocationPoint rlp = rfi.Location as LocationPoint;
                    if (rlp != null)
                    {
                        BoundingBoxXYZ rbb = re.get_BoundingBox(activeView);
                        if (rbb == null) rbb = re.get_BoundingBox(null);
                        if (rbb != null)
                        {
                            double cw = Math.Max(rbb.Max.X - rbb.Min.X, rbb.Max.Y - rbb.Min.Y);
                            XYZ cp = new XYZ(rlp.Point.X, rlp.Point.Y, 0);
                            refStarts.Add(new XYZ(cp.X, cp.Y - 0.01, 0));
                            refEnds.Add(new XYZ(cp.X, cp.Y + 0.01, 0));
                            refWidths.Add(cw);
                            refElemIds.Add(reId);
                        }
                    }
                }

                // 処理対象梁の情報
                var bStarts = new List<XYZ>();
                var bEnds = new List<XYZ>();
                var bIdx = new List<int>();

                for (int ei = 0; ei < elements.Count; ei++)
                {
                    var fi = elements[ei] as FamilyInstance;
                    if (fi == null) continue;
                    if (elements[ei].Category.Id.IntegerValue !=
                        (int)BuiltInCategory.OST_StructuralFraming) continue;
                    LocationCurve lc = fi.Location as LocationCurve;
                    if (lc?.Curve == null) continue;

                    var s = lc.Curve.GetEndPoint(0);
                    var e = lc.Curve.GetEndPoint(1);
                    bStarts.Add(new XYZ(s.X, s.Y, 0));
                    bEnds.Add(new XYZ(e.X, e.Y, 0));
                    bIdx.Add(ei);
                }

                double tol = 500.0 / 304.8;
                var processed = new HashSet<int>();
                var debugLines = new System.Collections.Generic.List<string>();
                debugLines.Add($"refElements={refStarts.Count}, targetBeams={bIdx.Count}");

                var outlines = new List<CurveLoop>();
                for (int bi = 0; bi < bIdx.Count; bi++)
                {
                    processed.Add(bIdx[bi]);
                    double sExt = offsetFeet;
                    double eExt = offsetFeet;

                    // 全構造要素に対してT字判定
                    for (int rj = 0; rj < refStarts.Count; rj++)
                    {
                        bool refIsProcessing = processingElementIds != null
                            && processingElementIds.Contains(refElemIds[rj]);

                        // 始端
                        if (sExt <= offsetFeet + 0.001)
                        {
                            double dist = PointToSegDist(bStarts[bi], refStarts[rj], refEnds[rj]);
                            if (dist < tol)
                            {
                                bool isTjunc = false;
                                double segLen = refStarts[rj].DistanceTo(refEnds[rj]);
                                if (segLen > 0.02)
                                {
                                    double p = ProjectParam(bStarts[bi], refStarts[rj], refEnds[rj]);
                                    isTjunc = (p > 0.05 && p < 0.95);
                                }
                                else
                                {
                                    isTjunc = true;
                                }

                                if (isTjunc)
                                {
                                    double hw = refWidths[rj] / 2.0;
                                    // 処理対象: offset端に合わせる
                                    // 非処理対象: 要素面まで（offsetは足さない）
                                    double ext = refIsProcessing ? hw + offsetFeet : hw;
                                    // offsetより小さければ延長不要
                                    if (ext > sExt)
                                    {
                                        sExt = ext;
                                        debugLines.Add($"  beam[{bi}]start T ref[{rj}] ext={sExt * 304.8:F0}mm rw={refWidths[rj] * 304.8:F0}mm {(refIsProcessing ? "proc" : "face")}");
                                    }
                                }
                            }
                        }

                        // 終端
                        if (eExt <= offsetFeet + 0.001)
                        {
                            double dist = PointToSegDist(bEnds[bi], refStarts[rj], refEnds[rj]);
                            if (dist < tol)
                            {
                                bool isTjunc = false;
                                double segLen = refStarts[rj].DistanceTo(refEnds[rj]);
                                if (segLen > 0.02)
                                {
                                    double p = ProjectParam(bEnds[bi], refStarts[rj], refEnds[rj]);
                                    isTjunc = (p > 0.05 && p < 0.95);
                                }
                                else
                                {
                                    isTjunc = true;
                                }

                                if (isTjunc)
                                {
                                    double hw = refWidths[rj] / 2.0;
                                    double ext = refIsProcessing ? hw + offsetFeet : hw;
                                    if (ext > eExt)
                                    {
                                        eExt = ext;
                                        debugLines.Add($"  beam[{bi}]end T ref[{rj}] ext={eExt * 304.8:F0}mm rw={refWidths[rj] * 304.8:F0}mm {(refIsProcessing ? "proc" : "face")}");
                                    }
                                }
                            }
                        }
                    }

                    debugLines.Add($"  beam[{bi}] sExt={sExt * 304.8:F0}mm eExt={eExt * 304.8:F0}mm");

                    var outline = BeamGeometryHelper.GetElementOffsetOutline(
                        elements[bIdx[bi]], activeView, offsetFeet, sExt, eExt);
                    if (outline != null) outlines.Add(outline);
                }

                // 梁以外（柱等）
                for (int ei = 0; ei < elements.Count; ei++)
                {
                    if (processed.Contains(ei)) continue;
                    var outline = BeamGeometryHelper.GetElementOffsetOutline(
                        elements[ei], activeView, offsetFeet);
                    if (outline != null) outlines.Add(outline);
                }

                // デバッグログ
                try
                {
                    System.IO.Directory.CreateDirectory(@"C:\temp");
                    System.IO.File.AppendAllText(@"C:\temp\FireProtection_debug.txt",
                        $"\n[{System.DateTime.Now:HH:mm:ss}] {typeName} offset={offsetFeet * 304.8:F0}mm\n"
                        + string.Join("\n", debugLines) + "\n");
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

        private static double GetColumnWidth(FamilyInstance column)
        {
            BoundingBoxXYZ bb = column.get_BoundingBox(null);
            if (bb != null)
            {
                double xExt = bb.Max.X - bb.Min.X;
                double yExt = bb.Max.Y - bb.Min.Y;
                // 小さい方の寸法 = 平面上の幅（大きい方は高さや奥行の可能性）
                return Math.Min(xExt, yExt);
            }
            return 400.0 / 304.8;
        }

        /// <summary>
        /// 点を線分に射影したパラメータ（0=始点、1=終点）
        /// </summary>
        private static double ProjectParam(XYZ pt, XYZ a, XYZ b)
        {
            double vx = b.X - a.X, vy = b.Y - a.Y;
            double wx = pt.X - a.X, wy = pt.Y - a.Y;
            double c2 = vx * vx + vy * vy;
            if (c2 < 1e-10) return 0;
            return (wx * vx + wy * vy) / c2;
        }

        /// <summary>
        /// 2D点から線分への最短距離
        /// </summary>
        private static double PointToSegDist(XYZ pt, XYZ a, XYZ b)
        {
            double t = ProjectParam(pt, a, b);
            if (t <= 0) return pt.DistanceTo(a);
            if (t >= 1) return pt.DistanceTo(b);
            return pt.DistanceTo(new XYZ(a.X + t * (b.X - a.X), a.Y + t * (b.Y - a.Y), 0));
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

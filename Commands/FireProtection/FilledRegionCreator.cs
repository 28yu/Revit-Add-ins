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
                Color fgColor = new Color(typeEntry.ColorR, typeEntry.ColorG, typeEntry.ColorB);
                Color bgColor = new Color(typeEntry.BgColorR, typeEntry.BgColorG, typeEntry.BgColorB);

                string regionTypeName = TypePrefix + typeName;
                ElementId regionTypeId = GetOrCreateFilledRegionType(
                    doc, regionTypeName, fillPatternId, fgColor,
                    typeEntry.ForegroundVisible, bgColor, typeEntry.BackgroundVisible);
                if (regionTypeId == null) continue;

                // 各要素のアウトラインを生成
                var outlines = new List<CurveLoop>();
                foreach (var elem in elements)
                {
                    var outline = BeamGeometryHelper.GetElementOffsetOutline(
                        elem, activeView, offsetFeet);
                    if (outline != null)
                        outlines.Add(outline);
                }

                // デバッグログ
                try
                {
                    System.IO.Directory.CreateDirectory(@"C:\temp");
                    System.IO.File.AppendAllText(@"C:\temp\FireProtection_debug.txt",
                        $"\n[{System.DateTime.Now:HH:mm:ss}] {typeName} offset={offsetFeet * 304.8:F0}mm" +
                        $" outlines={outlines.Count}\n");
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

        /// <summary>
        /// 断面ビュー: 柱の塗潰領域を梁オフセット端でクリップして作成
        /// </summary>
        public static int CreateSectionColumnRegions(
            Document doc, View view,
            Dictionary<string, List<Element>> colsByType,
            List<BoundingBoxXYZ> beamBBoxes,
            double beamOffset,
            ElementId fillPatternId, ElementId lineStyleId,
            List<FireProtectionTypeEntry> orderedTypes,
            bool overwriteExisting)
        {
            int created = 0;

            foreach (var typeEntry in orderedTypes)
            {
                if (!colsByType.ContainsKey(typeEntry.Name)) continue;
                var columns = colsByType[typeEntry.Name];

                string regionTypeName = TypePrefix + typeEntry.Name;
                var existingType = new FilteredElementCollector(doc)
                    .OfClass(typeof(FilledRegionType))
                    .Cast<FilledRegionType>()
                    .FirstOrDefault(t => t.Name == regionTypeName);

                // タイプがなければ作成（この種類に梁がないビューの場合）
                ElementId typeId;
                if (existingType != null)
                {
                    typeId = existingType.Id;
                }
                else
                {
                    Color fgColor = new Color(typeEntry.ColorR, typeEntry.ColorG, typeEntry.ColorB);
                    typeId = GetOrCreateFilledRegionType(
                        doc, regionTypeName, fillPatternId, fgColor,
                        typeEntry.ForegroundVisible);
                    if (typeId == null) continue;
                }

                foreach (var col in columns)
                {
                    var outline = BeamGeometryHelper.GetColumnOutlineClippedByBeams(
                        col, view, beamOffset, beamBBoxes, beamOffset);
                    if (outline == null) continue;

                    try
                    {
                        var region = FilledRegion.Create(doc, typeId,
                            view.Id, new List<CurveLoop> { outline });
                        if (lineStyleId != null && lineStyleId != ElementId.InvalidElementId)
                            region.SetLineStyleId(lineStyleId);
                        created++;
                    }
                    catch { }
                }
            }

            return created;
        }

        /// <summary>
        /// 柱の枠型（ドーナツ型）塗潰領域を作成
        /// </summary>
        public static int CreateColumnFrameRegions(
            Document doc,
            View activeView,
            List<Element> allColumns,
            string paramName,
            double aFeet, double bFeet,
            ElementId fillPatternId,
            ElementId lineStyleId,
            List<FireProtectionTypeEntry> orderedTypes,
            bool overwriteExisting)
        {
            int created = 0;

            // タイプごとにFilledRegionTypeを事前作成（1回だけ）
            var colTypeIds = new Dictionary<string, ElementId>();
            foreach (var typeEntry in orderedTypes)
            {
                string colTypeName = TypePrefix + "柱_" + typeEntry.Name;
                Color colFgColor = new Color(typeEntry.ColColorR, typeEntry.ColColorG, typeEntry.ColColorB);
                Color colBgColor = new Color(typeEntry.ColBgColorR, typeEntry.ColBgColorG, typeEntry.ColBgColorB);
                ElementId colTypeId = GetOrCreateFilledRegionType(
                    doc, colTypeName, fillPatternId, colFgColor,
                    typeEntry.ColForegroundVisible, colBgColor, typeEntry.ColBackgroundVisible);
                if (colTypeId != null)
                    colTypeIds[typeEntry.Name] = colTypeId;
            }

            foreach (var col in allColumns)
            {
                var fi = col as FamilyInstance;
                if (fi == null) continue;

                Parameter p = col.LookupParameter(paramName);
                if (p == null) continue;
                string value = p.StorageType == StorageType.String
                    ? p.AsString() : p.AsValueString();
                if (string.IsNullOrEmpty(value) || value.Trim().Length == 0) continue;
                value = value.Trim();

                if (!colTypeIds.ContainsKey(value)) continue;
                ElementId colTypeId = colTypeIds[value];

                LocationPoint lp = fi.Location as LocationPoint;
                if (lp == null) continue;
                XYZ center = new XYZ(lp.Point.X, lp.Point.Y, 0);

                try { System.IO.File.AppendAllText(@"C:\temp\FireProtection_debug.txt",
                    $"  柱枠作成: {value} center=({center.X * 304.8:F0},{center.Y * 304.8:F0}) A={aFeet * 304.8:F0} B={bFeet * 304.8:F0}\n"); } catch { }

                // 外側矩形: center ± (A + B)
                double outer = aFeet + bFeet;
                CurveLoop outerLoop = CreateRectLoop(center, outer);

                // 内側矩形: center ± A（穴）
                CurveLoop innerLoop = CreateRectLoop(center, aFeet);
                // 穴ループは時計回りにする必要がある
                innerLoop.Flip();

                try
                {
                    var region = FilledRegion.Create(doc, colTypeId,
                        activeView.Id, new List<CurveLoop> { outerLoop, innerLoop });
                    if (lineStyleId != null && lineStyleId != ElementId.InvalidElementId)
                        region.SetLineStyleId(lineStyleId);
                    created++;
                }
                catch (Exception ex)
                {
                    try
                    {
                        System.IO.File.AppendAllText(@"C:\temp\FireProtection_debug.txt",
                            $"  柱枠エラー: {ex.Message}\n");
                    }
                    catch { }
                }
            }

            return created;
        }

        private static CurveLoop CreateRectLoop(XYZ center, double halfSize)
        {
            XYZ p0 = new XYZ(center.X - halfSize, center.Y - halfSize, 0);
            XYZ p1 = new XYZ(center.X + halfSize, center.Y - halfSize, 0);
            XYZ p2 = new XYZ(center.X + halfSize, center.Y + halfSize, 0);
            XYZ p3 = new XYZ(center.X - halfSize, center.Y + halfSize, 0);

            CurveLoop loop = new CurveLoop();
            loop.Append(Line.CreateBound(p0, p1));
            loop.Append(Line.CreateBound(p1, p2));
            loop.Append(Line.CreateBound(p2, p3));
            loop.Append(Line.CreateBound(p3, p0));
            return loop;
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
            Document doc, string typeName, ElementId fillPatternId, Color fgColor,
            bool fgVisible = true, Color bgColor = null, bool bgVisible = false)
        {
            // 既存タイプがあれば再利用（複数ビュー処理時に削除しない）
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault(t => t.Name == typeName);

            if (existing != null)
                return existing.Id;

            var baseType = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault();

            if (baseType == null) return null;

            var newType = baseType.Duplicate(typeName) as FilledRegionType;
            if (newType == null) return null;

            // 前景
            if (fgVisible && fillPatternId != null && fillPatternId != ElementId.InvalidElementId)
            {
                newType.ForegroundPatternColor = fgColor;
                newType.ForegroundPatternId = fillPatternId;
            }
            else
            {
                newType.ForegroundPatternId = ElementId.InvalidElementId;
            }

            // 背景
            if (bgVisible && bgColor != null && fillPatternId != null
                && fillPatternId != ElementId.InvalidElementId)
            {
                newType.BackgroundPatternColor = bgColor;
                newType.BackgroundPatternId = fillPatternId;
            }
            else
            {
                newType.BackgroundPatternId = ElementId.InvalidElementId;
            }

            newType.IsMasking = false;

            return newType.Id;
        }

        /// <summary>
        /// 全ビューの処理前に1回だけ呼ぶ: 既存のFilledRegionTypeを削除
        /// </summary>
        public static void CleanupExistingTypes(Document doc)
        {
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

        private static void CleanupExistingRegions(Document doc, View view)
        {
            // このビューのRegionだけ削除（Typeは削除しない）
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
        }

        /// <summary>
        /// 梁用の色生成（薄めの色合い）
        /// </summary>
        internal static List<Color> GenerateBeamColors(int count)
        {
            var baseColors = new[]
            {
                new { R = 255, G = 153, B = 104 }, // 薄赤
                new { R = 128, G = 191, B = 255 }, // 薄青
                new { R = 100, G = 255, B = 100 }, // 薄緑
            };
            return GenerateFromBase(baseColors, count, 0.85f);
        }

        /// <summary>
        /// 柱用の色生成（濃いめの色合い）
        /// </summary>
        internal static List<Color> GenerateColumnColors(int count)
        {
            var baseColors = new[]
            {
                new { R = 255, G = 0, B = 0 },   // 濃赤
                new { R = 0, G = 0, B = 255 },   // 濃青
                new { R = 0, G = 128, B = 0 },   // 濃緑
            };
            return GenerateFromBase(baseColors, count, 0.7f);
        }

        /// <summary>
        /// 互換用（既存呼び出し）
        /// </summary>
        internal static List<Color> GenerateColors(int count)
        {
            return GenerateBeamColors(count);
        }

        private static List<Color> GenerateFromBase(
            dynamic[] baseColors, int count, float shiftFactor)
        {
            var colors = new List<Color>();
            for (int i = 0; i < count; i++)
            {
                int idx = i % baseColors.Length;
                int r = (int)baseColors[idx].R;
                int g = (int)baseColors[idx].G;
                int b = (int)baseColors[idx].B;

                if (i >= baseColors.Length)
                {
                    // 4色目以降: 色相をシフトして自動生成
                    int cycle = i / baseColors.Length;
                    float factor = shiftFactor + (cycle % 3) * 0.1f;
                    r = Math.Min((int)(r * factor + (1 - factor) * 128), 255);
                    g = Math.Min((int)(g * factor + (1 - factor) * 128), 255);
                    b = Math.Min((int)(b * factor + (1 - factor) * 128), 255);
                }

                colors.Add(new Color((byte)r, (byte)g, (byte)b));
            }
            return colors;
        }
    }
}

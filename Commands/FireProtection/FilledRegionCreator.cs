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
        public static int CreateFilledRegions(
            Document doc,
            View activeView,
            Dictionary<string, List<FamilyInstance>> beamsByType,
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

            var colors = GenerateColors(orderedTypes.Count);
            int totalCreated = 0;

            for (int i = 0; i < orderedTypes.Count; i++)
            {
                string typeName = orderedTypes[i].Name;
                if (!beamsByType.ContainsKey(typeName))
                    continue;

                var beams = beamsByType[typeName];
                double offsetFeet = offsetByType.ContainsKey(typeName)
                    ? offsetByType[typeName] : 0;
                Color color = colors[i];

                string regionTypeName = TypePrefix + typeName;
                ElementId regionTypeId = GetOrCreateFilledRegionType(
                    doc, regionTypeName, fillPatternId, color);
                if (regionTypeId == null) continue;

                var outlines = new List<CurveLoop>();
                foreach (var beam in beams)
                {
                    var outline = BeamGeometryHelper.GetBeamOffsetOutline(beam, offsetFeet);
                    if (outline != null)
                        outlines.Add(outline);
                }

                if (outlines.Count == 0) continue;

                var mergedOutlines = BeamGeometryHelper.MergeOutlines(outlines);

                try
                {
                    var region = FilledRegion.Create(
                        doc, regionTypeId, activeView.Id, mergedOutlines);

                    if (lineStyleId != null && lineStyleId != ElementId.InvalidElementId)
                    {
                        region.SetLineStyleId(lineStyleId);
                    }

                    totalCreated++;
                }
                catch
                {
                    foreach (var outline in outlines)
                    {
                        try
                        {
                            var region = FilledRegion.Create(
                                doc, regionTypeId, activeView.Id,
                                new List<CurveLoop> { outline });

                            if (lineStyleId != null && lineStyleId != ElementId.InvalidElementId)
                            {
                                region.SetLineStyleId(lineStyleId);
                            }

                            totalCreated++;
                        }
                        catch { }
                    }
                }
            }

            return totalCreated;
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

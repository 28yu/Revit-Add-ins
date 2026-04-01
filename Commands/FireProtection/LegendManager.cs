using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FireProtection
{
    /// <summary>
    /// 耐火被覆色分け凡例の製図ビューを作成
    /// </summary>
    public static class LegendManager
    {
        private const string LegendViewName = "耐火被覆色分け凡例";
        private const string RegionTypePrefix = "耐火被覆_凡例_";

        public static ElementId CreateLegendDraftingView(
            Document doc,
            List<FireProtectionTypeEntry> types,
            Dictionary<string, int> elementCounts,
            bool overwriteExisting,
            ElementId textNoteTypeId = null)
        {
            try
            {
                var existingView = FindExistingLegendView(doc);
                if (existingView != null)
                {
                    if (overwriteExisting)
                        doc.Delete(existingView.Id);
                    else
                        return null;
                }

                if (overwriteExisting)
                    CleanupLegendRegionTypes(doc);

                ElementId viewFamilyTypeId = GetDraftingViewFamilyTypeId(doc);
                if (viewFamilyTypeId == null) return null;

                ViewDrafting draftingView = ViewDrafting.Create(doc, viewFamilyTypeId);
                try { draftingView.Name = LegendViewName; }
                catch
                {
                    draftingView.Name = LegendViewName + "_" +
                        DateTime.Now.ToString("yyyyMMdd_HHmmss");
                }

                draftingView.Scale = 1;

                ElementId solidFillPatternId = GetSolidFillPatternId(doc);
                if (solidFillPatternId == null) return draftingView.Id;

                if (textNoteTypeId == null)
                    textNoteTypeId = GetDefaultTextNoteTypeId(doc);
                if (textNoteTypeId == null) return draftingView.Id;

                double textHeight = GetTextHeight(doc, textNoteTypeId);
                double rectHeight = textHeight * 1.8;
                double rectWidth = textHeight * 4.0;
                double textOffsetX = rectWidth + textHeight * 0.8;
                double rowSpacing = textHeight * 0.8;
                double titleGap = textHeight * 3.0;
                double textYOffset = (rectHeight + textHeight) / 2 + textHeight * 0.3;
                double rowPitch = rectHeight + rowSpacing;

                double startY = 0;

                XYZ titlePos = new XYZ(0, startY, 0);
                TextNote.Create(doc, draftingView.Id, titlePos,
                    "耐火被覆色分け凡例", textNoteTypeId);
                startY -= (textHeight + titleGap);

                for (int i = 0; i < types.Count; i++)
                {
                    var entry = types[i];
                    Color color = new Color(entry.ColorR, entry.ColorG, entry.ColorB);
                    int count = elementCounts.ContainsKey(entry.Name)
                        ? elementCounts[entry.Name] : 0;

                    double rowY = startY - rowPitch * i;

                    CreateColoredRectangle(doc, draftingView.Id,
                        solidFillPatternId, color, RegionTypePrefix + entry.Name,
                        0, rowY, rectWidth, rectHeight);

                    XYZ textPos = new XYZ(textOffsetX, rowY + textYOffset, 0);
                    string label = $"- {entry.Name}\uff08{count}\u672c\uff09";
                    TextNote.Create(doc, draftingView.Id, textPos, label, textNoteTypeId);
                }

                return draftingView.Id;
            }
            catch
            {
                return null;
            }
        }

        private static void CreateColoredRectangle(
            Document doc, ElementId viewId,
            ElementId solidFillPatternId, Color color,
            string typeName,
            double x, double y, double width, double height)
        {
            ElementId regionTypeId = GetOrCreateFilledRegionType(
                doc, solidFillPatternId, color, typeName);
            if (regionTypeId == null) return;

            XYZ p0 = new XYZ(x, y, 0);
            XYZ p1 = new XYZ(x + width, y, 0);
            XYZ p2 = new XYZ(x + width, y + height, 0);
            XYZ p3 = new XYZ(x, y + height, 0);

            CurveLoop loop = new CurveLoop();
            loop.Append(Line.CreateBound(p0, p1));
            loop.Append(Line.CreateBound(p1, p2));
            loop.Append(Line.CreateBound(p2, p3));
            loop.Append(Line.CreateBound(p3, p0));

            FilledRegion.Create(doc, regionTypeId, viewId,
                new List<CurveLoop> { loop });
        }

        private static ElementId GetOrCreateFilledRegionType(
            Document doc, ElementId solidFillPatternId,
            Color color, string typeName)
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
            newType.ForegroundPatternId = solidFillPatternId;
            newType.BackgroundPatternId = ElementId.InvalidElementId;
            newType.IsMasking = false;

            return newType.Id;
        }

        private static double GetTextHeight(Document doc, ElementId textNoteTypeId)
        {
            var tnt = doc.GetElement(textNoteTypeId) as TextNoteType;
            if (tnt != null)
            {
                Parameter sizeParam = tnt.get_Parameter(BuiltInParameter.TEXT_SIZE);
                if (sizeParam != null)
                    return sizeParam.AsDouble();
            }
            return 2.5 / 304.8;
        }

        private static ViewDrafting FindExistingLegendView(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewDrafting))
                .Cast<ViewDrafting>()
                .FirstOrDefault(v => v.Name == LegendViewName ||
                    v.Name.StartsWith(LegendViewName + "_"));
        }

        private static ElementId GetDraftingViewFamilyTypeId(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Drafting)?.Id;
        }

        private static ElementId GetDefaultTextNoteTypeId(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .FirstOrDefault()?.Id;
        }

        private static ElementId GetSolidFillPatternId(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill)?.Id;
        }

        private static void CleanupLegendRegionTypes(Document doc)
        {
            var legendTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .Where(t => t.Name.StartsWith(RegionTypePrefix))
                .ToList();

            foreach (var type in legendTypes)
            {
                try { doc.Delete(type.Id); } catch { }
            }
        }
    }
}

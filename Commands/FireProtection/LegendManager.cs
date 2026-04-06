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

        public static ElementId CreateLegendDraftingView(
            Document doc,
            List<FireProtectionTypeEntry> types,
            Dictionary<string, int> elementCounts,
            bool overwriteExisting,
            ElementId textNoteTypeId = null,
            bool includeColumnFrame = false,
            double columnA_feet = 0,
            double columnB_feet = 0)
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

                    double rowY = startY - rowPitch * i;

                    // 梁の色付き矩形（ビュー用タイプを再利用）
                    string viewTypeName = FilledRegionCreator.TypePrefix + entry.Name;
                    CreateColoredRectangle(doc, draftingView.Id,
                        solidFillPatternId, color, viewTypeName,
                        0, rowY, rectWidth, rectHeight);

                    XYZ textPos = new XYZ(textOffsetX, rowY + textYOffset, 0);
                    string label = $"- {entry.Name}";
                    TextNote.Create(doc, draftingView.Id, textPos, label, textNoteTypeId);
                }

                // 柱枠型の凡例行（平面/天伏ビュー用）
                if (includeColumnFrame && columnA_feet > 0 && columnB_feet > 0)
                {
                    double colStartY = startY - rowPitch * types.Count - textHeight * 2;

                    // 柱凡例タイトル
                    XYZ colTitlePos = new XYZ(0, colStartY, 0);
                    TextNote.Create(doc, draftingView.Id, colTitlePos,
                        "柱 耐火被覆", textNoteTypeId);
                    colStartY -= (textHeight + titleGap * 0.6);

                    for (int i = 0; i < types.Count; i++)
                    {
                        var entry = types[i];
                        Color color = new Color(entry.ColorR, entry.ColorG, entry.ColorB);

                        double rowY = colStartY - rowPitch * i;

                        // 枠型（外側矩形 + 内側穴）
                        string viewTypeName = FilledRegionCreator.TypePrefix + entry.Name;
                        double outerSize = rectHeight * 0.9;
                        double innerSize = outerSize * (columnA_feet / (columnA_feet + columnB_feet));
                        CreateFrameRectangle(doc, draftingView.Id,
                            solidFillPatternId, color, viewTypeName,
                            rectWidth / 2.0, rowY + rectHeight / 2.0,
                            outerSize / 2.0, innerSize / 2.0);

                        XYZ textPos = new XYZ(textOffsetX, rowY + textYOffset, 0);
                        string label = $"- {entry.Name}";
                        TextNote.Create(doc, draftingView.Id, textPos, label, textNoteTypeId);
                    }
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

        private static void CreateFrameRectangle(
            Document doc, ElementId viewId,
            ElementId solidFillPatternId, Color color,
            string typeName,
            double cx, double cy,
            double outerHalf, double innerHalf)
        {
            ElementId regionTypeId = GetOrCreateFilledRegionType(
                doc, solidFillPatternId, color, typeName);
            if (regionTypeId == null) return;

            // 外側（反時計回り）
            CurveLoop outer = new CurveLoop();
            XYZ o0 = new XYZ(cx - outerHalf, cy - outerHalf, 0);
            XYZ o1 = new XYZ(cx + outerHalf, cy - outerHalf, 0);
            XYZ o2 = new XYZ(cx + outerHalf, cy + outerHalf, 0);
            XYZ o3 = new XYZ(cx - outerHalf, cy + outerHalf, 0);
            outer.Append(Line.CreateBound(o0, o1));
            outer.Append(Line.CreateBound(o1, o2));
            outer.Append(Line.CreateBound(o2, o3));
            outer.Append(Line.CreateBound(o3, o0));

            // 内側（時計回り = 穴）
            CurveLoop inner = new CurveLoop();
            XYZ i0 = new XYZ(cx - innerHalf, cy - innerHalf, 0);
            XYZ i1 = new XYZ(cx + innerHalf, cy - innerHalf, 0);
            XYZ i2 = new XYZ(cx + innerHalf, cy + innerHalf, 0);
            XYZ i3 = new XYZ(cx - innerHalf, cy + innerHalf, 0);
            inner.Append(Line.CreateBound(i0, i1));
            inner.Append(Line.CreateBound(i1, i2));
            inner.Append(Line.CreateBound(i2, i3));
            inner.Append(Line.CreateBound(i3, i0));
            inner.Flip();

            FilledRegion.Create(doc, regionTypeId, viewId,
                new List<CurveLoop> { outer, inner });
        }

        private static ElementId GetOrCreateFilledRegionType(
            Document doc, ElementId solidFillPatternId,
            Color color, string typeName)
        {
            // ビューで作成済みのタイプがあればそのまま再利用
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
            // ビュー用タイプと共有のため、凡例独自のクリーンアップは不要
            // （FilledRegionCreatorのCleanupで処理済み）
        }
    }
}

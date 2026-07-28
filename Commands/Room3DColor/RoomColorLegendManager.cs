using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.Room3DColor
{
    /// <summary>
    /// 部屋3D色分けの凡例（製図ビュー）を作成
    /// </summary>
    public static class RoomColorLegendManager
    {
        private const string TypePrefix = "部屋3D色分け_";

        /// <summary>
        /// 凡例用の製図ビューを作成する。トランザクション内で呼ぶこと。
        /// </summary>
        public static ElementId CreateLegendDraftingView(
            Document doc,
            string legendViewName,
            List<RoomColorGroup> groups,
            ElementId textNoteTypeId = null)
        {
            try
            {
                // 同名の既存凡例ビューを削除
                var existingView = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewDrafting))
                    .Cast<ViewDrafting>()
                    .FirstOrDefault(v => v.Name == legendViewName ||
                        v.Name.StartsWith(legendViewName + "_"));
                if (existingView != null)
                {
                    try { doc.Delete(existingView.Id); } catch { }
                }

                CleanupLegendRegionTypes(doc);

                ElementId viewFamilyTypeId = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Drafting)?.Id;
                if (viewFamilyTypeId == null)
                    return null;

                ViewDrafting draftingView = ViewDrafting.Create(doc, viewFamilyTypeId);
                try { draftingView.Name = legendViewName; }
                catch { draftingView.Name = legendViewName + "_" +
                    Guid.NewGuid().ToString("N").Substring(0, 6); }

                draftingView.Scale = 1;

                ElementId solidFillPatternId = RoomColorManager.GetSolidFillPatternId(doc);
                if (solidFillPatternId == null)
                    return draftingView.Id;

                if (textNoteTypeId == null)
                    textNoteTypeId = new FilteredElementCollector(doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .FirstOrDefault()?.Id;
                if (textNoteTypeId == null)
                    return draftingView.Id;

                double textHeight = GetTextHeight(doc, textNoteTypeId);
                double rectHeight = textHeight * 1.8;
                double rectWidth = textHeight * 4.0;
                double textOffsetX = rectWidth + textHeight * 0.8;
                double rowSpacing = textHeight * 0.8;
                double titleGap = textHeight * 3.0;
                double textYOffset = (rectHeight + textHeight) / 2 + textHeight * 0.3;
                double rowPitch = rectHeight + rowSpacing;

                double startY = 0;

                // タイトル
                TextNote.Create(doc, draftingView.Id, new XYZ(0, startY, 0),
                    Loc.S("Room3D.LegendTitle"), textNoteTypeId);
                startY -= (textHeight + titleGap);

                for (int i = 0; i < groups.Count; i++)
                {
                    var g = groups[i];
                    double rowY = startY - rowPitch * i;
                    string typeName = TypePrefix + (string.IsNullOrEmpty(g.Label) ? i.ToString() : g.Label);

                    CreateColoredRectangle(doc, draftingView.Id,
                        solidFillPatternId, g.Color, typeName,
                        0, rowY, rectWidth, rectHeight);

                    XYZ textPos = new XYZ(textOffsetX, rowY + textYOffset, 0);
                    string label = string.Format(Loc.S("Room3D.LegendRow"), g.Label, g.RoomIds.Count);
                    TextNote.Create(doc, draftingView.Id, textPos, label, textNoteTypeId);
                }

                return draftingView.Id;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void CreateColoredRectangle(
            Document doc, ElementId viewId,
            ElementId solidFillPatternId, Color color, string typeName,
            double x, double y, double width, double height)
        {
            ElementId regionTypeId = GetOrCreateFilledRegionType(
                doc, solidFillPatternId, color, typeName);
            if (regionTypeId == null)
                return;

            XYZ p0 = new XYZ(x, y, 0);
            XYZ p1 = new XYZ(x + width, y, 0);
            XYZ p2 = new XYZ(x + width, y + height, 0);
            XYZ p3 = new XYZ(x, y + height, 0);

            CurveLoop loop = new CurveLoop();
            loop.Append(Line.CreateBound(p0, p1));
            loop.Append(Line.CreateBound(p1, p2));
            loop.Append(Line.CreateBound(p2, p3));
            loop.Append(Line.CreateBound(p3, p0));

            FilledRegion.Create(doc, regionTypeId, viewId, new List<CurveLoop> { loop });
        }

        private static ElementId GetOrCreateFilledRegionType(
            Document doc, ElementId solidFillPatternId, Color color, string typeName)
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
            if (baseType == null)
                return null;

            var newType = baseType.Duplicate(typeName) as FilledRegionType;
            if (newType == null)
                return null;

            newType.ForegroundPatternColor = color;
            newType.ForegroundPatternId = solidFillPatternId;
            newType.BackgroundPatternId = ElementId.InvalidElementId;
            newType.IsMasking = false;

            return newType.Id;
        }

        private static double GetTextHeight(Document doc, ElementId textNoteTypeId)
        {
            var textNoteType = doc.GetElement(textNoteTypeId) as TextNoteType;
            Parameter sizeParam = textNoteType?.get_Parameter(BuiltInParameter.TEXT_SIZE);
            if (sizeParam != null)
                return sizeParam.AsDouble();
            return 2.5 / 304.8;
        }

        private static void CleanupLegendRegionTypes(Document doc)
        {
            var legendTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .Where(t => t.Name.StartsWith(TypePrefix))
                .ToList();
            foreach (var type in legendTypes)
            {
                try { doc.Delete(type.Id); } catch { }
            }
        }
    }
}

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
        private const string LegendViewName = "耐火被覆色分け凡例"; // legend

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
                double mmToFt = 1.0 / 304.8;
                double colDivX = 20.0 * mmToFt;
                double frameWidth = colDivX + 65.0 * mmToFt;
                double rowH = 10.0 * mmToFt;
                double pad = 1.0 * mmToFt;                // 四周1mm余白
                double rectW = colDivX - pad * 2;          // 色四角幅 = 20mm - 2mm = 18mm
                double rectH = rowH - pad * 2;             // 色四角高さ = 10mm - 2mm = 8mm
                double textOffsetX = colDivX + pad;
                double titleGap = textHeight * 1.2;
                double colFrameThick = 2.0 * mmToFt;      // 柱塗り潰し厚み2mm

                double curY = 0;

                // タイトル: ◎耐火被覆仕様凡例
                TextNote.Create(doc, draftingView.Id, new XYZ(0, curY, 0),
                    "\u25ce\u8010\u706b\u88ab\u8986\u4ed5\u69d8\u51e1\u4f8b", textNoteTypeId);
                curY -= (textHeight + titleGap);

                // 全行数を計算
                int totalRows = types.Count;
                if (includeColumnFrame && columnA_feet > 0 && columnB_feet > 0)
                    totalRows += types.Count;

                double tableTop = curY;
                double tableLeft = 0;
                double tableRight = frameWidth;

                // 表の外枠（上線）
                var vw = doc.GetElement(draftingView.Id) as View;

                // 梁の行
                for (int i = 0; i < types.Count; i++)
                {
                    var entry = types[i];
                    Color color = new Color(entry.ColorR, entry.ColorG, entry.ColorB);
                    double rowTop = curY;
                    double rowBottom = curY - rowH;

                    string viewTypeName = FilledRegionCreator.TypePrefix + entry.Name;
                    CreateColoredRectangle(doc, draftingView.Id,
                        solidFillPatternId, color, viewTypeName,
                        pad, rowBottom + pad, rectW, rectH);

                    var tn = TextNote.Create(doc, draftingView.Id,
                        new XYZ(textOffsetX, rowBottom + rowH / 2.0, 0),
                        entry.Name, textNoteTypeId);
                    tn.VerticalAlignment = VerticalTextAlignment.Middle;

                    curY = rowBottom;
                }

                // 柱枠型の行
                if (includeColumnFrame && columnA_feet > 0 && columnB_feet > 0)
                {
                    for (int i = 0; i < types.Count; i++)
                    {
                        var entry = types[i];
                        double rowTop = curY;
                        double rowBottom = curY - rowH;

                        string colTypeName = FilledRegionCreator.TypePrefix + "柱_" + entry.Name;
                        double outerHalf = rectH / 2.0;
                        double innerHalf = outerHalf - colFrameThick; // 厚み2mm
                        Color colColor = new Color(entry.ColColorR, entry.ColColorG, entry.ColColorB);
                        CreateFrameRectangle(doc, draftingView.Id,
                            solidFillPatternId, colColor, colTypeName,
                            pad + rectW / 2.0, rowBottom + pad + rectH / 2.0,
                            outerHalf, innerHalf);

                        var ctn = TextNote.Create(doc, draftingView.Id,
                            new XYZ(textOffsetX, rowBottom + rowH / 2.0, 0),
                            $"\u67f1\uff1a{entry.Name}", textNoteTypeId);
                        ctn.VerticalAlignment = VerticalTextAlignment.Middle;

                        curY = rowBottom;
                    }
                }

                double tableBottom = curY;

                // 表のグリッド線（Excelセル表形式）
                if (vw != null)
                {
                    try
                    {
                        // 横線（各行の境界 + 上端 + 下端）
                        double y = tableTop;
                        for (int i = 0; i <= totalRows; i++)
                        {
                            doc.Create.NewDetailCurve(vw, Line.CreateBound(
                                new XYZ(tableLeft, y, 0),
                                new XYZ(tableRight, y, 0)));
                            y -= rowH;
                        }

                        // 縦線（左端、色四角列右端、右端）
                        doc.Create.NewDetailCurve(vw, Line.CreateBound(
                            new XYZ(tableLeft, tableTop, 0),
                            new XYZ(tableLeft, tableBottom, 0)));
                        doc.Create.NewDetailCurve(vw, Line.CreateBound(
                            new XYZ(colDivX, tableTop, 0),
                            new XYZ(colDivX, tableBottom, 0)));
                        doc.Create.NewDetailCurve(vw, Line.CreateBound(
                            new XYZ(tableRight, tableTop, 0),
                            new XYZ(tableRight, tableBottom, 0)));
                    }
                    catch { }
                }

                // 注記セクション（※毎に1つのTextNote）
                curY -= textHeight * 0.5; // 囲い線と定型文の間隔

                var noteBlocks = new[]
                {
                    "\u203b\u8010\u706b\u88ab\u8986\u4e0d\u8981\u7bc4\u56f2\n\u3000\u30fb\u8010\u98a8\u6881\u3001ALC\u30fbECP\u958b\u53e3\u88dc\u5f37\n\u3000\u30fb\u5e8a\u3092\u53d7\u3051\u306a\u3044EV\u6881\u30fb\u9593\u67f1\n\u3000\u30fb\u968e\u6bb5\u53d7\u3051\u6881\u30fb\u67f1\n\u3000\u30fb\u6c34\u5e73\u30d6\u30ec\u30fc\u30b9",
                    "\u203b\u534a\u6e7f\u5f0f\u5439\u4ed8\u30ed\u30c3\u30af\u30a6\u30fc\u30eb\u5de5\u6cd5\u306b\u5bfe\u3059\u308b\u30bb\u30e1\u30f3\u30c8\u30b9\u30e9\u30ea\u30fc\u5439\u4ed8\n\u3000\u306b\u3088\u308b\u8868\u9762\u786c\u5316\u51e6\u7406\u306f\u7bc4\u56f2\n\u3000\u30fbEV\u30b7\u30e3\u30d5\u30c8\u5185\n\u3000\u30fbEV\u6a5f\u68b0\u5ba4\n\u3000\u30fb\u96fb\u6c17\u5ba4\n\u3000\u30fb\u30b5\u30fc\u30d0\u30fc\u5ba4\u7b49\n\u3000\u30fb\u76f4\u5929\u4e95\u5ba4\u5185\u5916\u90e8\uff08\u30d4\u30ed\u30c6\u30a3\uff09",
                    "\u203b\u30b5\u30d3\u6b62\u3081\u7bc4\u56f2\u8a73\u7d30\u306f\u3001\u5225\u9014\u300c**\u9244\u5de5\u3000\u5857\u88c5\u7bc4\u56f2\u56f3\u300d\u53c2\u7167",
                };

                foreach (var block in noteBlocks)
                {
                    TextNote.Create(doc, draftingView.Id,
                        new XYZ(0, curY, 0), block, textNoteTypeId);
                    int lineCount = block.Split('\n').Length;
                    curY -= textHeight * 1.6 * lineCount + textHeight * 2.5;
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

            var fr = FilledRegion.Create(doc, regionTypeId, viewId,
                new List<CurveLoop> { outer, inner });

            // 柱枠の境界線を非表示に
            try
            {
                var invisStyle = new FilteredElementCollector(doc)
                    .OfClass(typeof(GraphicsStyle))
                    .Cast<GraphicsStyle>()
                    .FirstOrDefault(gs =>
                        gs.GraphicsStyleType == GraphicsStyleType.Projection &&
                        (gs.Name.Contains("非表示") || gs.Name.Contains("Invisible")));
                if (invisStyle != null)
                    fr.SetLineStyleId(invisStyle.Id);
            }
            catch { }
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

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.BeamCommon;

namespace Tools28.Commands.BeamTopLevel
{
    /// <summary>
    /// 梁天端色分けの凡例（凡例ビュー）を作成
    /// </summary>
    public static class LegendManager
    {
        internal const string LegendViewName = "梁天端色分け凡例";
        private const string ErrorFilterName = "梁天端_エラー";

        /// <summary>
        /// 階ごとの凡例ビュー名（例: "梁天端色分け凡例_1FL"）。再実行時の検索キーなので多言語化しない
        /// </summary>
        internal static string GetLegendViewName(string labelText)
        {
            return LegendViewName + "_" + labelText;
        }

        /// <summary>
        /// 凡例を作成する（凡例ビュー。モデルに凡例ビューが無い場合は製図ビュー）
        /// </summary>
        public static BeamLegendResult CreateLegendView(
            Document doc,
            Dictionary<string, int> levelGroups,
            bool overwriteExisting,
            int errorCount,
            ElementId textNoteTypeId,
            string labelText)
        {
            BeamLegendResult legend = null;
            try
            {
                string viewName = GetLegendViewName(labelText);
                legend = BeamLegendViews.Prepare(doc, viewName, overwriteExisting);
                View legendView = legend.View;
                if (legendView == null)
                    return legend;

                try { legendView.Scale = 1; } catch (Exception) { }

                ElementId solidFillPatternId = GetSolidFillPatternId(doc);
                if (solidFillPatternId == null)
                    return legend;

                var sortedLevels = levelGroups
                    .OrderBy(kv => FilterManager.ExtractNumericValue(kv.Key))
                    .ToList();

                // フィルタと同じ色（レベル差で固定）
                Dictionary<string, Color> colors = FilterManager.ResolveColors(doc, levelGroups.Keys);

                if (textNoteTypeId == null)
                    textNoteTypeId = GetDefaultTextNoteTypeId(doc);
                if (textNoteTypeId == null)
                    return legend;

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
                XYZ titlePos = new XYZ(0, startY, 0);
                TextNote.Create(doc, legendView.Id, titlePos,
                    viewName, textNoteTypeId);
                startY -= (textHeight + titleGap);

                // 各レベル行
                for (int i = 0; i < sortedLevels.Count; i++)
                {
                    string displayValue = sortedLevels[i].Key;
                    int beamCount = sortedLevels[i].Value;
                    Color color = colors[displayValue];
                    // 色の保存先のタイプを使う（各階の凡例で共通。削除しない）
                    string typeName = BeamLevelColors.ColorTypeName(FilterManager.ColorTypePrefix, displayValue);

                    double rowY = startY - rowPitch * i;

                    CreateColoredRectangle(doc, legendView.Id,
                        solidFillPatternId, color, typeName,
                        0, rowY, rectWidth, rectHeight);

                    XYZ textPos = new XYZ(textOffsetX, rowY + textYOffset, 0);
                    string label = $"- {displayValue}（{beamCount}本）";
                    TextNote.Create(doc, legendView.Id, textPos,
                        label, textNoteTypeId);
                }

                // エラー行
                if (errorCount > 0)
                {
                    double errorRowY = startY - rowPitch * sortedLevels.Count;

                    CreateColoredRectangle(doc, legendView.Id,
                        solidFillPatternId, new Color(255, 100, 100),
                        ErrorFilterName,
                        0, errorRowY, rectWidth, rectHeight);

                    XYZ errorTextPos = new XYZ(textOffsetX, errorRowY + textYOffset, 0);
                    TextNote.Create(doc, legendView.Id, errorTextPos,
                        "- エラー", textNoteTypeId);
                }

                return legend;
            }
            catch (Exception ex)
            {
                // 失敗の理由を完了画面とログに出す（黙って失敗すると、何が起きたか分からないため）
                DiagLog.Write($"[BeamTop Legend] 凡例の作成に失敗: {ex}");
                if (legend == null)
                    legend = new BeamLegendResult();
                legend.Error = ex.Message;
                return legend;
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

            var loops = new List<CurveLoop> { loop };
            FilledRegion.Create(doc, regionTypeId, viewId, loops);
        }

        private static ElementId GetOrCreateFilledRegionType(
            Document doc, ElementId solidFillPatternId,
            Color color, string typeName)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault(t => t.Name == typeName);

            // 他の階の凡例でも使っているため削除せず、そのまま使う
            if (existing != null)
                return existing.Id;

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
            if (textNoteType != null)
            {
                Parameter sizeParam = textNoteType.get_Parameter(
                    BuiltInParameter.TEXT_SIZE);
                if (sizeParam != null)
                    return sizeParam.AsDouble();
            }
            return 2.5 / 304.8;
        }

        private static ElementId GetDefaultTextNoteTypeId(Document doc)
        {
            var textNoteType = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .FirstOrDefault();

            return textNoteType?.Id;
        }

        private static ElementId GetSolidFillPatternId(Document doc)
        {
            var fillPattern = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);

            return fillPattern?.Id;
        }
    }
}

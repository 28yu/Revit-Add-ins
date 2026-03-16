using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamUnderLevel
{
    /// <summary>
    /// 梁下端色分け凡例の製図ビューを作成
    /// </summary>
    public static class LegendManager
    {
        private const string LegendViewName = "梁下端色分け凡例";
        // フィルタ名と同じプレフィックス
        private const string FilterPrefix = "梁下_";
        private const string ErrorFilterName = "梁下_エラー";

        /// <summary>
        /// 凡例用の製図ビューを作成
        /// </summary>
        public static ElementId CreateLegendDraftingView(
            Document doc,
            Dictionary<string, int> levelGroups,
            bool overwriteExisting)
        {
            try
            {
                // 既存の凡例ビューを確認
                ViewDrafting existingView = FindExistingLegendView(doc);
                if (existingView != null)
                {
                    if (overwriteExisting)
                    {
                        doc.Delete(existingView.Id);
                    }
                    else
                    {
                        return null;
                    }
                }

                // 既存の凡例用FilledRegionTypeをクリーンアップ
                if (overwriteExisting)
                {
                    CleanupLegendRegionTypes(doc);
                }

                // 製図ビュー用の ViewFamilyType を取得
                ElementId viewFamilyTypeId = GetDraftingViewFamilyTypeId(doc);
                if (viewFamilyTypeId == null)
                    return null;

                // 製図ビューを作成
                ViewDrafting draftingView = ViewDrafting.Create(doc, viewFamilyTypeId);
                try
                {
                    draftingView.Name = LegendViewName;
                }
                catch
                {
                    draftingView.Name = LegendViewName + "_" +
                        DateTime.Now.ToString("yyyyMMdd_HHmmss");
                }

                // 尺度を1:1に設定
                draftingView.Scale = 1;

                // ベタ塗りパターンを取得
                ElementId solidFillPatternId = GetSolidFillPatternId(doc);
                if (solidFillPatternId == null)
                    return draftingView.Id;

                // レベル値でソート（FilterManagerと同じロジック）
                var sortedLevels = levelGroups
                    .OrderBy(kv => FilterManager.ExtractNumericValue(kv.Key))
                    .ToList();

                // FilterManagerと同じ色を生成
                List<Color> colors = FilterManager.GenerateColors(sortedLevels.Count);

                // TextNoteTypeを取得し、テキストサイズを基準にレイアウト計算
                ElementId textNoteTypeId = GetDefaultTextNoteTypeId(doc);
                if (textNoteTypeId == null)
                    return draftingView.Id;

                double textHeight = GetTextHeight(doc, textNoteTypeId);

                // テキストサイズ基準のレイアウト定数
                double rectHeight = textHeight * 1.8;
                double rectWidth = textHeight * 4.0;
                double textOffsetX = rectWidth + textHeight * 0.8;
                double rowSpacing = textHeight * 0.8;
                double titleGap = textHeight * 3.0;

                // 凡例の各行を描画
                double currentY = 0;

                // タイトル
                XYZ titlePos = new XYZ(0, currentY, 0);
                TextNote.Create(doc, draftingView.Id, titlePos,
                    "梁下端色分け凡例", textNoteTypeId);
                currentY -= (textHeight + titleGap);

                // 各レベル行
                for (int i = 0; i < sortedLevels.Count; i++)
                {
                    string displayValue = sortedLevels[i].Key;
                    int beamCount = sortedLevels[i].Value;
                    Color color = colors[i];
                    string typeName = FilterPrefix + displayValue;

                    // 色付き矩形を作成
                    CreateColoredRectangle(doc, draftingView.Id,
                        solidFillPatternId, color, typeName,
                        0, currentY, rectWidth, rectHeight);

                    // テキストラベル（矩形の右横、垂直中央揃え）
                    // TextNoteの基準点は左上、矩形中央にテキスト中央を合わせる
                    XYZ textPos = new XYZ(
                        textOffsetX,
                        currentY + rectHeight / 2 + textHeight * 0.7,
                        0);
                    string label = $"- {displayValue}（{beamCount}本）";
                    TextNote.Create(doc, draftingView.Id, textPos,
                        label, textNoteTypeId);

                    currentY -= (rectHeight + rowSpacing);
                }

                // エラー行
                CreateColoredRectangle(doc, draftingView.Id,
                    solidFillPatternId, new Color(255, 100, 100),
                    ErrorFilterName,
                    0, currentY, rectWidth, rectHeight);

                XYZ errorTextPos = new XYZ(
                    textOffsetX,
                    currentY + rectHeight / 2 + textHeight * 0.4,
                    0);
                TextNote.Create(doc, draftingView.Id, errorTextPos,
                    "- エラー", textNoteTypeId);

                return draftingView.Id;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 色付き矩形（FilledRegion）を作成
        /// </summary>
        private static void CreateColoredRectangle(
            Document doc, ElementId viewId,
            ElementId solidFillPatternId, Color color,
            string typeName,
            double x, double y, double width, double height)
        {
            // FilledRegionTypeを作成（フィルタ名と同じ名前）
            ElementId regionTypeId = GetOrCreateFilledRegionType(
                doc, solidFillPatternId, color, typeName);
            if (regionTypeId == null)
                return;

            // 矩形の輪郭を定義
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

        /// <summary>
        /// 指定色のFilledRegionTypeを取得または作成
        /// タイプ名はフィルタ名と同じ（例: 梁下_2FL+2800）
        /// </summary>
        private static ElementId GetOrCreateFilledRegionType(
            Document doc, ElementId solidFillPatternId,
            Color color, string typeName)
        {
            // 既存タイプを検索
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault(t => t.Name == typeName);

            if (existing != null)
            {
                doc.Delete(existing.Id);
            }

            // ベースとなるFilledRegionTypeを取得
            var baseType = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault();

            if (baseType == null)
                return null;

            // 複製して色を設定
            var newType = baseType.Duplicate(typeName) as FilledRegionType;
            if (newType == null)
                return null;

            newType.ForegroundPatternColor = color;
            newType.ForegroundPatternId = solidFillPatternId;
            newType.BackgroundPatternId = ElementId.InvalidElementId;
            newType.IsMasking = false;

            return newType.Id;
        }

        /// <summary>
        /// TextNoteTypeからテキスト高さを取得（内部単位 feet）
        /// </summary>
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
            // フォールバック: 2.5mm
            return 2.5 / 304.8;
        }

        /// <summary>
        /// 既存の凡例製図ビューを検索
        /// </summary>
        private static ViewDrafting FindExistingLegendView(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewDrafting))
                .Cast<ViewDrafting>()
                .FirstOrDefault(v => v.Name == LegendViewName ||
                    v.Name.StartsWith(LegendViewName + "_"));
        }

        /// <summary>
        /// 製図ビューのViewFamilyTypeIdを取得
        /// </summary>
        private static ElementId GetDraftingViewFamilyTypeId(Document doc)
        {
            var viewFamilyType = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Drafting);

            return viewFamilyType?.Id;
        }

        /// <summary>
        /// デフォルトのTextNoteTypeIdを取得
        /// </summary>
        private static ElementId GetDefaultTextNoteTypeId(Document doc)
        {
            var textNoteType = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .FirstOrDefault();

            return textNoteType?.Id;
        }

        /// <summary>
        /// ベタ塗りパターンのElementIdを取得
        /// </summary>
        private static ElementId GetSolidFillPatternId(Document doc)
        {
            var fillPattern = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);

            return fillPattern?.Id;
        }

        /// <summary>
        /// 凡例用のFilledRegionTypeをクリーンアップ
        /// </summary>
        private static void CleanupLegendRegionTypes(Document doc)
        {
            var legendTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .Where(t => t.Name.StartsWith(FilterPrefix))
                .ToList();

            foreach (var type in legendTypes)
            {
                try { doc.Delete(type.Id); } catch { }
            }
        }
    }
}

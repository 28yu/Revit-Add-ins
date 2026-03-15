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
        private const string LegendTypePrefix = "梁下_凡例_";

        // レイアウト定数（mm → 内部単位はfeetだがMmToFeetで変換）
        private const double RectWidthMm = 15.0;
        private const double RectHeightMm = 8.0;
        private const double TextOffsetXMm = 20.0;
        private const double RowSpacingMm = 12.0;
        private const double TitleOffsetYMm = 5.0;
        private const double StartXMm = 0.0;
        private const double StartYMm = 0.0;

        /// <summary>
        /// 凡例用の製図ビューを作成
        /// </summary>
        /// <param name="doc">Revitドキュメント</param>
        /// <param name="levelGroups">レベル別グループ（表示値→梁数）</param>
        /// <param name="overwriteExisting">既存の凡例ビューを上書きするか</param>
        /// <returns>作成された製図ビューのId（失敗時はnull）</returns>
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

                // 製図ビュー用の ViewFamilyType を取得
                ElementId viewFamilyTypeId = GetDraftingViewFamilyTypeId(doc);
                if (viewFamilyTypeId == null)
                    return null;

                // 製図ビューを作成
                ViewDrafting draftingView = ViewDrafting.Create(doc, viewFamilyTypeId);
                // 名前の重複を避ける
                try
                {
                    draftingView.Name = LegendViewName;
                }
                catch
                {
                    draftingView.Name = LegendViewName + "_" +
                        DateTime.Now.ToString("yyyyMMdd_HHmmss");
                }

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

                // TextNoteTypeを取得
                ElementId textNoteTypeId = GetDefaultTextNoteTypeId(doc);
                if (textNoteTypeId == null)
                    return draftingView.Id;

                // 凡例の各行を描画
                double currentY = MmToFeet(StartYMm);

                // タイトル
                XYZ titlePos = new XYZ(MmToFeet(StartXMm), currentY, 0);
                TextNote.Create(doc, draftingView.Id, titlePos,
                    "梁下端色分け凡例", textNoteTypeId);
                currentY -= MmToFeet(TitleOffsetYMm + RowSpacingMm);

                // 各レベル行
                for (int i = 0; i < sortedLevels.Count; i++)
                {
                    string displayValue = sortedLevels[i].Key;
                    int beamCount = sortedLevels[i].Value;
                    Color color = colors[i];

                    // 色付き矩形を作成
                    CreateColoredRectangle(doc, draftingView.Id,
                        solidFillPatternId, color,
                        MmToFeet(StartXMm), currentY, i);

                    // テキストラベル
                    XYZ textPos = new XYZ(
                        MmToFeet(StartXMm + TextOffsetXMm),
                        currentY + MmToFeet(RectHeightMm / 2),
                        0);
                    string label = $"{displayValue}  ({beamCount}本)";
                    TextNote.Create(doc, draftingView.Id, textPos,
                        label, textNoteTypeId);

                    currentY -= MmToFeet(RowSpacingMm + RectHeightMm);
                }

                // エラー行
                CreateColoredRectangle(doc, draftingView.Id,
                    solidFillPatternId, new Color(255, 100, 100),
                    MmToFeet(StartXMm), currentY, sortedLevels.Count);

                XYZ errorTextPos = new XYZ(
                    MmToFeet(StartXMm + TextOffsetXMm),
                    currentY + MmToFeet(RectHeightMm / 2),
                    0);
                TextNote.Create(doc, draftingView.Id, errorTextPos,
                    "エラー", textNoteTypeId);

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
            double x, double y, int colorIndex)
        {
            double width = MmToFeet(RectWidthMm);
            double height = MmToFeet(RectHeightMm);

            // FilledRegionTypeを作成（色ごと）
            ElementId regionTypeId = GetOrCreateFilledRegionType(
                doc, solidFillPatternId, color, colorIndex);
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
        /// </summary>
        private static ElementId GetOrCreateFilledRegionType(
            Document doc, ElementId solidFillPatternId,
            Color color, int index)
        {
            string typeName = LegendTypePrefix + index;

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
            // 背景パターンをクリア（ベタ塗りのみ）
            newType.BackgroundPatternId = ElementId.InvalidElementId;
            // 境界線を非表示
            newType.IsMasking = false;

            return newType.Id;
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
        /// mm → feet 変換
        /// </summary>
        private static double MmToFeet(double mm)
        {
            return mm / 304.8;
        }

        /// <summary>
        /// 凡例作成時に生成したFilledRegionTypeをクリーンアップ
        /// </summary>
        public static void CleanupLegendTypes(Document doc)
        {
            var legendTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .Where(t => t.Name.StartsWith(LegendTypePrefix))
                .ToList();

            foreach (var type in legendTypes)
            {
                try { doc.Delete(type.Id); } catch { }
            }
        }
    }
}

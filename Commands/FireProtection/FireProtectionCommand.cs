using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.FireProtection
{
    /// <summary>
    /// 耐火被覆範囲色分け図コマンド（梁・柱の耐火被覆を色分け表示）
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FireProtectionCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                View activeView = doc.ActiveView;

                // シートビュー: 配置された全ビューを処理対象にする
                List<View> targetViews = new List<View>();
                bool isSheet = activeView.ViewType == ViewType.DrawingSheet;

                if (isSheet)
                {
                    var sheet = activeView as ViewSheet;
                    if (sheet != null)
                    {
                        foreach (ElementId vpId in sheet.GetAllViewports())
                        {
                            var vp = doc.GetElement(vpId) as Viewport;
                            if (vp == null) continue;
                            var v = doc.GetElement(vp.ViewId) as View;
                            if (v == null) continue;

                            if (v.ViewType == ViewType.FloorPlan ||
                                v.ViewType == ViewType.CeilingPlan ||
                                v.ViewType == ViewType.EngineeringPlan ||
                                v.ViewType == ViewType.Section)
                            {
                                targetViews.Add(v);
                            }
                        }
                    }

                    if (targetViews.Count == 0)
                    {
                        TaskDialog.Show("エラー",
                            "シート上に対応するビュー（平面、天伏、構造伏、断面）がありません。");
                        return Result.Cancelled;
                    }
                }
                else if (activeView.ViewType == ViewType.FloorPlan ||
                         activeView.ViewType == ViewType.CeilingPlan ||
                         activeView.ViewType == ViewType.EngineeringPlan ||
                         activeView.ViewType == ViewType.Section)
                {
                    targetViews.Add(activeView);
                }
                else
                {
                    TaskDialog.Show("エラー",
                        "平面ビュー、天井伏図、構造伏図、断面図、またはシートで実行してください。");
                    return Result.Cancelled;
                }

                // 全ビューの要素を収集してダイアログ用データを構築
                var allBeams = new List<Element>();
                var allColumns = new List<Element>();
                bool hasSectionView = false;

                foreach (var v in targetViews)
                {
                    var vBeams = new FilteredElementCollector(doc, v.Id)
                        .OfCategory(BuiltInCategory.OST_StructuralFraming)
                        .WhereElementIsNotElementType()
                        .Cast<Element>().ToList();
                    var vColumns = new FilteredElementCollector(doc, v.Id)
                        .OfCategory(BuiltInCategory.OST_StructuralColumns)
                        .WhereElementIsNotElementType()
                        .Cast<Element>().ToList();

                    allBeams.AddRange(vBeams);
                    allColumns.AddRange(vColumns);

                    if (v.ViewType == ViewType.Section)
                        hasSectionView = true;
                }

                // 重複要素を除去（複数ビューに同じ要素が表示される場合）
                allBeams = allBeams.GroupBy(e => e.Id.IntegerValue)
                    .Select(g => g.First()).ToList();
                allColumns = allColumns.GroupBy(e => e.Id.IntegerValue)
                    .Select(g => g.First()).ToList();

                if (allBeams.Count == 0 && allColumns.Count == 0)
                {
                    TaskDialog.Show("エラー",
                        "ビュー内に梁（構造フレーム）または柱（構造柱）が見つかりません。");
                    return Result.Cancelled;
                }

                // ビューテンプレート確認（各ビュー）
                using (Transaction tpl = new Transaction(doc, "ビューテンプレート解除"))
                {
                    bool needTemplate = targetViews.Any(
                        v => v.ViewTemplateId != ElementId.InvalidElementId);

                    if (needTemplate)
                    {
                        TaskDialogResult templateResult = TaskDialog.Show(
                            "ビューテンプレート確認",
                            "ビューテンプレートが設定されているビューがあります。\n" +
                            "塗潰領域を配置するにはテンプレートを解除する必要があります。\n\n" +
                            "テンプレートを解除しますか？",
                            TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                        if (templateResult == TaskDialogResult.Yes)
                        {
                            tpl.Start();
                            foreach (var v in targetViews)
                            {
                                if (v.ViewTemplateId != ElementId.InvalidElementId)
                                    v.ViewTemplateId = ElementId.InvalidElementId;
                            }
                            tpl.Commit();
                        }
                        else
                        {
                            return Result.Cancelled;
                        }
                    }
                }

                // パラメータ自動検出
                var beamParams = BeamGeometryHelper.DetectFireProtectionParameters(allBeams);
                var columnParams = BeamGeometryHelper.DetectFireProtectionParameters(allColumns);

                // 線種取得
                var lineStyles = CollectLineStyles(doc);

                // 塗りパターン取得
                var fillPatterns = new FilteredElementCollector(doc)
                    .OfClass(typeof(FillPatternElement))
                    .Cast<FillPatternElement>()
                    .OrderBy(fp => fp.Name)
                    .Select(fp => new FillPatternItem { Id = fp.Id, Name = fp.Name })
                    .ToList();

                // 文字タイプ取得
                var textNoteTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .OrderBy(t => t.Name)
                    .Select(t => new FpTextNoteTypeItem(t))
                    .ToList();

                string viewTypeName;
                if (isSheet)
                    viewTypeName = $"シート（{targetViews.Count}ビュー）";
                else
                {
                    switch (activeView.ViewType)
                    {
                        case ViewType.FloorPlan: viewTypeName = "平面ビュー"; break;
                        case ViewType.CeilingPlan: viewTypeName = "天井伏図"; break;
                        case ViewType.EngineeringPlan: viewTypeName = "構造伏図"; break;
                        case ViewType.Section: viewTypeName = "断面図"; break;
                        default: viewTypeName = activeView.ViewType.ToString(); break;
                    }
                }

                bool onlySections = targetViews.All(v => v.ViewType == ViewType.Section);

                var dialogData = new FireProtectionDialogData
                {
                    ViewName = activeView.Name,
                    ViewTypeName = viewTypeName,
                    IsSectionView = onlySections,
                    BeamCount = allBeams.Count,
                    ColumnCount = allColumns.Count,
                    HasBeams = allBeams.Count > 0,
                    HasColumns = allColumns.Count > 0,
                    BeamParameters = beamParams,
                    ColumnParameters = columnParams,
                    TextNoteTypes = textNoteTypes,
                    LineStyles = lineStyles,
                    FillPatterns = fillPatterns
                };

                var dialog = new FireProtectionDialog(dialogData);
                if (dialog.ShowDialog() != true)
                    return Result.Cancelled;

                var settings = dialog.GetResult();

                string paramName = settings.SelectedParameterName;
                if (paramName != null && paramName.Contains("（"))
                    paramName = paramName.Substring(0, paramName.IndexOf("（"));

                // オフセット辞書
                var offsetByType = new Dictionary<string, double>();
                foreach (var typeEntry in settings.Types)
                {
                    double offsetMm = settings.UseCommonOffset
                        ? settings.CommonOffsetMm
                        : typeEntry.OffsetMm;
                    offsetByType[typeEntry.Name] = offsetMm / 304.8;
                }

                // 実行: 各ビューに対して処理
                ElementId legendViewId = null;
                int regionCount = 0;

                using (Transaction trans = new Transaction(doc, "耐火被覆色分け"))
                {
                    trans.Start();

                    try
                    {
                        // Type削除は全ビュー処理前に1回だけ
                        if (settings.OverwriteExisting)
                        {
                            FilledRegionCreator.CleanupExistingTypes(doc);
                        }

                        foreach (var view in targetViews)
                        {
                            // このビューの要素を収集
                            var viewBeams = new FilteredElementCollector(doc, view.Id)
                                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                                .WhereElementIsNotElementType()
                                .Cast<Element>().ToList();
                            var viewColumns = new FilteredElementCollector(doc, view.Id)
                                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                                .WhereElementIsNotElementType()
                                .Cast<Element>().ToList();

                            // 梁は常に対象。柱は断面ビューでは別途クリップ処理
                            var viewTargets = new List<Element>();

                            if (view.ViewType == ViewType.Section)
                            {
                                // 断面ビュー: 直交する梁（断面が見える梁）を除外
                                // ビュー座標でX幅 < Y高さの梁は直交梁
                                Transform vt = null;
                                try { vt = view.CropBox.Transform.Inverse; } catch { }

                                foreach (var beam in viewBeams)
                                {
                                    if (vt == null) { viewTargets.Add(beam); continue; }

                                    BoundingBoxXYZ bbb = beam.get_BoundingBox(view)
                                        ?? beam.get_BoundingBox(null);
                                    if (bbb == null) { viewTargets.Add(beam); continue; }

                                    // ビュー座標に変換してX/Y範囲を比較
                                    double vMinX = double.MaxValue, vMaxX = double.MinValue;
                                    double vMinY = double.MaxValue, vMaxY = double.MinValue;
                                    for (int ix = 0; ix <= 1; ix++)
                                    for (int iy = 0; iy <= 1; iy++)
                                    for (int iz = 0; iz <= 1; iz++)
                                    {
                                        XYZ c = new XYZ(
                                            ix == 0 ? bbb.Min.X : bbb.Max.X,
                                            iy == 0 ? bbb.Min.Y : bbb.Max.Y,
                                            iz == 0 ? bbb.Min.Z : bbb.Max.Z);
                                        XYZ vp = vt.OfPoint(c);
                                        if (vp.X < vMinX) vMinX = vp.X;
                                        if (vp.X > vMaxX) vMaxX = vp.X;
                                        if (vp.Y < vMinY) vMinY = vp.Y;
                                        if (vp.Y > vMaxY) vMaxY = vp.Y;
                                    }

                                    double xExtent = vMaxX - vMinX;
                                    double yExtent = vMaxY - vMinY;

                                    // X幅 > Y高さ = 平行梁（対象）
                                    // X幅 <= Y高さ = 直交梁（除外）
                                    if (xExtent > yExtent)
                                        viewTargets.Add(beam);
                                }
                            }
                            else
                            {
                                viewTargets.AddRange(viewBeams);
                            }

                            // 要素をパラメータ値でグループ化
                            var elementsByType = new Dictionary<string, List<Element>>();
                            foreach (var elem in viewTargets)
                            {
                                Parameter p = elem.LookupParameter(paramName);
                                if (p == null) continue;
                                string value = p.StorageType == StorageType.String
                                    ? p.AsString() : p.AsValueString();
                                if (string.IsNullOrEmpty(value) || value.Trim().Length == 0)
                                    continue;
                                value = value.Trim();

                                if (!elementsByType.ContainsKey(value))
                                    elementsByType[value] = new List<Element>();
                                elementsByType[value].Add(elem);
                            }

                            // 断面ビュー: 柱のX範囲から梁クリップ境界を計算
                            double sClipMinX = double.NaN, sClipMaxX = double.NaN;
                            if (view.ViewType == ViewType.Section && viewColumns.Count > 0)
                            {
                                try
                                {
                                    Transform vInv = view.CropBox.Transform.Inverse;
                                    double cMinX = double.MaxValue, cMaxX = double.MinValue;
                                    double commonOff = settings.UseCommonOffset
                                        ? settings.CommonOffsetMm / 304.8
                                        : offsetByType.Values.FirstOrDefault();

                                    foreach (var col in viewColumns)
                                    {
                                        BoundingBoxXYZ cbb = col.get_BoundingBox(view)
                                            ?? col.get_BoundingBox(null);
                                        if (cbb == null) continue;

                                        for (int ix = 0; ix <= 1; ix++)
                                        for (int iy = 0; iy <= 1; iy++)
                                        for (int iz = 0; iz <= 1; iz++)
                                        {
                                            XYZ cc = new XYZ(
                                                ix == 0 ? cbb.Min.X : cbb.Max.X,
                                                iy == 0 ? cbb.Min.Y : cbb.Max.Y,
                                                iz == 0 ? cbb.Min.Z : cbb.Max.Z);
                                            XYZ cvp = vInv.OfPoint(cc);
                                            if (cvp.X < cMinX) cMinX = cvp.X;
                                            if (cvp.X > cMaxX) cMaxX = cvp.X;
                                        }
                                    }

                                    if (cMinX < double.MaxValue)
                                    {
                                        sClipMinX = cMinX - commonOff;
                                        sClipMaxX = cMaxX + commonOff;
                                    }
                                }
                                catch { }
                            }

                            // 梁塗潰領域
                            regionCount += FilledRegionCreator.CreateFilledRegions(
                                doc, view, elementsByType, offsetByType,
                                settings.FillPatternId, settings.LineStyleId,
                                settings.Types, settings.OverwriteExisting,
                                sClipMinX, sClipMaxX);

                            // 断面ビュー: 柱を梁オフセット端でクリップして配置
                            if (view.ViewType == ViewType.Section &&
                                viewColumns.Count > 0 && paramName != null)
                            {
                                // 平行梁のBBoxリストのみ収集（柱クリップ用）
                                // viewTargetsには直交梁が除外済み
                                var beamBBoxes = new List<BoundingBoxXYZ>();
                                foreach (var b in viewTargets)
                                {
                                    if (b.Category.Id.IntegerValue !=
                                        (int)BuiltInCategory.OST_StructuralFraming) continue;
                                    var bbb = b.get_BoundingBox(view)
                                              ?? b.get_BoundingBox(null);
                                    if (bbb != null) beamBBoxes.Add(bbb);
                                }

                                // 柱をパラメータ値でグループ化
                                var colsByType = new Dictionary<string, List<Element>>();
                                foreach (var col in viewColumns)
                                {
                                    Parameter p = col.LookupParameter(paramName);
                                    if (p == null) continue;
                                    string value = p.StorageType == StorageType.String
                                        ? p.AsString() : p.AsValueString();
                                    if (string.IsNullOrEmpty(value)) continue;
                                    value = value.Trim();
                                    if (!colsByType.ContainsKey(value))
                                        colsByType[value] = new List<Element>();
                                    colsByType[value].Add(col);
                                }

                                // 共通オフセット取得
                                double commonOffset = settings.UseCommonOffset
                                    ? settings.CommonOffsetMm / 304.8 : 0;

                                regionCount += FilledRegionCreator.CreateSectionColumnRegions(
                                    doc, view, colsByType, beamBBoxes,
                                    commonOffset > 0 ? commonOffset : offsetByType.Values.FirstOrDefault(),
                                    settings.FillPatternId, settings.LineStyleId,
                                    settings.Types, settings.OverwriteExisting);
                            }

                            // 柱枠型塗潰領域（平面/天伏/構造伏のみ）
                            // 天伏ではviewColumnsが空の場合があるため、ドキュメント全体の柱も参照
                            var colsForFrame = viewColumns.Count > 0 ? viewColumns
                                : new FilteredElementCollector(doc)
                                    .OfCategory(BuiltInCategory.OST_StructuralColumns)
                                    .WhereElementIsNotElementType()
                                    .Cast<Element>().ToList();
                            if (view.ViewType != ViewType.Section &&
                                colsForFrame.Count > 0 && paramName != null)
                            {
                                double aFeet = settings.ColumnA_mm / 304.8;
                                double bFeet = settings.ColumnB_mm / 304.8;
                                regionCount += FilledRegionCreator.CreateColumnFrameRegions(
                                    doc, view, colsForFrame, paramName,
                                    aFeet, bFeet,
                                    settings.FillPatternId, settings.LineStyleId,
                                    settings.Types, settings.OverwriteExisting);
                            }
                        }

                        // 凡例は1回だけ作成
                        bool hasColumnFrame = targetViews.Any(
                            v => v.ViewType != ViewType.Section);
                        var beamCountByType = new Dictionary<string, int>();
                        legendViewId = LegendManager.CreateLegendDraftingView(
                            doc, settings.Types, beamCountByType,
                            settings.OverwriteExisting, settings.TextNoteTypeId,
                            hasColumnFrame,
                            settings.ColumnA_mm / 304.8,
                            settings.ColumnB_mm / 304.8);

                        trans.Commit();
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        throw new Exception(
                            "トランザクション内でエラー: " + ex.Message +
                            "\n" + ex.StackTrace, ex);
                    }
                }

                // 凡例のシート自動配置（別トランザクション: ビューをコミット後に配置）
                if (isSheet && legendViewId != null)
                {
                    using (Transaction vpTrans = new Transaction(doc, "凡例シート配置"))
                    {
                        vpTrans.Start();
                        try
                        {
                            var sheet = activeView as ViewSheet;

                            bool canAdd = Viewport.CanAddViewToSheet(
                                doc, sheet.Id, legendViewId);

                            if (canAdd)
                            {
                                BoundingBoxUV outline = sheet.Outline;
                                double margin = 50.0 / 304.8;
                                XYZ position = new XYZ(
                                    outline.Max.U - margin,
                                    outline.Min.V + margin + 30.0 / 304.8, 0);

                                Viewport vp = Viewport.Create(
                                    doc, sheet.Id, legendViewId, position);

                                try
                                {
                                    System.IO.File.AppendAllText(
                                        @"C:\temp\FireProtection_debug.txt",
                                        $"\n凡例配置: VP={vp?.Id?.IntegerValue} pos=({position.X * 304.8:F0},{position.Y * 304.8:F0})\n");
                                }
                                catch { }
                            }

                            vpTrans.Commit();
                        }
                        catch (Exception vpEx)
                        {
                            vpTrans.RollBack();
                            try
                            {
                                System.IO.File.AppendAllText(
                                    @"C:\temp\FireProtection_debug.txt",
                                    $"\n凡例配置エラー: {vpEx.Message}\n");
                            }
                            catch { }
                        }
                    }
                }

                string legendInfo = legendViewId != null
                    ? "凡例ビュー「耐火被覆色分け凡例」を作成しました"
                    : "凡例ビュー作成をスキップしました";

                TaskDialog.Show("耐火被覆色分け - 完了",
                    $"処理が完了しました。\n\n" +
                    $"対象ビュー数: {targetViews.Count}\n" +
                    $"塗潰領域: {regionCount} 個作成\n" +
                    $"耐火被覆種類: {settings.Types.Count}\n" +
                    $"{legendInfo}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"耐火被覆色分け処理中にエラーが発生しました。\n\n{ex.Message}" +
                    $"\n\n--- スタックトレース ---\n{ex.StackTrace}" +
                    (ex.InnerException != null
                        ? $"\n\n--- InnerException ---\n{ex.InnerException.Message}\n{ex.InnerException.StackTrace}"
                        : "") +
                    "\n\nマニュアル: https://28yu.github.io/28tools-manual/" +
                    "\n配布サイト: https://28yu.github.io/28tools-download/";
                return Result.Failed;
            }
        }

        /// <summary>
        /// 塗潰領域の境界で選択できる線種を収集
        /// </summary>
        private static List<LineStyleItem> CollectLineStyles(Document doc)
        {
            var lineStyles = new List<LineStyleItem>();
            Category linesCat = doc.Settings.Categories
                .get_Item(BuiltInCategory.OST_Lines);
            if (linesCat == null) return lineStyles;

            var excludeKeywords = new[]
            {
                "スケッチ", "Sketch",
                "部屋を分割", "Room Separation",
                "スペースの分割", "Space Separator",
                "メッシュ筋", "Fabric", "Area Reinforcement",
                "回転の軸", "Axis of Rotation",
                "断熱層", "Insulation Batting",
            };

            foreach (Category subCat in linesCat.SubCategories)
            {
                string name = subCat.Name;
                bool excluded = false;
                foreach (var kw in excludeKeywords)
                {
                    if (name.Contains(kw)) { excluded = true; break; }
                }
                if (excluded) continue;

                GraphicsStyle gs = subCat.GetGraphicsStyle(GraphicsStyleType.Projection);
                if (gs != null)
                    lineStyles.Add(new LineStyleItem { Id = gs.Id, Name = name });
            }
            lineStyles = lineStyles.OrderBy(ls => ls.Name).ToList();

            // <非表示>
            if (!lineStyles.Any(ls => ls.Name.Contains("非表示") || ls.Name.Contains("Invisible")))
            {
                try
                {
                    Category invisCat = doc.Settings.Categories
                        .get_Item(BuiltInCategory.OST_InvisibleLines);
                    if (invisCat != null)
                    {
                        GraphicsStyle invisGs = invisCat.GetGraphicsStyle(
                            GraphicsStyleType.Projection);
                        if (invisGs != null)
                            lineStyles.Add(new LineStyleItem { Id = invisGs.Id, Name = invisCat.Name });
                    }
                }
                catch { }

                if (!lineStyles.Any(ls => ls.Name.Contains("非表示") || ls.Name.Contains("Invisible")))
                {
                    var invisStyle = new FilteredElementCollector(doc)
                        .OfClass(typeof(GraphicsStyle))
                        .Cast<GraphicsStyle>()
                        .FirstOrDefault(gs =>
                            gs.GraphicsStyleType == GraphicsStyleType.Projection &&
                            (gs.Name.Contains("非表示") || gs.Name.Contains("Invisible")));
                    if (invisStyle != null)
                        lineStyles.Add(new LineStyleItem { Id = invisStyle.Id, Name = invisStyle.Name });
                }
            }

            return lineStyles;
        }
    }
}

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

                // 対応ビュータイプ: 平面、天伏、構造伏、断面
                if (activeView.ViewType != ViewType.FloorPlan &&
                    activeView.ViewType != ViewType.CeilingPlan &&
                    activeView.ViewType != ViewType.EngineeringPlan &&
                    activeView.ViewType != ViewType.Section)
                {
                    TaskDialog.Show("エラー",
                        "平面ビュー、天井伏図、構造伏図、または断面図で実行してください。");
                    return Result.Cancelled;
                }

                // ビューテンプレート確認
                if (activeView.ViewTemplateId != ElementId.InvalidElementId)
                {
                    TaskDialogResult templateResult = TaskDialog.Show(
                        "ビューテンプレート確認",
                        "ビューテンプレートが設定されています。\n" +
                        "塗潰領域を配置するにはテンプレートを解除する必要があります。\n\n" +
                        "テンプレートを解除しますか？",
                        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                    if (templateResult == TaskDialogResult.Yes)
                    {
                        using (Transaction t = new Transaction(doc, "ビューテンプレート解除"))
                        {
                            t.Start();
                            activeView.ViewTemplateId = ElementId.InvalidElementId;
                            t.Commit();
                        }
                    }
                    else
                    {
                        return Result.Cancelled;
                    }
                }

                // 梁を取得
                var beams = new FilteredElementCollector(doc, activeView.Id)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();

                // 柱を取得
                var columns = new FilteredElementCollector(doc, activeView.Id)
                    .OfCategory(BuiltInCategory.OST_StructuralColumns)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();

                if (beams.Count == 0 && columns.Count == 0)
                {
                    TaskDialog.Show("エラー",
                        "ビュー内に梁（構造フレーム）または柱（構造柱）が見つかりません。");
                    return Result.Cancelled;
                }

                // パラメータ自動検出
                var beamParams = BeamGeometryHelper.DetectFireProtectionParameters(beams);
                var columnParams = BeamGeometryHelper.DetectFireProtectionParameters(columns);

                // 線種取得（塗潰領域の境界で選択できる線種と同じ）
                // Linesカテゴリのサブカテゴリからツール専用の線種を除外
                var lineStyles = new List<LineStyleItem>();
                Category linesCat = doc.Settings.Categories
                    .get_Item(BuiltInCategory.OST_Lines);
                if (linesCat != null)
                {
                    // ツール専用線種の除外キーワード（日英両対応）
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

                        GraphicsStyle gs = subCat.GetGraphicsStyle(
                            GraphicsStyleType.Projection);
                        if (gs != null)
                        {
                            lineStyles.Add(new LineStyleItem
                            {
                                Id = gs.Id,
                                Name = name
                            });
                        }
                    }
                    lineStyles = lineStyles.OrderBy(ls => ls.Name).ToList();

                    // <非表示>（Invisible Lines）を追加
                    if (!lineStyles.Any(ls => ls.Name.Contains("非表示") || ls.Name.Contains("Invisible")))
                    {
                        // 方法1: BuiltInCategory
                        try
                        {
                            Category invisCat = doc.Settings.Categories
                                .get_Item(BuiltInCategory.OST_InvisibleLines);
                            if (invisCat != null)
                            {
                                GraphicsStyle invisGs = invisCat.GetGraphicsStyle(
                                    GraphicsStyleType.Projection);
                                if (invisGs != null)
                                {
                                    lineStyles.Add(new LineStyleItem
                                    {
                                        Id = invisGs.Id,
                                        Name = invisCat.Name
                                    });
                                }
                            }
                        }
                        catch { }

                        // 方法2: 全GraphicsStyleから検索
                        if (!lineStyles.Any(ls => ls.Name.Contains("非表示") || ls.Name.Contains("Invisible")))
                        {
                            var invisStyle = new FilteredElementCollector(doc)
                                .OfClass(typeof(GraphicsStyle))
                                .Cast<GraphicsStyle>()
                                .FirstOrDefault(gs =>
                                    gs.GraphicsStyleType == GraphicsStyleType.Projection &&
                                    (gs.Name.Contains("非表示") || gs.Name.Contains("Invisible")));
                            if (invisStyle != null)
                            {
                                lineStyles.Add(new LineStyleItem
                                {
                                    Id = invisStyle.Id,
                                    Name = invisStyle.Name
                                });
                            }
                        }
                    }
                }

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
                switch (activeView.ViewType)
                {
                    case ViewType.FloorPlan: viewTypeName = "平面ビュー"; break;
                    case ViewType.CeilingPlan: viewTypeName = "天井伏図"; break;
                    case ViewType.EngineeringPlan: viewTypeName = "構造伏図"; break;
                    case ViewType.Section: viewTypeName = "断面図"; break;
                    default: viewTypeName = activeView.ViewType.ToString(); break;
                }

                var dialogData = new FireProtectionDialogData
                {
                    ViewName = activeView.Name,
                    ViewTypeName = viewTypeName,
                    IsSectionView = activeView.ViewType == ViewType.Section,
                    BeamCount = beams.Count,
                    ColumnCount = columns.Count,
                    HasBeams = beams.Count > 0,
                    HasColumns = columns.Count > 0,
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

                // 対象要素を収集
                var targetElements = new List<Element>();
                if (settings.IncludeBeams) targetElements.AddRange(beams);
                if (settings.IncludeColumns) targetElements.AddRange(columns);

                // パラメータ名を取得（表示用テキストを除去）
                string paramName = settings.SelectedParameterName;
                if (paramName != null && paramName.Contains("（"))
                    paramName = paramName.Substring(0, paramName.IndexOf("（"));

                // 要素をパラメータ値でグループ化
                var elementsByType = new Dictionary<string, List<Element>>();
                var beamCountByType = new Dictionary<string, int>();

                foreach (var elem in targetElements)
                {
                    Parameter p = elem.LookupParameter(paramName);
                    if (p == null) continue;

                    string value = null;
                    if (p.StorageType == StorageType.String)
                        value = p.AsString();
                    else
                        value = p.AsValueString();

                    if (string.IsNullOrEmpty(value) || value.Trim().Length == 0)
                        continue;

                    value = value.Trim();

                    if (!elementsByType.ContainsKey(value))
                    {
                        elementsByType[value] = new List<Element>();
                        beamCountByType[value] = 0;
                    }
                    elementsByType[value].Add(elem);
                    beamCountByType[value]++;
                }

                // オフセット辞書
                var offsetByType = new Dictionary<string, double>();
                foreach (var typeEntry in settings.Types)
                {
                    double offsetMm = settings.UseCommonOffset
                        ? settings.CommonOffsetMm
                        : typeEntry.OffsetMm;
                    offsetByType[typeEntry.Name] = offsetMm / 304.8;
                }

                // 実行
                ElementId legendViewId = null;
                int regionCount = 0;

                using (Transaction trans = new Transaction(doc, "耐火被覆色分け"))
                {
                    trans.Start();

                    try
                    {
                        regionCount = FilledRegionCreator.CreateFilledRegions(
                            doc, activeView, elementsByType, offsetByType,
                            settings.FillPatternId, settings.LineStyleId,
                            settings.Types, settings.OverwriteExisting);

                        legendViewId = LegendManager.CreateLegendDraftingView(
                            doc, settings.Types, beamCountByType,
                            settings.OverwriteExisting, settings.TextNoteTypeId);

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

                string legendInfo = legendViewId != null
                    ? "凡例ビュー「耐火被覆色分け凡例」を作成しました"
                    : "凡例ビュー作成をスキップしました";

                int totalElements = elementsByType.Values.Sum(b => b.Count);

                TaskDialog.Show("耐火被覆色分け - 完了",
                    $"処理が完了しました。\n\n" +
                    $"対象要素数: {totalElements}\n" +
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
    }
}

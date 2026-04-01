using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.FireProtection
{
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

                if (activeView.ViewType != ViewType.FloorPlan &&
                    activeView.ViewType != ViewType.CeilingPlan &&
                    activeView.ViewType != ViewType.EngineeringPlan)
                {
                    TaskDialog.Show("エラー",
                        "平面ビュー、天井伏図、または構造伏図で実行してください。");
                    return Result.Cancelled;
                }

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

                var beams = new FilteredElementCollector(doc, activeView.Id)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .ToList();

                if (beams.Count == 0)
                {
                    TaskDialog.Show("エラー", "ビュー内に梁（構造フレーム）が見つかりません。");
                    return Result.Cancelled;
                }

                var beamTypes = beams
                    .GroupBy(b => $"{b.Symbol.Family.Name}: {b.Symbol.Name}")
                    .Select(g => new BeamTypeInfo
                    {
                        FamilyName = g.First().Symbol.Family.Name,
                        TypeName = g.First().Symbol.Name,
                        Count = g.Count(),
                        Beams = g.ToList()
                    })
                    .OrderBy(bt => bt.DisplayName)
                    .ToList();

                var lineStyles = new FilteredElementCollector(doc)
                    .OfClass(typeof(GraphicsStyle))
                    .Cast<GraphicsStyle>()
                    .Where(gs => gs.GraphicsStyleType == GraphicsStyleType.Projection)
                    .OrderBy(gs => gs.Name)
                    .Select(gs => new LineStyleItem { Id = gs.Id, Name = gs.Name })
                    .ToList();

                var fillPatterns = new FilteredElementCollector(doc)
                    .OfClass(typeof(FillPatternElement))
                    .Cast<FillPatternElement>()
                    .OrderBy(fp => fp.Name)
                    .Select(fp => new FillPatternItem { Id = fp.Id, Name = fp.Name })
                    .ToList();

                var textNoteTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .OrderBy(t => t.Name)
                    .Select(t => new FpTextNoteTypeItem(t))
                    .ToList();

                Level refLevel = activeView.GenLevel;

                var dialogData = new FireProtectionDialogData
                {
                    ViewName = activeView.Name,
                    BeamCount = beams.Count,
                    RefLevel = refLevel,
                    BeamTypes = beamTypes,
                    TextNoteTypes = textNoteTypes,
                    LineStyles = lineStyles,
                    FillPatterns = fillPatterns
                };

                var dialog = new FireProtectionDialog(dialogData);
                if (dialog.ShowDialog() != true)
                    return Result.Cancelled;

                var settings = dialog.GetResult();

                var beamsByFpType = new Dictionary<string, List<FamilyInstance>>();
                var beamCountByType = new Dictionary<string, int>();

                foreach (var bt in beamTypes)
                {
                    string assignment;
                    if (!settings.BeamTypeAssignments.TryGetValue(
                            bt.DisplayName, out assignment))
                        continue;
                    if (assignment == "除外") continue;

                    if (!beamsByFpType.ContainsKey(assignment))
                        beamsByFpType[assignment] = new List<FamilyInstance>();
                    beamsByFpType[assignment].AddRange(bt.Beams);

                    if (!beamCountByType.ContainsKey(assignment))
                        beamCountByType[assignment] = 0;
                    beamCountByType[assignment] += bt.Count;
                }

                var offsetByType = new Dictionary<string, double>();
                foreach (var fpType in settings.Types)
                {
                    double offsetMm = settings.UseCommonOffset
                        ? settings.CommonOffsetMm
                        : fpType.OffsetMm;
                    offsetByType[fpType.Name] = offsetMm / 304.8;
                }

                ElementId legendViewId = null;
                int regionCount = 0;

                using (Transaction trans = new Transaction(doc, "耐火被覆色分け"))
                {
                    trans.Start();

                    try
                    {
                        regionCount = FilledRegionCreator.CreateFilledRegions(
                            doc, activeView, beamsByFpType, offsetByType,
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
                            "トランザクション内でエラーが発生しました: " + ex.Message, ex);
                    }
                }

                string legendInfo = legendViewId != null
                    ? "凡例ビュー「耐火被覆色分け凡例」を作成しました"
                    : "凡例ビュー作成をスキップしました";

                int totalBeams = beamsByFpType.Values.Sum(b => b.Count);

                TaskDialog.Show("耐火被覆色分け - 完了",
                    $"処理が完了しました。\n\n" +
                    $"対象梁数: {totalBeams}\n" +
                    $"塗潰領域: {regionCount} 個作成\n" +
                    $"耐火被覆種類: {settings.Types.Count}\n" +
                    $"{legendInfo}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"耐火被覆色分け処理中にエラーが発生しました。\n\n{ex.Message}" +
                    "\n\nマニュアル: https://28yu.github.io/28tools-manual/" +
                    "\n配布サイト: https://28yu.github.io/28tools-download/";
                return Result.Failed;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.WallArea
{
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class WallAreaCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                View activeView = doc.ActiveView;

                // アクティブビューに表示されている壁を取得
                var walls = new FilteredElementCollector(doc, activeView.Id)
                    .OfClass(typeof(Wall))
                    .Cast<Wall>()
                    .ToList();

                if (walls.Count == 0)
                {
                    TaskDialog.Show("壁面積集計", "このビューに壁がありません。");
                    return Result.Succeeded;
                }

                // 壁タイプ別にグループ化して面積を集計
                var summary = walls
                    .GroupBy(w => w.WallType.Name)
                    .Select(g => new
                    {
                        TypeName = g.Key,
                        Count = g.Count(),
                        TotalArea = g.Sum(w =>
                        {
                            var areaParam = w.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                            return areaParam != null ? areaParam.AsDouble() : 0.0;
                        })
                    })
                    .OrderByDescending(x => x.TotalArea)
                    .ToList();

                // 単位変換（内部単位 → ㎡）
                double grandTotalArea = 0;
                int grandTotalCount = 0;

                var sb = new StringBuilder();
                foreach (var item in summary)
                {
                    double areaSqm = UnitUtils.ConvertFromInternalUnits(item.TotalArea, UnitTypeId.SquareMeters);
                    grandTotalArea += areaSqm;
                    grandTotalCount += item.Count;
                    sb.AppendLine($"{item.TypeName}:  {areaSqm:F2} ㎡（{item.Count}本）");
                }

                sb.AppendLine("─────────────────");
                sb.AppendLine($"合計:  {grandTotalArea:F2} ㎡（{grandTotalCount}本）");

                // ダイアログ表示
                var dialog = new TaskDialog("壁面積集計");
                dialog.MainInstruction = $"壁面積集計（ビュー: {activeView.Name}）";
                dialog.MainContent = sb.ToString();
                dialog.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}

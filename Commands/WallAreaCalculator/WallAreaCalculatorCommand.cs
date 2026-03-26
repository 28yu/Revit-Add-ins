using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.WallAreaCalculator
{
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class WallAreaCalculatorCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                Document doc = uiDoc.Document;
                View activeView = doc.ActiveView;

                // アクティブビューに表示されている壁を取得
                FilteredElementCollector collector = new FilteredElementCollector(doc, activeView.Id);
                List<Wall> walls = collector
                    .OfClass(typeof(Wall))
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .ToList();

                if (walls.Count == 0)
                {
                    TaskDialog.Show("壁面積集計", "このビューに壁がありません。");
                    return Result.Succeeded;
                }

                // 壁タイプごとにグループ化して面積を集計
                var groupedWalls = walls
                    .GroupBy(w => w.WallType.Name)
                    .OrderByDescending(g => GetGroupArea(g))
                    .ToList();

                StringBuilder sb = new StringBuilder();
                double totalArea = 0;
                int totalCount = 0;

                foreach (var group in groupedWalls)
                {
                    double areaInternalUnits = GetGroupArea(group);
                    double areaSqm = UnitUtils.ConvertFromInternalUnits(areaInternalUnits, UnitTypeId.SquareMeters);
                    int count = group.Count();

                    sb.AppendLine(string.Format("{0}:  {1:F2} ㎡（{2}本）",
                        group.Key, areaSqm, count));

                    totalArea += areaInternalUnits;
                    totalCount += count;
                }

                double totalAreaSqm = UnitUtils.ConvertFromInternalUnits(totalArea, UnitTypeId.SquareMeters);

                sb.AppendLine("─────────────────");
                sb.AppendLine(string.Format("合計:  {0:F2} ㎡（{1}本）",
                    totalAreaSqm, totalCount));

                TaskDialog dialog = new TaskDialog("壁面積集計");
                dialog.MainInstruction = string.Format("壁面積集計（ビュー: {0}）", activeView.Name);
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

        /// <summary>
        /// グループ内の壁面積合計を内部単位で取得
        /// </summary>
        private static double GetGroupArea(IGrouping<string, Wall> group)
        {
            double total = 0;
            foreach (Wall wall in group)
            {
                Parameter areaParam = wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                if (areaParam != null)
                {
                    total += areaParam.AsDouble();
                }
            }
            return total;
        }
    }
}

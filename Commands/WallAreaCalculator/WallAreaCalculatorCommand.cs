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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
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

                // 壁タイプ別にグループ化して面積を集計
                var groupedWalls = walls
                    .GroupBy(w => w.WallType.Name)
                    .OrderByDescending(g => g.Sum(w => GetWallArea(w)))
                    .ToList();

                double totalArea = 0;
                int totalCount = 0;
                StringBuilder sb = new StringBuilder();

                foreach (var group in groupedWalls)
                {
                    string typeName = group.Key;
                    int count = group.Count();
                    double area = group.Sum(w => GetWallArea(w));

                    // 内部単位（平方フィート）から平方メートルに変換
                    double areaM2 = UnitUtils.ConvertFromInternalUnits(area, UnitTypeId.SquareMeters);

                    sb.AppendLine($"{typeName}:  {areaM2:F2} \u33a1\uff08{count}\u672c\uff09");

                    totalArea += area;
                    totalCount += count;
                }

                double totalAreaM2 = UnitUtils.ConvertFromInternalUnits(totalArea, UnitTypeId.SquareMeters);

                sb.AppendLine("\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500");
                sb.AppendLine($"\u5408\u8a08:  {totalAreaM2:F2} \u33a1\uff08{totalCount}\u672c\uff09");

                string title = $"\u58c1\u9762\u7a4d\u96c6\u8a08\uff08\u30d3\u30e5\u30fc: {activeView.Name}\uff09";
                TaskDialog.Show(title, sb.ToString());

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("\u30a8\u30e9\u30fc", $"\u58c1\u9762\u7a4d\u96c6\u8a08\u4e2d\u306b\u30a8\u30e9\u30fc\u304c\u767a\u751f\u3057\u307e\u3057\u305f\u3002\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        /// <summary>
        /// 壁の面積を取得する（内部単位: 平方フィート）
        /// </summary>
        private double GetWallArea(Wall wall)
        {
            Parameter areaParam = wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
            if (areaParam != null && areaParam.HasValue)
            {
                return areaParam.AsDouble();
            }
            return 0;
        }
    }
}

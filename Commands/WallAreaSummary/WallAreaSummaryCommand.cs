using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.WallAreaSummary
{
    [Transaction(TransactionMode.Manual)]
    public class WallAreaSummaryCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            // アクティブなビューに表示されている壁を全て取得
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
            var summary = walls
                .GroupBy(w => w.WallType.Name)
                .Select(g => new
                {
                    TypeName = g.Key,
                    Count = g.Count(),
                    TotalArea = g.Sum(w =>
                    {
                        Parameter areaParam = w.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                        if (areaParam != null && areaParam.HasValue)
                        {
#if REVIT2021 || REVIT2022
                            return UnitUtils.ConvertFromInternalUnits(areaParam.AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                            return UnitUtils.ConvertFromInternalUnits(areaParam.AsDouble(), UnitTypeId.SquareMeters);
#endif
                        }
                        return 0.0;
                    })
                })
                .OrderByDescending(x => x.TotalArea)
                .ToList();

            // 表示テキストを組み立て
            string result = "";
            int totalCount = 0;
            double totalArea = 0.0;

            foreach (var item in summary)
            {
                result += $"{item.TypeName}:  {item.TotalArea:F2} ㎡（{item.Count}本）\n";
                totalCount += item.Count;
                totalArea += item.TotalArea;
            }

            result += "─────────────────\n";
            result += $"合計:  {totalArea:F2} ㎡（{totalCount}本）";

            TaskDialog dialog = new TaskDialog("壁面積集計");
            dialog.MainInstruction = $"壁面積集計（ビュー: {activeView.Name}）";
            dialog.MainContent = result;
            dialog.Show();

            return Result.Succeeded;
        }
    }
}
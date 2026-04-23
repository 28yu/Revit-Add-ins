using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    internal static class ElementCollector
    {
        private static readonly Dictionary<string, BuiltInCategory> _nameToCat
            = new Dictionary<string, BuiltInCategory>
            {
                { "StructuralColumns", BuiltInCategory.OST_StructuralColumns },
                { "StructuralFraming", BuiltInCategory.OST_StructuralFraming },
                { "Walls", BuiltInCategory.OST_Walls },
                { "Floors", BuiltInCategory.OST_Floors },
                { "StructuralFoundation", BuiltInCategory.OST_StructuralFoundation },
                { "Stairs", BuiltInCategory.OST_Stairs },
            };

        internal static List<Element> Collect(Document doc, FormworkSettings settings, View activeView)
        {
            var result = new List<Element>();
            var seenIds = new HashSet<int>();

            foreach (var key in settings.IncludedCategories)
            {
                if (!_nameToCat.TryGetValue(key, out var bic))
                    continue;

                FilteredElementCollector col = settings.Scope == CalculationScope.CurrentView
                    ? new FilteredElementCollector(doc, activeView.Id)
                    : new FilteredElementCollector(doc);

                var elems = col.OfCategory(bic).WhereElementIsNotElementType().ToList();

                if (bic == BuiltInCategory.OST_Walls)
                {
                    // 壁本体のみ（カーテンウォール等は除外）。WallSweep は別途追加する。
                    elems = elems.Where(e =>
                    {
                        if (!(e is Wall w)) return false;
                        var wt = doc.GetElement(w.GetTypeId()) as WallType;
                        if (wt == null) return true;
                        return wt.Function == WallFunction.Retaining ||
                               wt.Kind == WallKind.Basic;
                    }).ToList();

                    foreach (var e in elems)
                    {
                        if (seenIds.Add(e.Id.IntegerValue)) result.Add(e);
                    }

                    // 壁スイープ・リビール（独立した WallSweep 要素として存在）も追加
                    FilteredElementCollector swCol = settings.Scope == CalculationScope.CurrentView
                        ? new FilteredElementCollector(doc, activeView.Id)
                        : new FilteredElementCollector(doc);
                    var sweeps = swCol.OfClass(typeof(WallSweep))
                        .WhereElementIsNotElementType()
                        .ToList();
                    foreach (var sw in sweeps)
                    {
                        if (seenIds.Add(sw.Id.IntegerValue)) result.Add(sw);
                    }
                }
                else
                {
                    foreach (var e in elems)
                    {
                        if (seenIds.Add(e.Id.IntegerValue)) result.Add(e);
                    }
                }
            }
            return result;
        }

        internal static CategoryGroup ToCategoryGroup(Element elem)
        {
            if (elem == null) return CategoryGroup.Other;
            if (elem is WallSweep) return CategoryGroup.Wall;
            if (elem.Category == null) return CategoryGroup.Other;
            switch ((BuiltInCategory)elem.Category.Id.IntegerValue)
            {
                case BuiltInCategory.OST_StructuralColumns: return CategoryGroup.Column;
                case BuiltInCategory.OST_StructuralFraming: return CategoryGroup.Beam;
                case BuiltInCategory.OST_Walls: return CategoryGroup.Wall;
                case BuiltInCategory.OST_Floors: return CategoryGroup.Slab;
                case BuiltInCategory.OST_StructuralFoundation: return CategoryGroup.Foundation;
                case BuiltInCategory.OST_Stairs: return CategoryGroup.Stairs;
                default: return CategoryGroup.Other;
            }
        }

        internal static string GetParameterString(Element elem, string paramName)
        {
            if (elem == null || string.IsNullOrEmpty(paramName)) return string.Empty;
            Parameter p = elem.LookupParameter(paramName);
            if (p == null) return string.Empty;
            switch (p.StorageType)
            {
                case StorageType.String:
                    return p.AsString() ?? string.Empty;
                case StorageType.Integer:
                    return p.AsInteger().ToString();
                case StorageType.Double:
                    return p.AsValueString() ?? p.AsDouble().ToString("F2");
                case StorageType.ElementId:
                    return p.AsValueString() ?? string.Empty;
                default:
                    return string.Empty;
            }
        }
    }
}

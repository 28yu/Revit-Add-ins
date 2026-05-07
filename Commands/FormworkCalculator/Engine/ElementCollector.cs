using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    internal static class ElementCollector
    {
        /// <summary>
        /// 鉄骨と判定されて除外された要素 1 件分の情報。
        /// </summary>
        internal class ExcludedSteelEntry
        {
            public Element Element;
            public SteelMemberDetector.DetectionLayer Layer;
            public string Reason = string.Empty;
        }

        /// <summary>
        /// 要素収集結果。型枠算出対象の Targets と、鉄骨として除外された ExcludedSteel の 2 リスト。
        /// </summary>
        internal class CollectionResult
        {
            public List<Element> Targets = new List<Element>();
            public List<ExcludedSteelEntry> ExcludedSteel = new List<ExcludedSteelEntry>();
        }

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

        internal static CollectionResult CollectAndClassify(
            Document doc, FormworkSettings settings, View activeView)
        {
            var raw = Collect(doc, settings, activeView);
            var cr = new CollectionResult();

            bool exclude = settings?.ExcludeSteelMembers ?? true;

            FormworkDebugLog.Section("Steel Member Detection");
            int steelCount = 0;
            foreach (var elem in raw)
            {
                bool isCheckTarget = exclude && IsSteelDetectionTarget(elem);
                if (isCheckTarget)
                {
                    var det = SteelMemberDetector.Detect(elem, doc);
                    if (det != null && det.IsSteel)
                    {
                        cr.ExcludedSteel.Add(new ExcludedSteelEntry
                        {
                            Element = elem,
                            Layer = det.Layer,
                            Reason = det.Reason ?? string.Empty,
                        });
                        FormworkDebugLog.Log(
                            $"  [SteelExclude] E{elem.Id.IntegerValue} " +
                            $"Cat={elem.Category?.Name} Name='{elem.Name}' " +
                            $"L={det.Layer} reason={det.Reason}");
                        steelCount++;
                        continue;
                    }
                    else if (det != null && FormworkDebugLog.Enabled)
                    {
                        FormworkDebugLog.Log(
                            $"  [SteelKeep]    E{elem.Id.IntegerValue} " +
                            $"Cat={elem.Category?.Name} Name='{elem.Name}' " +
                            $"reason={det.Reason}");
                    }
                }
                cr.Targets.Add(elem);
            }
            FormworkDebugLog.Log(
                $"  Steel detection: target={raw.Count} excluded={steelCount} kept={cr.Targets.Count}");
            FormworkDebugLog.Flush();

            return cr;
        }

        /// <summary>
        /// 鉄骨検出の対象カテゴリ判定。構造柱・構造フレームのみ対象とする。
        /// </summary>
        private static bool IsSteelDetectionTarget(Element elem)
        {
            if (elem?.Category == null) return false;
            int catId = elem.Category.Id.IntegerValue;
            return catId == (int)BuiltInCategory.OST_StructuralColumns
                || catId == (int)BuiltInCategory.OST_StructuralFraming;
        }

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

        /// <summary>
        /// 要素の関連レベル名を取得する（スラブ・壁・柱・梁など各カテゴリを個別に対応）。
        /// 見つからない場合は空文字を返す。
        /// </summary>
        internal static string GetElementLevelName(Element elem)
        {
            if (elem == null) return string.Empty;
            var doc = elem.Document;

            ElementId levelId = null;
            try
            {
                // 要素型固有のプロパティ
                if (elem is Wall w) levelId = w.LevelId;
                else if (elem is Floor floor) levelId = floor.LevelId;
                else if (elem.LevelId != null && elem.LevelId != ElementId.InvalidElementId)
                    levelId = elem.LevelId;
            }
            catch { }

            if (levelId == null || levelId == ElementId.InvalidElementId)
            {
                // 代表的なパラメータから取得
                var candidates = new[]
                {
                    BuiltInParameter.LEVEL_PARAM,
                    BuiltInParameter.FAMILY_LEVEL_PARAM,
                    BuiltInParameter.SCHEDULE_LEVEL_PARAM,
                    BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                    BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
                };
                foreach (var bip in candidates)
                {
                    try
                    {
                        var p = elem.get_Parameter(bip);
                        if (p != null && p.HasValue && p.StorageType == StorageType.ElementId)
                        {
                            var id = p.AsElementId();
                            if (id != null && id != ElementId.InvalidElementId)
                            {
                                levelId = id;
                                break;
                            }
                        }
                    }
                    catch { }
                }
            }

            if (levelId == null || levelId == ElementId.InvalidElementId) return string.Empty;
            var level = doc.GetElement(levelId) as Level;
            return level?.Name ?? string.Empty;
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

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 型枠数量算出のメインクラス。
    /// 要素収集 → Solid 結合 → 面分類 → 開口加算 → 集計。
    /// </summary>
    internal class FormworkCalcEngine
    {
        private readonly Document _doc;
        private readonly FormworkSettings _settings;
        private readonly View _activeView;
        private readonly IProgress<string> _progress;

        internal FormworkCalcEngine(Document doc, FormworkSettings settings, View activeView, IProgress<string> progress = null)
        {
            _doc = doc;
            _settings = settings;
            _activeView = activeView;
            _progress = progress;
        }

        internal FormworkResult Run()
        {
            var result = new FormworkResult();

            _progress?.Report("要素を収集中...");
            var elements = ElementCollector.Collect(_doc, _settings, _activeView);
            result.ProcessedElementCount = elements.Count;

            if (elements.Count == 0)
            {
                return result;
            }

            // GL 高さ（m → feet）
            double? glFeet = null;
            if (_settings.UseGLDeduction)
            {
                glFeet = UnitUtils.ConvertToInternalUnits(_settings.GLElevationMeters, UnitTypeId.Meters);
            }

            _progress?.Report($"要素ごとに処理中... 0 / {elements.Count}");

            int idx = 0;
            var openingDeltas = OpeningProcessor.Compute(_doc, elements);
            var openingMap = openingDeltas.ToDictionary(o => o.HostElementId, o => o);

            // 個別要素ごとに計算（Phase 1実装: Boolean 結合は部位内 + 近接に限定）
            var elementsByGroup = elements.GroupBy(e => ElementCollector.ToCategoryGroup(e)).ToList();

            foreach (var grp in elementsByGroup)
            {
                foreach (var elem in grp)
                {
                    idx++;
                    if (idx % 10 == 0)
                        _progress?.Report($"要素ごとに処理中... {idx} / {elements.Count}");

                    try
                    {
                        var er = ProcessElement(elem, glFeet, openingMap);
                        if (er != null)
                        {
                            result.ElementResults.Add(er);
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Errors.Add(new ErrorLogEntry
                        {
                            ElementId = elem.Id.IntegerValue,
                            CategoryName = elem.Category?.Name ?? string.Empty,
                            ElementName = elem.Name,
                            ErrorKind = "ProcessError",
                            Message = ex.Message,
                        });
                    }
                }
            }

            Aggregate(result);
            return result;
        }

        private ElementResult ProcessElement(
            Element elem,
            double? glFeet,
            Dictionary<int, OpeningProcessor.OpeningDelta> openingMap)
        {
            var solids = SolidUnionProcessor.GetSolids(elem);
            if (solids.Count == 0) return null;

            // 同一要素内のソリッドを結合
            Solid unioned = SolidUnionProcessor.Union(solids);
            var finalSolids = unioned != null ? new List<Solid> { unioned } : solids;

            double minZ = FaceClassifier.GetMinZ(finalSolids);

            var faceInfos = FaceClassifier.ClassifyAll(finalSolids, glFeet, minZ);

            // 基礎底面以外の最下面を控除 (OST_StructuralFoundation のみ)
            bool isFoundation = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Foundation;
            if (!isFoundation)
            {
                // 非基礎要素: 最下の下向き面は DeductedContact として扱う（土中・スラブ上等）
                foreach (var fi in faceInfos)
                {
                    if (fi.FaceType == FaceType.DeductedBottom)
                    {
                        fi.FaceType = FaceType.DeductedContact;
                    }
                }
            }

            double formwork = 0;
            double dedTop = 0;
            double dedBottom = 0;
            double dedContact = 0;
            double inclined = 0;

            foreach (var fi in faceInfos)
            {
                double aM2 = FeetSqToM2(fi.Area);
                switch (fi.FaceType)
                {
                    case FaceType.FormworkRequired:
                        formwork += aM2;
                        break;
                    case FaceType.DeductedTop:
                        dedTop += aM2;
                        break;
                    case FaceType.DeductedBottom:
                        dedBottom += aM2;
                        break;
                    case FaceType.DeductedContact:
                    case FaceType.DeductedBelowGL:
                        dedContact += aM2;
                        break;
                    case FaceType.Inclined:
                        inclined += aM2;
                        break;
                }
            }

            double openingDeducted = 0;
            double openingAdded = 0;
            if (openingMap.TryGetValue(elem.Id.IntegerValue, out var od))
            {
                openingDeducted = FeetSqToM2(od.DeductedArea);
                openingAdded = FeetSqToM2(od.AddedEdgeArea);
                formwork = Math.Max(0, formwork - openingDeducted + openingAdded);
            }

            var er = new ElementResult
            {
                ElementId = elem.Id.IntegerValue,
                ElementName = elem.Name ?? string.Empty,
                Category = ElementCollector.ToCategoryGroup(elem),
                CategoryName = elem.Category?.Name ?? string.Empty,
                Zone = _settings.GroupByZone
                    ? NormalizeParamValue(ElementCollector.GetParameterString(elem, _settings.ZoneParameterName))
                    : string.Empty,
                FormworkType = _settings.GroupByFormworkType
                    ? NormalizeParamValue(ElementCollector.GetParameterString(elem, _settings.FormworkTypeParameterName))
                    : string.Empty,
                FormworkArea = formwork,
                DeductedTopArea = dedTop,
                DeductedBottomArea = dedBottom,
                DeductedContactArea = dedContact,
                InclinedArea = inclined,
                OpeningAreaDeducted = openingDeducted,
                OpeningEdgeAreaAdded = openingAdded,
            };

            return er;
        }

        private static string NormalizeParamValue(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "未設定";
            return s.Trim();
        }

        private static double FeetSqToM2(double feetSq)
        {
            return UnitUtils.ConvertFromInternalUnits(feetSq, UnitTypeId.SquareMeters);
        }

        private static void Aggregate(FormworkResult result)
        {
            double total = 0, deducted = 0, inclined = 0;

            var byCat = new Dictionary<CategoryGroup, CategoryResult>();
            var byZone = new Dictionary<string, ZoneResult>();
            var byType = new Dictionary<string, FormworkTypeResult>();

            foreach (var er in result.ElementResults)
            {
                total += er.FormworkArea;
                deducted += er.DeductedTopArea + er.DeductedBottomArea + er.DeductedContactArea;
                inclined += er.InclinedArea;

                if (!byCat.TryGetValue(er.Category, out var c))
                {
                    c = new CategoryResult { Category = er.Category, CategoryName = er.CategoryName };
                    byCat[er.Category] = c;
                }
                c.FormworkArea += er.FormworkArea;
                c.DeductedArea += er.DeductedTopArea + er.DeductedBottomArea + er.DeductedContactArea;
                c.ElementCount++;

                if (!string.IsNullOrEmpty(er.Zone))
                {
                    if (!byZone.TryGetValue(er.Zone, out var z))
                    {
                        z = new ZoneResult { Zone = er.Zone };
                        byZone[er.Zone] = z;
                    }
                    z.FormworkArea += er.FormworkArea;
                    z.ElementCount++;
                }

                if (!string.IsNullOrEmpty(er.FormworkType))
                {
                    if (!byType.TryGetValue(er.FormworkType, out var t))
                    {
                        t = new FormworkTypeResult { FormworkType = er.FormworkType };
                        byType[er.FormworkType] = t;
                    }
                    t.FormworkArea += er.FormworkArea;
                    t.ElementCount++;
                }
            }

            result.TotalFormworkArea = total;
            result.TotalDeductedArea = deducted;
            result.InclinedFaceArea = inclined;
            result.CategoryResults = byCat.Values.OrderBy(c => c.Category).ToList();
            result.ZoneResults = byZone.Values.OrderBy(z => z.Zone).ToList();
            result.TypeResults = byType.Values.OrderBy(t => t.FormworkType).ToList();
        }
    }
}

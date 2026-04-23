using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 型枠数量算出のメインクラス。
    /// Pass 1: 要素毎の Solid 取得・面分類
    /// Pass 2: ReferenceIntersector で要素間接触面の検出 → DeductedContact
    /// Pass 3: 開口加算・集計
    ///
    /// Pass 2 には View3D が必須なので、コマンド側から渡してもらう。
    /// </summary>
    internal class FormworkCalcEngine
    {
        private readonly Document _doc;
        private readonly FormworkSettings _settings;
        private readonly View _activeView;
        private readonly View3D _rayView;
        private readonly IProgress<string> _progress;

        internal FormworkCalcEngine(
            Document doc,
            FormworkSettings settings,
            View activeView,
            View3D rayView,
            IProgress<string> progress = null)
        {
            _doc = doc;
            _settings = settings;
            _activeView = activeView;
            _rayView = rayView;
            _progress = progress;
        }

        internal FormworkResult Run()
        {
            var result = new FormworkResult();

            _progress?.Report("要素を収集中...");
            var elements = ElementCollector.Collect(_doc, _settings, _activeView);
            result.ProcessedElementCount = elements.Count;
            if (elements.Count == 0) return result;

            double? glFeet = null;
            if (_settings.UseGLDeduction)
                glFeet = UnitUtils.ConvertToInternalUnits(_settings.GLElevationMeters, UnitTypeId.Meters);

            var openingDeltas = OpeningProcessor.Compute(_doc, elements);
            var openingMap = openingDeltas.ToDictionary(o => o.HostElementId, o => o);

            // Pass 1: 要素毎に面を分類
            _progress?.Report($"Pass 1: 面分類中... 0 / {elements.Count}");
            var contexts = new List<ContactFaceDetector.ElementFacesContext>();
            var elemByContext = new Dictionary<int, Element>();

            int idx = 0;
            foreach (var elem in elements)
            {
                idx++;
                if (idx % 10 == 0)
                    _progress?.Report($"Pass 1: 面分類中... {idx} / {elements.Count}");

                try
                {
                    var ctx = ClassifyElementFaces(elem, glFeet);
                    if (ctx != null)
                    {
                        contexts.Add(ctx);
                        elemByContext[ctx.ElementId] = elem;
                    }
                }
                catch (Exception ex)
                {
                    result.Errors.Add(new ErrorLogEntry
                    {
                        ElementId = elem.Id.IntegerValue,
                        CategoryName = elem.Category?.Name ?? string.Empty,
                        ElementName = elem.Name,
                        ErrorKind = "ClassifyError",
                        Message = ex.Message,
                    });
                }
            }

            // Pass 2: ReferenceIntersector で接触面を検出して控除
            _progress?.Report("Pass 2: 接触面を検出中...");
            ContactFaceDetector.RefineContactFaces(_doc, _rayView, contexts);

            // Pass 3: 開口加算 + ElementResult 作成
            _progress?.Report("Pass 3: 集計中...");
            foreach (var ctx in contexts)
            {
                var elem = elemByContext[ctx.ElementId];
                try
                {
                    var er = BuildElementResult(elem, ctx, openingMap);
                    if (er != null) result.ElementResults.Add(er);
                }
                catch (Exception ex)
                {
                    result.Errors.Add(new ErrorLogEntry
                    {
                        ElementId = ctx.ElementId,
                        CategoryName = elem.Category?.Name ?? string.Empty,
                        ElementName = elem.Name,
                        ErrorKind = "AggregateError",
                        Message = ex.Message,
                    });
                }
            }

            Aggregate(result);
            return result;
        }

        private ContactFaceDetector.ElementFacesContext ClassifyElementFaces(
            Element elem, double? glFeet)
        {
            var solids = SolidUnionProcessor.GetSolids(elem);
            if (solids.Count == 0) return null;

            Solid unioned = SolidUnionProcessor.Union(solids);
            var finalSolids = unioned != null ? new List<Solid> { unioned } : solids;

            var (minZ, maxZ) = FaceClassifier.GetZRange(finalSolids);
            var faceInfos = FaceClassifier.ClassifyAll(finalSolids, glFeet, minZ, maxZ);

            // 基礎以外の最下面は FormworkRequired として扱う（梁底・柱底等は形枠必要）
            // 他要素との接触は ContactDetector が個別に検出する
            bool isFoundation = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Foundation;
            if (!isFoundation)
            {
                foreach (var fi in faceInfos)
                {
                    if (fi.FaceType == FaceType.DeductedBottom)
                        fi.FaceType = FaceType.FormworkRequired;
                }
            }

            BoundingBoxXYZ bb = null;
            try { bb = elem.get_BoundingBox(null); } catch { }

            return new ContactFaceDetector.ElementFacesContext
            {
                ElementId = elem.Id.IntegerValue,
                BB = bb,
                Faces = faceInfos,
            };
        }

        private ElementResult BuildElementResult(
            Element elem,
            ContactFaceDetector.ElementFacesContext ctx,
            Dictionary<int, OpeningProcessor.OpeningDelta> openingMap)
        {
            double formwork = 0, dedTop = 0, dedBottom = 0, dedContact = 0, inclined = 0;

            foreach (var fi in ctx.Faces)
            {
                double aM2 = FeetSqToM2(fi.Area);
                switch (fi.FaceType)
                {
                    case FaceType.FormworkRequired: formwork += aM2; break;
                    case FaceType.DeductedTop: dedTop += aM2; break;
                    case FaceType.DeductedBottom: dedBottom += aM2; break;
                    case FaceType.DeductedContact:
                    case FaceType.DeductedBelowGL: dedContact += aM2; break;
                    case FaceType.Inclined: inclined += aM2; break;
                }
            }

            double openingDeducted = 0;
            double openingAdded = 0;
            if (openingMap.TryGetValue(ctx.ElementId, out var od))
            {
                openingDeducted = FeetSqToM2(od.DeductedArea);
                openingAdded = FeetSqToM2(od.AddedEdgeArea);
                formwork = Math.Max(0, formwork - openingDeducted + openingAdded);
            }

            return new ElementResult
            {
                ElementId = ctx.ElementId,
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
        }

        /// <summary>
        /// 可視化のため、分類後の面を要素IDごとに提供する。
        /// View3D が必要（ReferenceIntersector のため）。
        /// </summary>
        internal static Dictionary<int, List<FaceClassifier.FaceInfo>> RecomputeFaces(
            Document doc,
            View3D rayView,
            FormworkResult result,
            FormworkSettings settings)
        {
            var map = new Dictionary<int, List<FaceClassifier.FaceInfo>>();
            double? glFeet = null;
            if (settings.UseGLDeduction)
                glFeet = UnitUtils.ConvertToInternalUnits(settings.GLElevationMeters, UnitTypeId.Meters);

            var contexts = new List<ContactFaceDetector.ElementFacesContext>();
            var elements = result.ElementResults
                .Select(er => doc.GetElement(new ElementId(er.ElementId)))
                .Where(e => e != null).ToList();

            foreach (var elem in elements)
            {
                var solids = SolidUnionProcessor.GetSolids(elem);
                if (solids.Count == 0) continue;
                var unioned = SolidUnionProcessor.Union(solids);
                var final = unioned != null ? new List<Solid> { unioned } : solids;
                var (minZ, maxZ) = FaceClassifier.GetZRange(final);
                var faces = FaceClassifier.ClassifyAll(final, glFeet, minZ, maxZ);

                bool isFoundation = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Foundation;
                if (!isFoundation)
                {
                    foreach (var f in faces)
                    {
                        if (f.FaceType == FaceType.DeductedBottom)
                            f.FaceType = FaceType.FormworkRequired;
                    }
                }

                BoundingBoxXYZ bb = null;
                try { bb = elem.get_BoundingBox(null); } catch { }

                contexts.Add(new ContactFaceDetector.ElementFacesContext
                {
                    ElementId = elem.Id.IntegerValue,
                    BB = bb,
                    Faces = faces,
                });
            }

            ContactFaceDetector.RefineContactFaces(doc, rayView, contexts);

            foreach (var ctx in contexts)
                map[ctx.ElementId] = ctx.Faces;

            return map;
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

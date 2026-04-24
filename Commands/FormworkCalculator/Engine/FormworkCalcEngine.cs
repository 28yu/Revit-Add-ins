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
    /// Pass 2: 幾何学的な接触面検出 → DeductedContact
    /// Pass 3: 開口加算・集計
    /// </summary>
    internal class FormworkCalcEngine
    {
        private readonly Document _doc;
        private readonly FormworkSettings _settings;
        private readonly View _activeView;
        private readonly IProgress<string> _progress;

        internal FormworkCalcEngine(
            Document doc,
            FormworkSettings settings,
            View activeView,
            IProgress<string> progress = null)
        {
            _doc = doc;
            _settings = settings;
            _activeView = activeView;
            _progress = progress;
        }

        internal FormworkResult Run()
        {
            var result = new FormworkResult();

            FormworkDebugLog.Initialize(_settings?.EnableDebugLog == true);
            try
            {
                return RunCore(result);
            }
            finally
            {
                FormworkDebugLog.Close();
            }
        }

        private FormworkResult RunCore(FormworkResult result)
        {
            _progress?.Report("要素を収集中...");
            var elements = ElementCollector.Collect(_doc, _settings, _activeView);
            result.ProcessedElementCount = elements.Count;
            FormworkDebugLog.Log($"Collected elements: {elements.Count}");
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

            // Pass 1 完了時: 各要素の面分類内訳をログ
            LogPass1Summary(contexts);

            // Pass 2: 幾何学的な直接検査で接触面を検出して控除
            _progress?.Report("Pass 2: 接触面を検出中...");
            ContactFaceDetector.RefineContactFaces(contexts);

            // Pass 2 完了時: 最終 FormworkRequired 面数
            LogPostPass2Summary(contexts);

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

        private static void LogPass1Summary(List<ContactFaceDetector.ElementFacesContext> contexts)
        {
            if (!FormworkDebugLog.Enabled) return;
            FormworkDebugLog.Section($"Pass 1: Face Classification (elements={contexts.Count})");
            int totalReq = 0, totalTop = 0, totalBot = 0, totalCon = 0, totalBGL = 0, totalInc = 0, totalErr = 0;
            foreach (var ctx in contexts)
            {
                int req = 0, top = 0, bot = 0, con = 0, bgl = 0, inc = 0, err = 0;
                foreach (var f in ctx.Faces)
                {
                    switch (f.FaceType)
                    {
                        case FaceType.FormworkRequired: req++; break;
                        case FaceType.DeductedTop: top++; break;
                        case FaceType.DeductedBottom: bot++; break;
                        case FaceType.DeductedContact: con++; break;
                        case FaceType.DeductedBelowGL: bgl++; break;
                        case FaceType.Inclined: inc++; break;
                        case FaceType.Error: err++; break;
                    }
                }
                totalReq += req; totalTop += top; totalBot += bot; totalCon += con;
                totalBGL += bgl; totalInc += inc; totalErr += err;
                FormworkDebugLog.Log(
                    $"  E{ctx.ElementId} ({ctx.Category}/{ctx.CategoryName}) " +
                    $"total={ctx.Faces.Count} Req={req} Top={top} Bot={bot} Con={con} BGL={bgl} Inc={inc} Err={err}");
            }
            FormworkDebugLog.Log(
                $"  TOTAL: Req={totalReq} Top={totalTop} Bot={totalBot} Con={totalCon} " +
                $"BGL={totalBGL} Inc={totalInc} Err={totalErr}");
            FormworkDebugLog.Flush();
        }

        private static void LogPostPass2Summary(List<ContactFaceDetector.ElementFacesContext> contexts)
        {
            if (!FormworkDebugLog.Enabled) return;
            int req = 0, con = 0;
            foreach (var ctx in contexts)
            {
                foreach (var f in ctx.Faces)
                {
                    if (f.FaceType == FaceType.FormworkRequired) req++;
                    else if (f.FaceType == FaceType.DeductedContact) con++;
                }
            }
            FormworkDebugLog.Section("Post Pass 2 Totals");
            FormworkDebugLog.Log($"  FormworkRequired: {req}");
            FormworkDebugLog.Log($"  DeductedContact:  {con}");
            FormworkDebugLog.Flush();
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
            bool isSlab = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Slab;

            if (!isFoundation)
            {
                foreach (var fi in faceInfos)
                {
                    if (fi.FaceType == FaceType.DeductedBottom)
                        fi.FaceType = FaceType.FormworkRequired;
                }
            }

            // 床(スラブ)の天端は常に型枠不要: 最上部以外の段差・凹みを含め全ての上向き水平面を除外
            if (isSlab)
            {
                foreach (var fi in faceInfos)
                {
                    if (fi.Normal != null && fi.Normal.Z > 0.99 &&
                        fi.FaceType == FaceType.FormworkRequired)
                    {
                        fi.FaceType = FaceType.DeductedTop;
                    }
                }
            }

            BoundingBoxXYZ bb = null;
            try { bb = elem.get_BoundingBox(null); } catch { }

            return new ContactFaceDetector.ElementFacesContext
            {
                ElementId = elem.Id.IntegerValue,
                Category = ElementCollector.ToCategoryGroup(elem),
                CategoryName = elem.Category?.Name ?? string.Empty,
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
        /// 幾何学的な接触検出を使うので View3D は不要。
        /// </summary>
        internal static Dictionary<int, List<FaceClassifier.FaceInfo>> RecomputeFaces(
            Document doc,
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
                bool isSlab = ElementCollector.ToCategoryGroup(elem) == CategoryGroup.Slab;

                if (!isFoundation)
                {
                    foreach (var f in faces)
                    {
                        if (f.FaceType == FaceType.DeductedBottom)
                            f.FaceType = FaceType.FormworkRequired;
                    }
                }

                if (isSlab)
                {
                    foreach (var f in faces)
                    {
                        if (f.Normal != null && f.Normal.Z > 0.99 &&
                            f.FaceType == FaceType.FormworkRequired)
                        {
                            f.FaceType = FaceType.DeductedTop;
                        }
                    }
                }

                BoundingBoxXYZ bb = null;
                try { bb = elem.get_BoundingBox(null); } catch { }

                contexts.Add(new ContactFaceDetector.ElementFacesContext
                {
                    ElementId = elem.Id.IntegerValue,
                    Category = ElementCollector.ToCategoryGroup(elem),
                    CategoryName = elem.Category?.Name ?? string.Empty,
                    BB = bb,
                    Faces = faces,
                });
            }

            ContactFaceDetector.RefineContactFaces(contexts);

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

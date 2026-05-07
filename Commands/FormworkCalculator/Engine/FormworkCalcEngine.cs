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
            // 注: ここで Close() しない。後続の CreateVisualization / CreateSchedule もログ出力するため、
            // クローズは Command 側 (FormworkCalculatorCommand) で全処理完了後に行う。
            return RunCore(result);
        }

        private FormworkResult RunCore(FormworkResult result)
        {
            _progress?.Report("要素を収集中...");
            var collection = ElementCollector.CollectAndClassify(_doc, _settings, _activeView);
            var elements = collection.Targets;
            result.ProcessedElementCount = elements.Count;
            FormworkDebugLog.Log(
                $"Collected elements: targets={elements.Count} " +
                $"excluded={collection.Excluded.Count}");

            // 除外要素を ExcludedResult として記録 (集計には含めない)
            foreach (var ex in collection.Excluded)
            {
                var e = ex.Element;
                result.ExcludedResults.Add(new ExcludedResult
                {
                    ElementId = e.Id.IntegerValue,
                    ElementName = e.Name ?? string.Empty,
                    Category = ElementCollector.ToCategoryGroup(e),
                    CategoryName = e.Category?.Name ?? string.Empty,
                    Kind = ex.Kind,
                    DetectionLayer = ex.Layer,
                    DetectionReason = ex.Reason,
                });
            }

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

            // Pass 2b: WallSweep (スイープ・リビール) の host 壁面を直接 DeductedContact 化
            // (Reveal 等で WallSweep 自身がソリッドを持たないケースに対応)
            WallSweepFaceDeductor.DeductWallFacesNearSweeps(_doc, contexts);

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

            // 壁スイープ・リビールを ElementResults から除外し ExcludedResults に移す
            // (Pass 2 で Wall の top 面に対する contact deduction を有効化した後に実施)
            MoveWallSweepsToExcluded(_doc, result);

            // 診断ログ: 各要素の集計結果を一覧出力 (⚠️ マーカーで疑わしい要素を強調)
            LogElementDiagnostics(_doc, contexts, result);

            Aggregate(result);
            return result;
        }

        /// <summary>
        /// 各要素の集計結果と面の状態を診断ログに出力する。
        /// 疑わしいパターンには ⚠️ マーカーを付け、原因特定の手がかりとする。
        /// 検索しやすいように 1 要素 1 行で出力。
        /// </summary>
        private static void LogElementDiagnostics(
            Document doc,
            List<ContactFaceDetector.ElementFacesContext> contexts,
            FormworkResult result)
        {
            if (!FormworkDebugLog.Enabled) return;

            var ctxById = new Dictionary<int, ContactFaceDetector.ElementFacesContext>();
            foreach (var c in contexts) ctxById[c.ElementId] = c;

            FormworkDebugLog.Section($"Element Diagnostics ({result.ElementResults.Count} elements)");
            FormworkDebugLog.Log("  fmt: [ElemDiag] E<id> (cat) 'name' type='type' formwork=Xm² faces=R/T/B/C/B(GL)/I parts=N");
            FormworkDebugLog.Log("       ⚠️ markers: ZERO=formwork≈0, MOSTLY_DED=控除>>formwork, ALL_PARTIAL=全FormworkRequired面が部分接触");

            foreach (var er in result.ElementResults)
            {
                if (!ctxById.TryGetValue(er.ElementId, out var ctx)) continue;

                string typeName = "?";
                try
                {
                    var elem = doc.GetElement(new ElementId(er.ElementId));
                    if (elem != null)
                    {
                        var t = doc.GetElement(elem.GetTypeId());
                        typeName = t?.Name ?? "?";
                    }
                }
                catch { }

                int reqCount = 0, topCount = 0, botCount = 0, conCount = 0, bglCount = 0, incCount = 0;
                int partialFaceCount = 0;
                int allReqFaces = 0;
                foreach (var fi in ctx.Faces)
                {
                    switch (fi.FaceType)
                    {
                        case FaceType.FormworkRequired:
                            reqCount++;
                            allReqFaces++;
                            if (fi.PartialContacts.Count > 0) partialFaceCount++;
                            break;
                        case FaceType.DeductedTop: topCount++; break;
                        case FaceType.DeductedBottom: botCount++; break;
                        case FaceType.DeductedContact: conCount++; break;
                        case FaceType.DeductedBelowGL: bglCount++; break;
                        case FaceType.Inclined: incCount++; break;
                    }
                }

                // ⚠️ マーカー判定 (見やすくするため絞り込み)
                var marks = new List<string>();
                // CONTACT_HEAVY: dedTop/dedBot は基礎・スラブで自然に大きいため除外し、
                //                純粋な contact 控除 (dedCon) が formwork の 3 倍以上の場合のみ警告
                if (er.FormworkArea > 0 && er.DeductedContactArea > er.FormworkArea * 3)
                    marks.Add("⚠️CONTACT_HEAVY");
                // ZERO: FormworkRequired 面が無く、contact 面が多い (柱の埋め込み等)
                if (er.FormworkArea < 0.01 && conCount >= 4)
                    marks.Add("⚠️ZERO_EMBEDDED");
                else if (er.FormworkArea < 0.01 && reqCount > 0)
                    marks.Add("⚠️ZERO");
                if (allReqFaces > 0 && partialFaceCount == allReqFaces)
                    marks.Add("⚠️ALL_PARTIAL");
                string marker = marks.Count > 0 ? " " + string.Join("|", marks) : "";

                // 寸法情報 (BB から要素の概算寸法を算出)
                string dim = "?";
                try
                {
                    if (ctx.BB != null)
                    {
                        double lx = (ctx.BB.Max.X - ctx.BB.Min.X) * 304.8;  // ft → mm
                        double ly = (ctx.BB.Max.Y - ctx.BB.Min.Y) * 304.8;
                        double lz = (ctx.BB.Max.Z - ctx.BB.Min.Z) * 304.8;
                        dim = $"{lx:F0}x{ly:F0}x{lz:F0}";
                    }
                }
                catch { }

                string cat = CategoryShort(er.Category);
                FormworkDebugLog.Log(
                    $"[ElemDiag] E{er.ElementId} ({cat}) '{er.ElementName}' type='{typeName}' " +
                    $"dim={dim}mm formwork={er.FormworkArea:F2}m² faces={reqCount}/{topCount}/{botCount}/{conCount}/{bglCount}/{incCount} " +
                    $"parts={partialFaceCount} dedTop={er.DeductedTopArea:F2} dedBot={er.DeductedBottomArea:F2} " +
                    $"dedCon={er.DeductedContactArea:F2}{marker}");
            }
            FormworkDebugLog.Flush();
        }

        private static string CategoryShort(CategoryGroup cg)
        {
            switch (cg)
            {
                case CategoryGroup.Column: return "柱";
                case CategoryGroup.Beam: return "梁";
                case CategoryGroup.Wall: return "壁";
                case CategoryGroup.Slab: return "スラブ";
                case CategoryGroup.Foundation: return "基礎";
                case CategoryGroup.Stairs: return "階段";
                default: return "他";
            }
        }

        /// <summary>
        /// WallSweep (壁スイープ・リビール) を ElementResults から ExcludedResults に移動する。
        /// 早期除外せず Pass 1/2 に参加させる理由: Wall の top 面が reveal で切り取られた
        /// 部分を contact detection によって DeductedContact 化するため。
        /// </summary>
        private static void MoveWallSweepsToExcluded(Document doc, FormworkResult result)
        {
            var toMove = new List<ElementResult>();
            foreach (var er in result.ElementResults)
            {
                var elem = doc.GetElement(new ElementId(er.ElementId));
                if (elem is WallSweep) toMove.Add(er);
            }
            if (toMove.Count == 0) return;

            foreach (var er in toMove)
            {
                var elem = doc.GetElement(new ElementId(er.ElementId));
                result.ElementResults.Remove(er);
                result.ExcludedResults.Add(new ExcludedResult
                {
                    ElementId = er.ElementId,
                    ElementName = er.ElementName,
                    Category = CategoryGroup.Wall,
                    CategoryName = elem?.Category?.Name ?? string.Empty,
                    Kind = ExclusionKind.WallSweep,
                    DetectionLayer = "WallSweep",
                    DetectionReason = "壁スイープ・リビール (post-processed after contact detection)",
                });
            }
            FormworkDebugLog.Log($"  Moved {toMove.Count} WallSweep elements to ExcludedResults");
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
                    case FaceType.FormworkRequired:
                        {
                            // 部分接触の面積を控除する。
                            // Naive sum (ContactArea を単純合計) は接触領域同士が重なって
                            // いると過大評価され、面積が 0 にクランプされてしまう。
                            // Clipper を使った 2D 矩形差分で正確な残面積を計算する。
                            string clipperStatus;
                            double effectiveFeetSq = ComputeAndSetEffectiveArea(fi, out clipperStatus);

                            // 部分接触の詳細をログ (問題追跡用)
                            if (fi.PartialContacts.Count > 0 && FormworkDebugLog.Enabled)
                            {
                                var sb = new System.Text.StringBuilder();
                                sb.Append($"  [FaceDiag] E{ctx.ElementId} face[{ctx.Faces.IndexOf(fi)}] ");
                                sb.Append($"area={fi.Area:F4} partials={fi.PartialContacts.Count} [");
                                foreach (var pc in fi.PartialContacts)
                                    sb.Append($" E{pc.OtherElementId}:a={pc.ContactArea:F4}");
                                sb.Append($" ] {clipperStatus} eff={effectiveFeetSq:F4}");
                                FormworkDebugLog.Log(sb.ToString());
                            }

                            formwork += fi.EffectiveAreaM2;
                            // 部分接触分も控除面積として集計に加える
                            double partialDed = Math.Max(0, fi.Area - effectiveFeetSq);
                            dedContact += FeetSqToM2(partialDed);
                        }
                        break;
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

            bool hasPartialContact = false;
            foreach (var fi in ctx.Faces)
            {
                if (fi.FaceType == FaceType.FormworkRequired && fi.PartialContacts.Count > 0)
                { hasPartialContact = true; break; }
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
                HasPartialContact = hasPartialContact,
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

            // 各面の有効面積を計算して FaceInfo に保存。
            // RecomputeFaces は新しい FaceInfo インスタンスを作るため、BuildElementResult
            // 側で設定した EffectiveAreaM2 はここには来ない。可視化で使うために再計算する。
            foreach (var ctx in contexts)
            {
                foreach (var fi in ctx.Faces)
                {
                    if (fi.FaceType == FaceType.FormworkRequired)
                    {
                        string _;
                        ComputeAndSetEffectiveArea(fi, out _);
                    }
                }
                map[ctx.ElementId] = ctx.Faces;
            }

            return map;
        }

        /// <summary>
        /// FormworkRequired 面の有効面積 (部分接触控除後) を計算し、fi.EffectiveAreaM2 に保存する。
        /// 戻り値は feet² 単位の有効面積 (ログ出力用)、out clipperStatus は診断文字列。
        /// </summary>
        private static double ComputeAndSetEffectiveArea(
            FaceClassifier.FaceInfo fi, out string clipperStatus)
        {
            clipperStatus = "no-partial";
            double effectiveFeetSq = fi.Area;
            if (fi.PartialContacts.Count > 0)
            {
                var clip = PartialContactClipper.TryClip(fi);
                if (clip.Success && clip.Solids.Count > 0)
                {
                    double area = 0;
                    foreach (var s in clip.Solids)
                        area += s.Volume / PartialContactClipper.ThicknessFeet;
                    effectiveFeetSq = area;
                    clipperStatus = "clipper-OK";
                }
                else
                {
                    double partial = 0;
                    foreach (var pc in fi.PartialContacts) partial += pc.ContactArea;
                    double rawSum = partial;
                    partial = Math.Min(partial, fi.Area * 0.95);
                    effectiveFeetSq = Math.Max(0, fi.Area - partial);
                    clipperStatus = $"clipper-FAIL({clip.FailReason}) rawSum={rawSum:F4}";
                }
            }
            fi.EffectiveAreaM2 = FeetSqToM2(effectiveFeetSq);
            return effectiveFeetSq;
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

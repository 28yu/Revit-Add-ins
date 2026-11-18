using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Tools28.Localization;

namespace Tools28.Commands.LineMerge
{
    /// <summary>
    /// 直線まとめ: 選択した詳細線・モデル線（CAD を分解した線を含む）のうち、
    /// 同一直線上にある直線どうしを1本にまとめる。角度のある線はつながない。
    /// いちばん長い線を伸ばして残し、ほかの線を削除する（残した線の寸法などを消さないため）。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class LineMergeCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
                if (ids.Count == 0)
                {
                    try
                    {
                        ids = uidoc.Selection.PickObjects(ObjectType.Element, new LineSelectionFilter(),
                                Loc.S("LineMerge.PickPrompt"))
                            .Select(r => r.ElementId).ToList();
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        return Result.Cancelled;
                    }
                }

                // 選択を「まとめられる直線」と、それ以外に分ける
                var groups = new Dictionary<string, List<MergeSegment<ElementId>>>();
                int lineCount = 0, skippedCurve = 0, skippedLocked = 0, cadCount = 0;
                foreach (ElementId id in ids)
                {
                    Element e = doc.GetElement(id);
                    if (e is ImportInstance) { cadCount++; continue; }
                    if (!(e is CurveElement ce) || !IsTargetLine(ce)) continue;

                    if (!(ce.GeometryCurve is Line line) || !line.IsBound) { skippedCurve++; continue; }
                    if (ce.Pinned || ce.GroupId != ElementId.InvalidElementId) { skippedLocked++; continue; }

                    string key = GroupKey(ce);
                    if (!groups.TryGetValue(key, out var list)) groups[key] = list = new List<MergeSegment<ElementId>>();
                    list.Add(new MergeSegment<ElementId>
                    {
                        Tag = id,
                        Start = ToVec(line.GetEndPoint(0)),
                        End = ToVec(line.GetEndPoint(1)),
                    });
                    lineCount++;
                }

                if (lineCount == 0)
                {
                    string msg = Loc.S("LineMerge.NoLines");
                    if (cadCount > 0) msg += "\n\n" + Loc.S("LineMerge.CadHint");
                    TaskDialog.Show(Loc.S("LineMerge.Title"), msg);
                    return Result.Cancelled;
                }

                var results = groups.Values.SelectMany(CollinearLineMerger.Merge).ToList();
                if (results.Count == 0)
                {
                    TaskDialog.Show(Loc.S("LineMerge.Title"),
                        Loc.S("LineMerge.NothingToMerge") + Notes(skippedCurve, skippedLocked, cadCount));
                    return Result.Cancelled;
                }

                int removed = 0, failed = 0;
                using (var t = new Transaction(doc, Loc.S("LineMerge.Txn")))
                {
                    t.Start();
                    SetQuietFailureHandling(t);

                    // 削除はまとめて1回で行う（組ごとに Delete を呼ぶと、そのたびに Revit がモデルを
                    // 再計算して非常に遅くなるため）。まとめて消せなかったときだけ組ごとにやり直す。
                    var merged = new List<MergeResult<ElementId>>();
                    try
                    {
                        doc.Delete(results.SelectMany(r => r.Remove).Select(s => s.Tag).ToList());
                        merged.AddRange(results);
                    }
                    catch (Exception ex)
                    {
                        DiagLog.Write($"LineMerge: 一括削除に失敗したため組ごとに削除します: {ex.Message}");
                        foreach (var r in results)
                        {
                            try
                            {
                                doc.Delete(r.Remove.Select(s => s.Tag).ToList());
                                merged.Add(r);
                            }
                            catch (Exception ex2)
                            {
                                failed++;
                                DiagLog.Write($"LineMerge: 1組を削除できませんでした: {ex2.Message}");
                            }
                        }
                    }

                    // 残す線を、まとめた長さまで伸ばす
                    foreach (var r in merged)
                    {
                        try
                        {
                            var keep = (CurveElement)doc.GetElement(r.Keep.Tag);
                            keep.SetGeometryCurve(Line.CreateBound(ToXyz(r.NewStart), ToXyz(r.NewEnd)), true);
                        }
                        catch (Exception ex)
                        {
                            // 伸ばせなくても、重なっていた線は消えているので見た目はほぼ変わらない
                            DiagLog.Write($"LineMerge: 線を伸ばせませんでした: {ex.Message}");
                        }
                        removed += r.Remove.Count;
                    }
                    t.Commit();
                }

                string done = string.Format(Loc.S("LineMerge.Done"), lineCount, lineCount - removed, results.Count - failed);
                if (failed > 0) done += "\n" + string.Format(Loc.S("LineMerge.Failed"), failed);
                TaskDialog.Show(Loc.S("LineMerge.Title"), done + Notes(skippedCurve, skippedLocked, cadCount));
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show(Loc.S("Common.Error"), string.Format(Loc.S("LineMerge.Error"), ex.Message));
                return Result.Failed;
            }
        }

        /// <summary>詳細線とモデル線だけを対象にする（部屋境界線・エリア境界線などは触らない）。</summary>
        internal static bool IsTargetLine(CurveElement ce)
        {
            return ce.CurveElementType == CurveElementType.DetailCurve
                || ce.CurveElementType == CurveElementType.ModelCurve;
        }

        /// <summary>
        /// 同じ組にまとめてよい線の区分。詳細線は同じビュー、モデル線は同じ作業面のものだけをまとめる
        /// （別のビューや別の面の線を1本にはできないため）。
        /// </summary>
        private static string GroupKey(CurveElement ce)
        {
            if (ce.CurveElementType == CurveElementType.DetailCurve)
                return "D:" + ce.OwnerViewId.IntValue();
            return "M:" + (ce.SketchPlane?.Id ?? ElementId.InvalidElementId).IntValue();
        }

        private static string Notes(int curve, int locked, int cad)
        {
            string s = "";
            if (curve > 0) s += "\n" + string.Format(Loc.S("LineMerge.SkippedCurve"), curve);
            if (locked > 0) s += "\n" + string.Format(Loc.S("LineMerge.SkippedLocked"), locked);
            if (cad > 0) s += "\n" + Loc.S("LineMerge.CadHint");
            return s.Length > 0 ? "\n" + s : s;
        }

        /// <summary>
        /// 線を伸ばしたときの「わずかに軸からずれています」などの警告で、確定のたびに
        /// 警告画面が出て止まらないよう、警告だけは黙って消す（エラーはそのまま Revit に任せる）。
        /// </summary>
        private static void SetQuietFailureHandling(Transaction t)
        {
            var opt = t.GetFailureHandlingOptions();
            opt.SetFailuresPreprocessor(new SwallowWarnings());
            t.SetFailureHandlingOptions(opt);
        }

        private sealed class SwallowWarnings : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
            {
                foreach (var f in a.GetFailureMessages())
                {
                    if (f.GetSeverity() == FailureSeverity.Warning) a.DeleteWarning(f);
                }
                return FailureProcessingResult.Continue;
            }
        }

        private static Vec3 ToVec(XYZ p) => new Vec3(p.X, p.Y, p.Z);
        private static XYZ ToXyz(Vec3 v) => new XYZ(v.X, v.Y, v.Z);

        /// <summary>選択するときは詳細線・モデル線だけをクリックできるようにする。</summary>
        private class LineSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is CurveElement ce && IsTargetLine(ce);
            public bool AllowReference(Reference reference, XYZ position) => false;
        }
    }
}

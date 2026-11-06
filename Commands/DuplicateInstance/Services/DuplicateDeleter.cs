using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.DuplicateInstance.Models;
using Tools28.Localization;

namespace Tools28.Commands.DuplicateInstance.Services
{
    /// <summary>削除の結果（削除できた要素・できなかった要素と原因・Revit の警告）</summary>
    public class DeleteResult
    {
        public bool Cancelled { get; set; }
        public HashSet<long> DeletedIds { get; } = new HashSet<long>();
        public List<KeyValuePair<DuplicateRow, string>> Failed { get; } = new List<KeyValuePair<DuplicateRow, string>>();
        public List<string> Warnings { get; } = new List<string>();
        public string FatalError { get; set; }
    }

    /// <summary>進み具合の通知先（ダイアログの進捗ウィンドウ）</summary>
    public interface IDeleteProgress
    {
        void Report(int done, int total, string status, bool force = false);
        void BeginCommit();
        bool IsCancelRequested { get; }
    }

    /// <summary>
    /// チェックした要素を削除する。大量でも途中経過を出せるよう少しずつ消し、
    /// 消せなかった要素はその原因を記録する。
    /// 確定時にエラーが出た場合は、エラーの原因になった要素だけを外してやり直す。
    /// </summary>
    public class DuplicateDeleter
    {
        private const int ChunkSize = 100;
        private const int MaxAttempts = 5;

        public DeleteResult Delete(Document doc, List<DuplicateRow> targets, IDeleteProgress progress)
        {
            var result = new DeleteResult();
            var failed = new Dictionary<long, string>();

            // --- 事前確認（削除する前から消せないと分かるもの） ---
            var remaining = new List<DuplicateRow>();
            foreach (var r in targets)
            {
                string reason = PreCheck(doc, r);
                if (reason != null) failed[r.IdValue] = reason;
                else remaining.Add(r);
            }

            for (int attempt = 1; attempt <= MaxAttempts && remaining.Count > 0; attempt++)
            {
                string status = attempt == 1
                    ? Loc.S("DupInst.Progress.Deleting")
                    : string.Format(Loc.S("DupInst.Progress.Retry"), attempt - 1);
                progress.Report(0, remaining.Count, status, force: true);

                var collector = new FailureCollector();
                TransactionStatus ts;
                using (var t = new Transaction(doc, Loc.S("DupInst.Txn.Delete")))
                {
                    t.Start();
                    var fho = t.GetFailureHandlingOptions();
                    fho = fho.SetForcedModalHandling(false);
                    fho = fho.SetClearAfterRollback(true);
                    fho = fho.SetFailuresPreprocessor(collector);
                    t.SetFailureHandlingOptions(fho);

                    int done = 0;
                    for (int i = 0; i < remaining.Count; i += ChunkSize)
                    {
                        if (progress.IsCancelRequested)
                        {
                            t.RollBack();
                            result.Cancelled = true;
                            return result;
                        }
                        var chunk = remaining.Skip(i).Take(ChunkSize)
                                             .Where(r => !failed.ContainsKey(r.IdValue) && doc.GetElement(r.Id) != null)
                                             .ToList();
                        if (chunk.Count > 0)
                        {
                            try
                            {
                                doc.Delete(chunk.Select(r => r.Id).ToList());
                            }
                            catch
                            {
                                // まとめて消せないときは1つずつ消し、消せなかった要素の原因を記録する
                                foreach (var r in chunk)
                                {
                                    if (doc.GetElement(r.Id) == null) continue;   // 他の要素と一緒に消えた
                                    try { doc.Delete(r.Id); }
                                    catch (Exception ex) { failed[r.IdValue] = ExceptionReason(doc, r, ex); }
                                }
                            }
                        }
                        done = Math.Min(remaining.Count, i + ChunkSize);
                        progress.Report(done, remaining.Count, null);
                    }

                    progress.Report(remaining.Count, remaining.Count, null, force: true);
                    if (progress.IsCancelRequested)
                    {
                        t.RollBack();
                        result.Cancelled = true;
                        return result;
                    }
                    progress.BeginCommit();
                    try { ts = t.Commit(); }
                    catch (Exception ex)
                    {
                        result.FatalError = ex.Message;
                        ts = TransactionStatus.RolledBack;
                    }
                }

                foreach (var w in collector.Warnings)
                    if (!result.Warnings.Contains(w)) result.Warnings.Add(w);

                if (ts == TransactionStatus.Committed)
                {
                    foreach (var r in remaining)
                    {
                        if (failed.ContainsKey(r.IdValue)) continue;
                        if (doc.GetElement(r.Id) == null) result.DeletedIds.Add(r.IdValue);
                        else failed[r.IdValue] = Loc.S("DupInst.Reason.Remained");
                    }
                    remaining.Clear();
                    break;
                }

                // 確定できなかった: エラーの原因になった要素を外してやり直す
                bool matched = false;
                foreach (var err in collector.Errors)
                {
                    foreach (var r in remaining)
                    {
                        if (failed.ContainsKey(r.IdValue)) continue;
                        if (err.Value.Contains(r.IdValue))
                        {
                            failed[r.IdValue] = string.Format(Loc.S("DupInst.Reason.Error"), err.Key);
                            matched = true;
                        }
                    }
                }
                if (!matched)
                {
                    // どの要素が原因か分からない場合は、残り全部を「確定できなかった」とする
                    string text = collector.Errors.Count > 0
                        ? string.Join(" / ", collector.Errors.Select(e => e.Key).Distinct())
                        : (result.FatalError ?? Loc.S("DupInst.Reason.Unknown"));
                    foreach (var r in remaining)
                        if (!failed.ContainsKey(r.IdValue))
                            failed[r.IdValue] = string.Format(Loc.S("DupInst.Reason.Error"), text);
                    remaining.Clear();
                    break;
                }
                remaining = remaining.Where(r => !failed.ContainsKey(r.IdValue)).ToList();
                // 1つずつ消して失敗した要素も、やり直しではもう一度試す価値が無いので外したまま
            }

            foreach (var r in targets)
                if (failed.TryGetValue(r.IdValue, out var reason))
                    result.Failed.Add(new KeyValuePair<DuplicateRow, string>(r, reason));
            return result;
        }

        /// <summary>削除する前に分かる「消せない理由」。問題なければ null。</summary>
        private static string PreCheck(Document doc, DuplicateRow r)
        {
            var e = doc.GetElement(r.Id);
            if (e == null) return Loc.S("DupInst.Reason.Missing");

            if (doc.IsWorkshared)
            {
                try
                {
                    var st = WorksharingUtils.GetCheckoutStatus(doc, r.Id, out string owner);
                    if (st == CheckoutStatus.OwnedByOtherUser)
                        return string.Format(Loc.S("DupInst.Reason.OwnedByOther"), owner);
                }
                catch { }
                try
                {
                    var us = WorksharingUtils.GetModelUpdatesStatus(doc, r.Id);
                    if (us == ModelUpdatesStatus.UpdatedInCentral || us == ModelUpdatesStatus.DeletedInCentral)
                        return Loc.S("DupInst.Reason.NotCurrent");
                }
                catch { }
            }
            return null;
        }

        private static string ExceptionReason(Document doc, DuplicateRow r, Exception ex)
        {
            string msg = ex.Message ?? ex.GetType().Name;
            Element e = null;
            try { e = doc.GetElement(r.Id); } catch { }

            if (e != null && e.GroupId != null && e.GroupId != ElementId.InvalidElementId)
                return string.Format(Loc.S("DupInst.Reason.InGroup"), msg);
            if (ex is Autodesk.Revit.Exceptions.ModificationForbiddenException)
                return string.Format(Loc.S("DupInst.Reason.Forbidden"), msg);
            return string.Format(Loc.S("DupInst.Reason.Refused"), msg);
        }

        /// <summary>
        /// 確定時の警告は画面に出さずに記録して消す。エラーは原因の要素 ID と一緒に記録し、
        /// 確定を取り消す（取り消し後に、その要素を外してやり直すため）。
        /// </summary>
        private class FailureCollector : IFailuresPreprocessor
        {
            public List<string> Warnings { get; } = new List<string>();
            public List<KeyValuePair<string, HashSet<long>>> Errors { get; } = new List<KeyValuePair<string, HashSet<long>>>();

            public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
            {
                bool hasError = false;
                try
                {
                    foreach (var f in a.GetFailureMessages())
                    {
                        string text = f?.GetDescriptionText() ?? "";
                        if (f.GetSeverity() == FailureSeverity.Warning)
                        {
                            if (!string.IsNullOrWhiteSpace(text) && !Warnings.Contains(text)) Warnings.Add(text);
                            continue;
                        }
                        hasError = true;
                        var ids = new HashSet<long>();
                        foreach (var id in f.GetFailingElementIds() ?? new List<ElementId>()) ids.Add(id.LongValue());
                        foreach (var id in f.GetAdditionalElementIds() ?? new List<ElementId>()) ids.Add(id.LongValue());
                        Errors.Add(new KeyValuePair<string, HashSet<long>>(text, ids));
                    }
                }
                catch { }

                if (hasError) return FailureProcessingResult.ProceedWithRollBack;
                a.DeleteAllWarnings();
                return FailureProcessingResult.Continue;
            }
        }
    }
}

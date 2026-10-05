using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using Autodesk.Revit.DB;
using Tools28.Commands.ExcelExportImport.Services;
using Tools28.Localization;

namespace Tools28.Commands.ParameterValueReplace.Services
{
    /// <summary>値の探し方</summary>
    public enum MatchMode
    {
        Exact,      // 値全体が一致（前後の空白は無視）
        Contains    // 値の一部に含まれる（含まれる部分だけを置き換える）
    }

    /// <summary>探す範囲</summary>
    public enum ReplaceScope
    {
        EntireProject,
        ActiveView,
        Selection
    }

    /// <summary>
    /// 置換対象のまとまり（カテゴリ × パラメータ × インスタンス/タイプ）。ダイアログの1行。
    /// </summary>
    public class ReplaceGroup : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public string CategoryName { get; set; }
        public string ParameterName { get; set; }
        public bool IsType { get; set; }

        /// <summary>同名パラメータを区別するための種別（組み込み/共有/プロジェクト）</summary>
        public string KindText { get; set; }

        /// <summary>このまとまりで置き換えるパラメータ（要素ごと）</summary>
        public List<Parameter> Targets { get; } = new List<Parameter>();

        /// <summary>見つかった値の例（部分一致のときに確認用）</summary>
        public string SampleValue { get; set; }

        public int Count => Targets.Count;

        public string ScopeText => IsType ? Loc.S("ValueReplace.Scope.Type") : Loc.S("ValueReplace.Scope.Instance");

        private bool _isSelected = true;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }
    }

    /// <summary>置換の結果</summary>
    public class ReplaceResult
    {
        public int Success;
        public int Failed;
        public int Skipped;   // 実行時点で値が変わっていて対象外になったもの
        public bool Cancelled;
    }

    /// <summary>
    /// モデル内の文字パラメータから指定の値を探し、別の値（空欄可）に置き換えるサービス。
    ///
    /// 【目的】Excel 連携（書き出し→編集→読み込み）を使わずに、モデル全体の値を一括で置き換える。
    /// 大容量モデルでは数百万セルの書き出し・読み込みに非常に時間がかかるため、
    /// 「この値を消す（例: 未使用 → 空欄）」のような要素に依存しない変更は Revit 上で直接行う。
    ///
    /// 【対象】書き込み可能な文字（String）パラメータのみ。数値・要素参照・読み取り専用は対象外。
    /// 【画面を止めない工夫】走査・書き込みとも反復子にして、一定件数ごとに呼び出し側へ制御を返す。
    /// </summary>
    public static class ValueReplaceService
    {
        /// <summary>この件数の要素ごとに走査の反復子が yield する</summary>
        private const int ScanYieldEvery = 500;

        /// <summary>1つのトランザクションで書き込む件数（失敗時の巻き戻し範囲を小さくし、進み具合を表示するため）</summary>
        public const int WriteBatchSize = 2000;

        /// <summary>値が条件に一致するか</summary>
        public static bool IsMatch(string value, string find, MatchMode mode)
        {
            if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(find)) return false;
            return mode == MatchMode.Exact
                ? string.Equals(value.Trim(), find.Trim(), StringComparison.Ordinal)
                : value.IndexOf(find, StringComparison.Ordinal) >= 0;
        }

        /// <summary>置き換え後の値を作る</summary>
        public static string BuildNewValue(string current, string find, string replace, MatchMode mode)
        {
            replace = replace ?? "";
            return mode == MatchMode.Exact ? replace : (current ?? "").Replace(find, replace);
        }

        /// <summary>
        /// 範囲内の要素を走査して、値が一致するパラメータをまとまりごとに集める（反復子）。
        /// 処理済み要素数を yield する。結果は groups に入る。
        /// </summary>
        public static IEnumerable<int> Scan(
            Document doc, string find, MatchMode mode, ReplaceScope scope, bool includeTypes,
            View activeView, ICollection<ElementId> selectionIds,
            Dictionary<string, ReplaceGroup> groups, CancellationToken ct)
        {
            var instances = CollectInstances(doc, scope, activeView, selectionIds);
            var typeIds = new HashSet<ElementId>();
            int processed = 0;

            foreach (var e in instances)
            {
                if (ct.IsCancellationRequested) yield break;

                ScanElement(e, false, find, mode, groups);

                if (includeTypes)
                {
                    ElementId tid = null;
                    try { tid = e.GetTypeId(); } catch { }
                    if (tid != null && tid != ElementId.InvalidElementId)
                        typeIds.Add(tid);
                }

                processed++;
                if (processed % ScanYieldEvery == 0)
                    yield return processed;
            }

            if (includeTypes)
            {
                // プロジェクト全体ならインスタンスの無いタイプも対象にする
                if (scope == ReplaceScope.EntireProject)
                {
                    foreach (var t in new FilteredElementCollector(doc).WhereElementIsElementType())
                    {
                        if (HasCategory(t)) typeIds.Add(t.Id);
                    }
                }

                foreach (var tid in typeIds)
                {
                    if (ct.IsCancellationRequested) yield break;

                    Element t = null;
                    try { t = doc.GetElement(tid); } catch { }
                    if (t != null)
                        ScanElement(t, true, find, mode, groups);

                    processed++;
                    if (processed % ScanYieldEvery == 0)
                        yield return processed;
                }
            }

            yield return processed;
        }

        private static IEnumerable<Element> CollectInstances(
            Document doc, ReplaceScope scope, View activeView, ICollection<ElementId> selectionIds)
        {
            FilteredElementCollector collector;
            switch (scope)
            {
                case ReplaceScope.ActiveView:
                    if (activeView == null) return Enumerable.Empty<Element>();
                    collector = new FilteredElementCollector(doc, activeView.Id);
                    break;
                case ReplaceScope.Selection:
                    if (selectionIds == null || selectionIds.Count == 0) return Enumerable.Empty<Element>();
                    collector = new FilteredElementCollector(doc, selectionIds);
                    break;
                default:
                    collector = new FilteredElementCollector(doc);
                    break;
            }

            // カテゴリの無い内部要素（設定・ビュー内部オブジェクト等）は対象外
            return collector.WhereElementIsNotElementType().Where(HasCategory);
        }

        private static bool HasCategory(Element e)
        {
            try { return e?.Category != null; }
            catch { return false; }
        }

        private static void ScanElement(Element e, bool isType, string find, MatchMode mode,
                                        Dictionary<string, ReplaceGroup> groups)
        {
            string categoryName;
            try { categoryName = e.Category?.Name ?? ""; }
            catch { categoryName = ""; }

            ParameterSet ps;
            try { ps = e.Parameters; }
            catch { return; }

            foreach (Parameter p in ps)
            {
                try
                {
                    if (p == null || p.StorageType != StorageType.String || p.IsReadOnly || !p.HasValue)
                        continue;

                    string value = p.AsString();
                    if (!IsMatch(value, find, mode))
                        continue;

                    string name = p.Definition?.Name ?? "";
                    long pid = ParameterService.ParamIdToLong(p.Id);
                    string key = categoryName + "|" + name + "|" + (isType ? "T" : "I") + "|" + pid;

                    if (!groups.TryGetValue(key, out var g))
                    {
                        g = new ReplaceGroup
                        {
                            CategoryName = categoryName,
                            ParameterName = name,
                            IsType = isType,
                            KindText = ParameterKindHelper.Label(ParameterKindHelper.Determine(p)),
                            SampleValue = value
                        };
                        groups[key] = g;
                    }
                    g.Targets.Add(p);
                }
                catch
                {
                    // 個別パラメータの読み取り失敗は無視して続行
                }
            }
        }

        /// <summary>
        /// 選択されたまとまりの値を置き換える（反復子）。
        /// WriteBatchSize 件ごとに1トランザクションで確定し、処理済み件数を yield する。
        /// 全体は呼び出し側の TransactionGroup でまとめ、キャンセル時は巻き戻す。
        /// </summary>
        public static IEnumerable<int> Apply(
            Document doc, IEnumerable<ReplaceGroup> selectedGroups, string find, string replace, MatchMode mode,
            ReplaceResult result, CancellationToken ct)
        {
            var targets = selectedGroups.SelectMany(g => g.Targets).ToList();
            int done = 0;

            for (int start = 0; start < targets.Count; start += WriteBatchSize)
            {
                if (ct.IsCancellationRequested)
                {
                    result.Cancelled = true;
                    yield break;
                }

                int end = Math.Min(start + WriteBatchSize, targets.Count);
                int batchSuccess = 0, batchSkipped = 0;

                using (var t = new Transaction(doc, Loc.S("ValueReplace.Txn")))
                {
                    t.Start();
                    var opts = t.GetFailureHandlingOptions();
                    opts.SetFailuresPreprocessor(new WarningSwallower());
                    t.SetFailureHandlingOptions(opts);

                    for (int i = start; i < end; i++)
                    {
                        var p = targets[i];
                        try
                        {
                            // 走査後に値が変わっている場合に備えて、書き込む直前にもう一度確認する
                            string current = p.AsString();
                            if (!IsMatch(current, find, mode))
                            {
                                batchSkipped++;
                                continue;
                            }
                            if (p.Set(BuildNewValue(current, find, replace, mode)))
                                batchSuccess++;
                            else
                                result.Failed++;
                        }
                        catch (Exception ex)
                        {
                            result.Failed++;
                            DiagLog.Write($"[ValueReplace] 書き込み失敗 param='{SafeName(p)}': {ex.Message}");
                        }
                    }

                    var status = t.Commit();
                    if (status == TransactionStatus.Committed)
                    {
                        result.Success += batchSuccess;
                        result.Skipped += batchSkipped;
                    }
                    else
                    {
                        // このまとまりは確定できなかった（無視できないエラー等）→ 全件失敗扱い
                        result.Failed += batchSuccess;
                        result.Skipped += batchSkipped;
                        DiagLog.Write($"[ValueReplace] トランザクション確定失敗 status={status} 件数={end - start}");
                    }
                }

                done = end;
                yield return done;
            }
        }

        private static string SafeName(Parameter p)
        {
            try { return p?.Definition?.Name ?? ""; }
            catch { return ""; }
        }

        /// <summary>警告はコミットを止めないよう削除する（エラーは Revit の既定処理に任せる）</summary>
        private class WarningSwallower : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                foreach (var f in failuresAccessor.GetFailureMessages())
                {
                    if (f.GetSeverity() == FailureSeverity.Warning)
                        failuresAccessor.DeleteWarning(f);
                }
                return FailureProcessingResult.Continue;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.ParameterValueReplace.Services
{
    /// <summary>移動先にすでに値がある要素の扱い</summary>
    public enum ExistingTargetMode
    {
        Keep,       // 移動先の値を残し、その要素は移動しない（移動元もそのまま）
        Overwrite   // 移動元の値で上書きする
    }

    /// <summary>1つのパラメータ名について、グループごとの集計（読み込み時）</summary>
    public class GroupStats
    {
        /// <summary>そのグループのパラメータを持つ要素数</summary>
        public int ElementCount;
        /// <summary>そのグループのパラメータに値が入っている要素数</summary>
        public int ValueCount;
    }

    /// <summary>
    /// ダイアログの1行（パラメータ名ごと）。読み込み時の集計と、確認（プレビュー）の結果を持つ。
    /// </summary>
    public class MoveRow : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnChanged(string n) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));

        public string ParameterName { get; set; }

        /// <summary>グループ表示名 → 集計</summary>
        public Dictionary<string, GroupStats> Groups { get; } = new Dictionary<string, GroupStats>();

        private bool _isSelected = true;
        public bool IsSelected
        {
            get => _isSelected;
            set { if (_isSelected != value) { _isSelected = value; OnChanged(nameof(IsSelected)); } }
        }

        // --- 選択中のグループでの集計（表示用）---
        private int _sourceValueCount;
        public int SourceValueCount { get => _sourceValueCount; set { _sourceValueCount = value; OnChanged(nameof(SourceValueCount)); } }

        private int _targetElementCount;
        public int TargetElementCount { get => _targetElementCount; set { _targetElementCount = value; OnChanged(nameof(TargetElementCount)); } }

        // --- 確認（プレビュー）の結果 ---
        private string _movable = "", _noTarget = "", _targetHasValue = "", _conflict = "";
        public string MovableText { get => _movable; set { _movable = value; OnChanged(nameof(MovableText)); } }
        public string NoTargetText { get => _noTarget; set { _noTarget = value; OnChanged(nameof(NoTargetText)); } }
        public string TargetHasValueText { get => _targetHasValue; set { _targetHasValue = value; OnChanged(nameof(TargetHasValueText)); } }
        public string ConflictText { get => _conflict; set { _conflict = value; OnChanged(nameof(ConflictText)); } }

        public void ClearPreview()
        {
            MovableText = NoTargetText = TargetHasValueText = ConflictText = "";
        }
    }

    /// <summary>1要素・1パラメータ名ぶんの移動操作</summary>
    public class MoveOp
    {
        public Parameter Target;
        public string Value;
        public List<Parameter> Sources = new List<Parameter>();
    }

    /// <summary>確認（プレビュー）の集計</summary>
    public class MovePlan
    {
        public List<MoveOp> Ops = new List<MoveOp>();
        // パラメータ名 → 件数
        public Dictionary<string, int> Movable = new Dictionary<string, int>();
        public Dictionary<string, int> NoTarget = new Dictionary<string, int>();
        public Dictionary<string, int> TargetHasValue = new Dictionary<string, int>();
        public Dictionary<string, int> Conflict = new Dictionary<string, int>();

        public static void Add(Dictionary<string, int> d, string key)
        {
            d.TryGetValue(key, out int n);
            d[key] = n + 1;
        }
    }

    /// <summary>移動の結果</summary>
    public class MoveResult
    {
        public int Success;
        public int Failed;
        public int Skipped;
        public bool Cancelled;
    }

    /// <summary>
    /// 同じ名前でパラメータグループが違う文字パラメータの間で、値を要素ごとに移動するサービス。
    ///
    /// 【用途】GUID 違いの同名共有パラメータ（例: 「☆財産区分」がモデル プロパティとセットの2つ）に
    /// 値が分かれて入っているのを、一方（例: セット）へまとめる。Excel 上で列を移す作業を Revit 上で直接行う。
    ///
    /// 【判定】要素ごと・パラメータ名ごとに
    ///  - 移動元（指定グループ）に値が無い → 対象外
    ///  - 移動先（指定グループ）のパラメータを持たない → 「移動先なし」（移動しない。移動元もそのまま）
    ///  - 移動元に違う値が複数（同じグループに同名が複数）→ 「値の食い違い」（移動しない）
    ///  - 移動先にすでに違う値 → 「残す」なら移動しない / 「上書き」なら上書き
    ///  - それ以外 → 移動先へ書き込み、移動元を空欄にする
    /// 値を失わないことを優先し、判断できないものは移動しない。
    /// </summary>
    public static class ValueMoveService
    {
        private const int YieldEvery = 500;
        public const int WriteBatchSize = 2000;

        /// <summary>
        /// 範囲内の要素の文字パラメータを調べ、「同じ名前で2つ以上のグループにある」パラメータ名の一覧を作る（反復子）。
        /// </summary>
        public static IEnumerable<int> Load(
            Document doc, ReplaceScope scope, bool includeTypes, View activeView, ICollection<ElementId> selectionIds,
            Dictionary<string, MoveRow> rows, CancellationToken ct)
        {
            string otherLabel = OtherGroupLabel(doc);
            var groupCache = new Dictionary<long, string>();
            int processed = 0;
            foreach (var e in EnumerateTargets(doc, scope, includeTypes, activeView, selectionIds))
            {
                if (ct.IsCancellationRequested) yield break;

                foreach (var kv in GroupedParameters(e, otherLabel, groupCache))
                {
                    string name = kv.Key.Item1, group = kv.Key.Item2;
                    if (!rows.TryGetValue(name, out var row))
                    {
                        row = new MoveRow { ParameterName = name };
                        rows[name] = row;
                    }
                    if (!row.Groups.TryGetValue(group, out var st))
                    {
                        st = new GroupStats();
                        row.Groups[group] = st;
                    }
                    st.ElementCount++;
                    if (kv.Value.Any(p => !string.IsNullOrEmpty(SafeString(p))))
                        st.ValueCount++;
                }

                processed++;
                if (processed % YieldEvery == 0)
                    yield return processed;
            }

            // 1つのグループにしか無い名前は移動の対象にならないので除く
            foreach (var key in rows.Where(kv => kv.Value.Groups.Count < 2).Select(kv => kv.Key).ToList())
                rows.Remove(key);

            yield return processed;
        }

        /// <summary>
        /// 選んだパラメータ名について、実際の移動操作を組み立てる（確認＝プレビュー。まだ書き込まない）。
        /// </summary>
        public static IEnumerable<int> Plan(
            Document doc, ReplaceScope scope, bool includeTypes, View activeView, ICollection<ElementId> selectionIds,
            ICollection<string> names, string sourceGroup, string targetGroup, ExistingTargetMode mode,
            MovePlan plan, CancellationToken ct)
        {
            string otherLabel = OtherGroupLabel(doc);
            var groupCache = new Dictionary<long, string>();
            var nameSet = new HashSet<string>(names);
            int processed = 0;

            foreach (var e in EnumerateTargets(doc, scope, includeTypes, activeView, selectionIds))
            {
                if (ct.IsCancellationRequested) yield break;

                var grouped = GroupedParameters(e, otherLabel, groupCache, nameSet);
                foreach (string name in nameSet)
                {
                    grouped.TryGetValue(Tuple.Create(name, sourceGroup), out var sources);
                    if (sources == null) continue;

                    var values = sources.Select(SafeString).Where(v => !string.IsNullOrEmpty(v)).Distinct().ToList();
                    if (values.Count == 0) continue;   // 移動元が空 → 対象外

                    grouped.TryGetValue(Tuple.Create(name, targetGroup), out var targets);
                    var writableTargets = targets?.Where(p => !p.IsReadOnly).ToList();
                    if (writableTargets == null || writableTargets.Count != 1)
                    {
                        // 移動先を持たない（または同じグループに同名が複数あって決められない）
                        MovePlan.Add(plan.NoTarget, name);
                        continue;
                    }
                    if (values.Count > 1)
                    {
                        MovePlan.Add(plan.Conflict, name);
                        continue;
                    }

                    var target = writableTargets[0];
                    string value = values[0];
                    string current = SafeString(target);
                    if (!string.IsNullOrEmpty(current) && current != value)
                    {
                        MovePlan.Add(plan.TargetHasValue, name);
                        if (mode == ExistingTargetMode.Keep)
                            continue;   // 移動先の値を残し、この要素は移動しない
                    }

                    var op = new MoveOp { Target = target, Value = value };
                    op.Sources.AddRange(sources.Where(p => !p.IsReadOnly && !string.IsNullOrEmpty(SafeString(p))));
                    plan.Ops.Add(op);
                    MovePlan.Add(plan.Movable, name);
                }

                processed++;
                if (processed % YieldEvery == 0)
                    yield return processed;
            }
            yield return processed;
        }

        /// <summary>移動を実行する（反復子）。WriteBatchSize 件ごとに1トランザクション。</summary>
        public static IEnumerable<int> Apply(Document doc, List<MoveOp> ops, MoveResult result, CancellationToken ct)
        {
            for (int start = 0; start < ops.Count; start += WriteBatchSize)
            {
                if (ct.IsCancellationRequested)
                {
                    result.Cancelled = true;
                    yield break;
                }

                int end = Math.Min(start + WriteBatchSize, ops.Count);
                int batchSuccess = 0, batchSkipped = 0, batchFailed = 0;

                using (var t = new Transaction(doc, Loc.S("ValueMove.Txn")))
                {
                    t.Start();
                    var opts = t.GetFailureHandlingOptions();
                    opts.SetFailuresPreprocessor(new WarningSwallower());
                    t.SetFailureHandlingOptions(opts);

                    for (int i = start; i < end; i++)
                    {
                        var op = ops[i];
                        try
                        {
                            // 確認後に移動元が変わっていたら（別操作で消された等）書き込まない
                            if (!op.Sources.Any(p => SafeString(p) == op.Value) && SafeString(op.Target) != op.Value)
                            {
                                batchSkipped++;
                                continue;
                            }

                            if (SafeString(op.Target) != op.Value && !op.Target.Set(op.Value))
                            {
                                batchFailed++;
                                continue;
                            }
                            foreach (var s in op.Sources)
                                s.Set("");
                            batchSuccess++;
                        }
                        catch (Exception ex)
                        {
                            batchFailed++;
                            DiagLog.Write($"[ValueMove] 書き込み失敗: {ex.Message}");
                        }
                    }

                    if (t.Commit() == TransactionStatus.Committed)
                    {
                        result.Success += batchSuccess;
                        result.Skipped += batchSkipped;
                        result.Failed += batchFailed;
                    }
                    else
                    {
                        result.Failed += end - start - batchSkipped;
                        result.Skipped += batchSkipped;
                        DiagLog.Write($"[ValueMove] トランザクション確定失敗 件数={end - start}");
                    }
                }

                yield return end;
            }
        }

        // ===================== 共通 =====================

        /// <summary>範囲内の要素（インスタンス、必要ならそのタイプ）を重複なく列挙する</summary>
        private static IEnumerable<Element> EnumerateTargets(
            Document doc, ReplaceScope scope, bool includeTypes, View activeView, ICollection<ElementId> selectionIds)
        {
            FilteredElementCollector collector;
            switch (scope)
            {
                case ReplaceScope.ActiveView:
                    if (activeView == null) yield break;
                    collector = new FilteredElementCollector(doc, activeView.Id);
                    break;
                case ReplaceScope.Selection:
                    if (selectionIds == null || selectionIds.Count == 0) yield break;
                    collector = new FilteredElementCollector(doc, selectionIds);
                    break;
                default:
                    collector = new FilteredElementCollector(doc);
                    break;
            }

            var typeIds = new HashSet<ElementId>();
            foreach (var e in collector.WhereElementIsNotElementType())
            {
                if (!HasCategory(e)) continue;
                yield return e;

                if (includeTypes)
                {
                    ElementId tid = null;
                    try { tid = e.GetTypeId(); } catch { }
                    if (tid != null && tid != ElementId.InvalidElementId)
                        typeIds.Add(tid);
                }
            }

            if (!includeTypes) yield break;

            if (scope == ReplaceScope.EntireProject)
            {
                foreach (var t in new FilteredElementCollector(doc).WhereElementIsElementType())
                    if (HasCategory(t)) typeIds.Add(t.Id);
            }
            foreach (var tid in typeIds)
            {
                Element t = null;
                try { t = doc.GetElement(tid); } catch { }
                if (t != null) yield return t;
            }
        }

        /// <summary>要素の文字パラメータを（名前, グループ表示名）ごとにまとめる。names 指定時はその名前だけ。</summary>
        private static Dictionary<Tuple<string, string>, List<Parameter>> GroupedParameters(
            Element e, string otherLabel, Dictionary<long, string> groupCache, HashSet<string> names = null)
        {
            var result = new Dictionary<Tuple<string, string>, List<Parameter>>();
            ParameterSet ps;
            try { ps = e.Parameters; }
            catch { return result; }

            foreach (Parameter p in ps)
            {
                try
                {
                    if (p == null || p.StorageType != StorageType.String) continue;
                    var def = p.Definition;
                    string name = def?.Name;
                    if (string.IsNullOrEmpty(name)) continue;
                    if (names != null && !names.Contains(name)) continue;

                    // グループ表示名はパラメータ（識別番号）ごとに決まるので、1回だけ求めて使い回す
                    long pid = ExcelExportImport.Services.ParameterService.ParamIdToLong(p.Id);
                    if (!groupCache.TryGetValue(pid, out string group))
                    {
                        group = ParameterGroupHelper.GetGroupLabel(def, otherLabel);
                        groupCache[pid] = group;
                    }
                    var key = Tuple.Create(name, group);
                    if (!result.TryGetValue(key, out var list))
                    {
                        list = new List<Parameter>();
                        result[key] = list;
                    }
                    list.Add(p);
                }
                catch
                {
                    // 個別パラメータの読み取り失敗は無視
                }
            }
            return result;
        }

        private static string OtherGroupLabel(Document doc)
            => Loc.S("Export.ParamGroup.Other", RevitUiLanguage.Resolve(doc));

        private static string SafeString(Parameter p)
        {
            try { return p != null && p.HasValue ? (p.AsString() ?? "") : ""; }
            catch { return ""; }
        }

        private static bool HasCategory(Element e)
        {
            try { return e?.Category != null; }
            catch { return false; }
        }

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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Autodesk.Revit.DB;
using Tools28.Commands.ParameterCleanup.Models;
using Tools28.Localization;

namespace Tools28.Commands.ParameterCleanup.Services
{
    /// <summary>
    /// プロジェクト内の削除可能なパラメータ（プロジェクト／共有／グローバル）を列挙し、
    /// 各パラメータが「どこで使われているか」「値が入っている要素があるか」を判定するサービス。
    ///
    /// 【v2 の重要な変更】
    /// 値の判定を ParameterBindings（プロジェクトパラメータのバインド）に依存させない。
    /// バインドされていない共有パラメータ（ファミリ内部で定義され、ファミリ読込により
    /// プロジェクトに登録された SharedParameterElement）でも、要素側から実際に値を読み取る。
    /// これにより「バインドなし」と表示されて値の有無が分からないまま削除する事故を防ぐ。
    ///
    /// 大容量モデルでの UI フリーズ回避方針:
    ///   - 列挙 (EnumerateParameters) は ParameterBindings とパラメータ要素の走査のみで軽量。
    ///   - 使用箇所索引 (ParameterUsageIndex) をドキュメント1回走査で構築し、全パラメータで共有。
    ///   - 値の有無判定 (ScanRow) は反復子として実装し、一定件数ごとに yield して
    ///     呼び出し側（UI スレッド）がメッセージポンプへ制御を返せるようにする。
    ///   - 値が1件でも見つかれば即 break（early-exit）。
    /// </summary>
    public class ParameterScanner
    {
        /// <summary>この件数ごとに反復子が yield して UI へ制御を返す</summary>
        private const int YieldEvery = 2000;

        /// <summary>使用箇所テキストに載せるグループ数の上限</summary>
        private const int MaxUsageLabels = 12;

        private ParameterUsageIndex _index;

        /// <summary>構築済みの使用箇所索引（未構築なら null）</summary>
        public ParameterUsageIndex Index => _index;

        /// <summary>
        /// 削除可能な全パラメータを列挙する（軽量・同期）。
        /// </summary>
        public List<ParamRow> EnumerateParameters(Document doc)
        {
            var rows = new List<ParamRow>();

            // --- バインド情報を先に構築 ---
            // ⚠ 名前ではなく InternalDefinition.Id（= ParameterElement の Id）でキー化する。
            //    名前キーだと同名パラメータが複数ある場合に取り違え、
            //    「バインド済みなのにバインドなし」「その逆」の誤判定が起きる。
            //    Id が取れない環境向けに名前キーもフォールバックとして併用する。
            var bindingById = new Dictionary<int, BindingInfo>();
            var bindingByName = new Dictionary<string, BindingInfo>();
            try
            {
                var it = doc.ParameterBindings.ForwardIterator();
                it.Reset();
                while (it.MoveNext())
                {
                    Definition def = it.Key;
                    if (def == null || string.IsNullOrEmpty(def.Name)) continue;

                    var info = new BindingInfo();
                    var binding = it.Current as Binding;
                    if (binding is InstanceBinding ib)
                    {
                        info.IsTypeBinding = false;
                        CollectCategories(ib.Categories, info.Categories);
                    }
                    else if (binding is TypeBinding tb)
                    {
                        info.IsTypeBinding = true;
                        CollectCategories(tb.Categories, info.Categories);
                    }

                    bindingByName[def.Name] = info;

                    try
                    {
                        var idef = def as InternalDefinition;
                        if (idef != null && idef.Id != null && idef.Id != ElementId.InvalidElementId)
                            bindingById[idef.Id.IntValue()] = info;
                    }
                    catch { }
                }
            }
            catch { }

            // --- 集計表（スケジュール）での参照を先に構築（軽量：集計表を舐めるだけ）---
            var scheduleRefs = BuildScheduleReferences(doc);

            // --- プロジェクト／共有パラメータ（ParameterElement / SharedParameterElement）---
            var paramElems = new Dictionary<ElementId, ParameterElement>();
            AddParameterElements(doc, typeof(ParameterElement), paramElems);
            AddParameterElements(doc, typeof(SharedParameterElement), paramElems);

            foreach (var pe in paramElems.Values)
            {
                Definition def;
                try { def = pe.GetDefinition(); }
                catch { def = null; }
                if (def == null || string.IsNullOrEmpty(def.Name)) continue;

                var row = new ParamRow
                {
                    Name = def.Name,
                    Id = pe.Id,
                    Definition = def,
                    Kind = (pe is SharedParameterElement) ? ParamKind.Shared : ParamKind.Project
                };

                var spe = pe as SharedParameterElement;
                if (spe != null)
                {
                    try { row.SharedGuid = spe.GuidValue; }
                    catch { row.SharedGuid = Guid.Empty; }
                }

                BindingInfo bi;
                if (!bindingById.TryGetValue(pe.Id.IntValue(), out bi))
                    bindingByName.TryGetValue(def.Name, out bi);

                if (bi != null && bi.Categories.Count > 0)
                {
                    row.IsTypeBinding = bi.IsTypeBinding;
                    row.BoundCategories = bi.Categories;
                    var names = bi.Categories.Select(c => c.Name).ToList();
                    names.Sort(StringComparer.CurrentCulture);
                    row.CategoriesText = string.Join(", ", names);
                }
                else
                {
                    // プロジェクトのカテゴリにはバインドされていない。
                    // ただしファミリ内部で定義された共有パラメータとして要素上に実在する場合があるため、
                    // 「判定対象外」にはせず、使用箇所索引を使って必ず値を確認する。
                    row.IsTypeBinding = null;
                    row.State = ValueState.Unchecked;
                }

                List<string> sref;
                if (scheduleRefs.TryGetValue(row.Id, out sref))
                    row.ScheduleRefText = string.Join(", ", sref);

                rows.Add(row);
            }

            // --- グローバルパラメータ ---
            try
            {
                if (GlobalParametersManager.AreGlobalParametersAllowed(doc))
                {
                    foreach (var gid in GlobalParametersManager.GetAllGlobalParameters(doc))
                    {
                        var gp = doc.GetElement(gid) as GlobalParameter;
                        if (gp == null) continue;
                        rows.Add(new ParamRow
                        {
                            Name = gp.Name,
                            Id = gp.Id,
                            Kind = ParamKind.Global,
                            IsTypeBinding = null,
                            State = ValueState.NotApplicable,
                            GlobalValueText = FormatGlobalValue(gp)
                        });
                    }
                }
            }
            catch { }

            // 同名フラグ付与
            foreach (var grp in rows.GroupBy(r => r.Name))
            {
                if (grp.Count() > 1)
                    foreach (var r in grp) r.IsDuplicateName = true;
            }

            return rows
                .OrderByDescending(r => r.IsDuplicateName)
                .ThenBy(r => r.Name)
                .ToList();
        }

        /// <summary>
        /// 使用箇所索引を構築する反復子（全パラメータで共有する1回だけの走査）。
        /// </summary>
        public IEnumerable<int> BuildIndex(Document doc, IEnumerable<ParamRow> rows, CancellationToken ct)
        {
            var ids = new HashSet<int>();
            var guids = new Dictionary<Guid, int>();
            foreach (var r in rows)
            {
                if (r.Kind == ParamKind.Global) continue;
                if (r.Id == null || r.Id == ElementId.InvalidElementId) continue;
                int pid = r.Id.IntValue();
                if (pid <= 0) continue;
                ids.Add(pid);
                if (r.SharedGuid != Guid.Empty && !guids.ContainsKey(r.SharedGuid))
                    guids[r.SharedGuid] = pid;
            }

            _index = new ParameterUsageIndex(doc);
            foreach (var n in _index.Build(ids, guids, ct))
                yield return n;
        }

        private static void AddParameterElements(Document doc, Type t, Dictionary<ElementId, ParameterElement> map)
        {
            try
            {
                foreach (var e in new FilteredElementCollector(doc).OfClass(t))
                {
                    if (e is GlobalParameter) continue;       // グローバルは別途処理
                    if (e is ParameterElement pe && !map.ContainsKey(pe.Id))
                        map[pe.Id] = pe;
                }
            }
            catch { }
        }

        /// <summary>バインドマップの1エントリ分のバインド情報</summary>
        private class BindingInfo
        {
            public bool? IsTypeBinding;
            public List<Category> Categories = new List<Category>();
        }

        private static void CollectCategories(CategorySet cats, List<Category> bucket)
        {
            if (cats == null) return;
            foreach (Category c in cats)
            {
                if (c != null) bucket.Add(c);
            }
        }

        /// <summary>
        /// 集計表（ViewSchedule）のフィールドが参照するパラメータ要素Id -> 集計表名リスト を構築する。
        /// 全要素走査ではなく集計表とそのフィールドを舐めるだけなので軽量。
        /// 組み込みパラメータ（負のId）は除外し、ユーザー作成パラメータのみ対象とする。
        /// </summary>
        private static Dictionary<ElementId, List<string>> BuildScheduleReferences(Document doc)
        {
            var map = new Dictionary<ElementId, List<string>>();
            try
            {
                var schedules = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSchedule))
                    .Cast<ViewSchedule>();

                foreach (var vs in schedules)
                {
                    if (vs == null || vs.IsTemplate) continue;

                    ScheduleDefinition sd;
                    try { sd = vs.Definition; } catch { continue; }
                    if (sd == null) continue;

                    int count;
                    try { count = sd.GetFieldCount(); } catch { continue; }

                    for (int i = 0; i < count; i++)
                    {
                        ScheduleField f;
                        try { f = sd.GetField(i); } catch { continue; }
                        if (f == null) continue;

                        ElementId pid;
                        try { pid = f.ParameterId; } catch { continue; }
                        if (pid == null || pid == ElementId.InvalidElementId) continue;
                        if (pid.IntValue() <= 0) continue;   // 組み込みパラメータを除外

                        List<string> list;
                        if (!map.TryGetValue(pid, out list))
                        {
                            list = new List<string>();
                            map[pid] = list;
                        }
                        if (!list.Contains(vs.Name)) list.Add(vs.Name);
                    }
                }
            }
            catch { }
            return map;
        }

        /// <summary>
        /// 1パラメータ分の使用箇所と値有無を判定する反復子。
        /// 一定件数ごとに「処理済み要素数」を yield する。列挙完了時に row.State を確定させる。
        /// キャンセル時は State を変更せず yield break する（呼び出し側でリセット）。
        /// </summary>
        public IEnumerable<int> ScanRow(Document doc, ParamRow row, CancellationToken ct)
        {
            if (row == null) yield break;

            if (row.Kind == ParamKind.Global)
            {
                row.State = ValueState.NotApplicable;
                yield break;
            }

            if (_index == null || !_index.IsBuilt)
            {
                // 索引未構築なら判定不能（呼び出し側で BuildIndex を先に回すこと）
                if (row.State == ValueState.Checking) row.State = ValueState.Unchecked;
                yield break;
            }

            var groups = _index.GetGroups(row.Id);

            if (groups.Count == 0)
            {
                // どの要素にも存在しない = 定義だけが残っている未使用パラメータ
                row.UsageText = "";
                row.UsageElementCount = 0;
                row.SampleText = "";
                row.State = ValueState.NotFound;
                yield return 0;
                yield break;
            }

            row.UsageText = BuildUsageText(groups);
            row.UsageElementCount = groups.Sum(g => g.Elements.Count);

            int processed = 0;
            bool found = false;
            var access = new ParameterAccessor(row);

            foreach (var g in groups)
            {
                if (found) break;

                foreach (var eid in g.Elements)
                {
                    if (ct.IsCancellationRequested) yield break;

                    Element e = null;
                    try { e = doc.GetElement(eid); }
                    catch { e = null; }

                    if (e != null)
                    {
                        var p = access.Get(e);
                        if (HasRealValue(p))
                        {
                            found = true;
                            row.SampleText = BuildSampleText(g, e, p);
                            break;
                        }
                    }

                    processed++;
                    if (processed % YieldEvery == 0)
                        yield return processed;
                }
            }

            if (!found) row.SampleText = "";
            row.State = found ? ValueState.HasValue : ValueState.Empty;
            yield return processed;
        }

        /// <summary>
        /// 要素からパラメータを取り出す方法を1度だけ決定してキャッシュするヘルパー。
        /// 共有パラメータは GUID、非共有は Definition で引ける。
        /// どちらも効かない場合のみ Parameters 列挙（低速）にフォールバックする。
        /// </summary>
        private class ParameterAccessor
        {
            private enum Mode { Undetermined, Guid, Definition, Enumerate }

            private readonly ParamRow _row;
            private readonly int _paramId;
            private Mode _mode = Mode.Undetermined;

            public ParameterAccessor(ParamRow row)
            {
                _row = row;
                _paramId = (row.Id != null) ? row.Id.IntValue() : 0;
            }

            public Parameter Get(Element e)
            {
                Parameter p;

                // 有効と分かっている方法を先に試す（取得できなければ他の方法も試す）
                switch (_mode)
                {
                    case Mode.Guid:
                        p = ByGuid(e); if (p != null) return p; break;
                    case Mode.Definition:
                        p = ByDefinition(e); if (p != null) return p; break;
                    case Mode.Enumerate:
                        return ByEnumerate(e);
                }

                p = ByGuid(e);
                if (p != null) { _mode = Mode.Guid; return p; }

                p = ByDefinition(e);
                if (p != null) { _mode = Mode.Definition; return p; }

                p = ByEnumerate(e);
                if (p != null) { _mode = Mode.Enumerate; return p; }

                return null;   // この要素に無いだけの可能性があるので Mode は確定させない
            }

            private Parameter ByGuid(Element e)
            {
                if (_row.SharedGuid == Guid.Empty) return null;
                try { return e.get_Parameter(_row.SharedGuid); }
                catch { return null; }
            }

            private Parameter ByDefinition(Element e)
            {
                if (_row.Definition == null) return null;
                try { return e.get_Parameter(_row.Definition); }
                catch { return null; }
            }

            private Parameter ByEnumerate(Element e)
            {
                if (_paramId <= 0) return null;
                try
                {
                    foreach (Parameter q in e.Parameters)
                    {
                        if (q == null || q.Id == null) continue;
                        if (q.Id.IntValue() == _paramId) return q;
                    }
                }
                catch { }
                return null;
            }
        }

        private static string BuildUsageText(List<ParameterUsageIndex.UsageGroup> groups)
        {
            // 「カテゴリ名: 件数」を件数の多い順に並べる
            var byCat = new Dictionary<string, int>();
            foreach (var g in groups)
            {
                string cat = string.IsNullOrEmpty(g.CategoryName) ? "-" : g.CategoryName;
                if (g.IsElementType) cat += Loc.S("ParamCleanup.Usage.TypeSuffix");
                int cur;
                byCat.TryGetValue(cat, out cur);
                byCat[cat] = cur + g.Elements.Count;
            }

            var parts = byCat.OrderByDescending(kv => kv.Value)
                             .Take(MaxUsageLabels)
                             .Select(kv => string.Format("{0} ({1})", kv.Key, kv.Value))
                             .ToList();
            if (byCat.Count > MaxUsageLabels) parts.Add("…");
            return string.Join(", ", parts);
        }

        private static string BuildSampleText(ParameterUsageIndex.UsageGroup g, Element e, Parameter p)
        {
            string val = "";
            try
            {
                val = p.AsValueString();
                if (string.IsNullOrEmpty(val))
                {
                    switch (p.StorageType)
                    {
                        case StorageType.String: val = p.AsString() ?? ""; break;
                        case StorageType.Integer: val = p.AsInteger().ToString(); break;
                        case StorageType.Double: val = p.AsDouble().ToString("0.###"); break;
                        case StorageType.ElementId:
                            var id = p.AsElementId();
                            var t = (id != null) ? e.Document.GetElement(id) : null;
                            val = (t != null) ? (t.Name ?? id.IntValue().ToString()) : (id != null ? id.IntValue().ToString() : "");
                            break;
                    }
                }
            }
            catch { }

            string label = !string.IsNullOrEmpty(g.Label) ? g.Label : (g.CategoryName ?? "");
            string eid = "";
            try { eid = e.Id.IntValue().ToString(); }
            catch { }

            return string.Format(Loc.S("ParamCleanup.Usage.Sample"), label, eid, val);
        }

        /// <summary>
        /// 値が「実質的に入っている」か判定。
        /// 文字列は空白のみを除外、ElementId は無効IDを除外。
        /// 数値・整数（Yes/No 等）は常に値を持つため安全側で「値あり」とみなす。
        /// </summary>
        private static bool HasRealValue(Parameter p)
        {
            if (p == null || !p.HasValue) return false;

            switch (p.StorageType)
            {
                case StorageType.String:
                    return !string.IsNullOrWhiteSpace(p.AsString());
                case StorageType.ElementId:
                    var id = p.AsElementId();
                    return id != null && id != ElementId.InvalidElementId;
                case StorageType.Integer:
                case StorageType.Double:
                    return true;
                default:
                    return false;
            }
        }

        private static string FormatGlobalValue(GlobalParameter gp)
        {
            try
            {
                var v = gp.GetValue();
                if (v is StringParameterValue sv) return sv.Value ?? "";
                if (v is DoubleParameterValue dv) return dv.Value.ToString();
                if (v is IntegerParameterValue iv) return iv.Value.ToString();
                if (v is ElementIdParameterValue ev) return ev.Value?.ToString() ?? "";
            }
            catch { }
            return "";
        }

        public void ClearCache()
        {
            if (_index != null) _index.Clear();
            _index = null;
        }
    }
}

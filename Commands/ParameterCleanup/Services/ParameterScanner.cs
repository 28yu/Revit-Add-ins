using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Autodesk.Revit.DB;
using Tools28.Commands.ParameterCleanup.Models;

namespace Tools28.Commands.ParameterCleanup.Services
{
    /// <summary>
    /// プロジェクト内の削除可能なパラメータ（プロジェクト／共有／グローバル）を列挙し、
    /// 各パラメータに値が入っている要素が1つでもあるかを判定するサービス。
    ///
    /// 大容量モデルでの UI フリーズ回避方針:
    ///   - 列挙 (EnumerateParameters) は ParameterBindings とパラメータ要素の走査のみで軽量。
    ///   - 値の有無判定 (ScanRow) は反復子として実装し、一定件数ごとに yield して
    ///     呼び出し側（UI スレッド）がメッセージポンプへ制御を返せるようにする。
    ///   - 値が1件でも見つかれば即 break（early-exit）。
    ///   - カテゴリ単位で要素リストをキャッシュし、同カテゴリの複数パラメータで再利用。
    /// </summary>
    public class ParameterScanner
    {
        /// <summary>この件数ごとに反復子が yield して UI へ制御を返す</summary>
        private const int YieldEvery = 2000;

        // (カテゴリId + I/T) -> 要素リスト のキャッシュ
        private readonly Dictionary<string, List<Element>> _elementCache
            = new Dictionary<string, List<Element>>();

        /// <summary>
        /// 削除可能な全パラメータを列挙する（軽量・同期）。
        /// </summary>
        public List<ParamRow> EnumerateParameters(Document doc)
        {
            var rows = new List<ParamRow>();

            // --- バインド情報を先に構築 ---
            // get_Item(Definition) は解決に失敗する場合があるため、実績のある
            // ForwardIterator でバインドマップを一度走査し、パラメータ名でキー化する。
            // （プロジェクトにバインド済みのパラメータ名は実質一意なので名前キーで安全）
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

                if (bindingByName.TryGetValue(def.Name, out var bi) && bi.Categories.Count > 0)
                {
                    row.IsTypeBinding = bi.IsTypeBinding;
                    row.BoundCategories = bi.Categories;
                    var names = bi.Categories.Select(c => c.Name).ToList();
                    names.Sort(StringComparer.CurrentCulture);
                    row.CategoriesText = string.Join(", ", names);
                    // State は Unchecked のまま（スキャン対象）
                }
                else
                {
                    // どのカテゴリにもバインドされていない（ファミリ用に読込済みの共有パラメータ等）
                    row.IsTypeBinding = null;
                    row.State = ValueState.NotApplicable;
                }

                if (scheduleRefs.TryGetValue(row.Id, out var sref))
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

                        if (!map.TryGetValue(pid, out var list))
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
        /// 1パラメータ分の値有無を判定する反復子。
        /// 一定件数ごとに「処理済み要素数」を yield する。列挙完了時に row.State を確定させる。
        /// キャンセル時は State を変更せず yield break する（呼び出し側でリセット）。
        /// </summary>
        public IEnumerable<int> ScanRow(Document doc, ParamRow row, CancellationToken ct)
        {
            if (!row.IsScannable)
            {
                if (row.State == ValueState.Checking || row.State == ValueState.Unchecked)
                    row.State = ValueState.NotApplicable;
                yield break;
            }

            bool isType = row.IsTypeBinding == true;
            bool found = false;
            int processed = 0;

            foreach (var cat in row.BoundCategories)
            {
                var elements = GetElements(doc, cat, isType);
                foreach (var e in elements)
                {
                    if (ct.IsCancellationRequested) yield break;

                    Parameter p = null;
                    try { p = e.get_Parameter(row.Definition); }
                    catch { p = null; }

                    if (HasRealValue(p)) { found = true; break; }

                    processed++;
                    if (processed % YieldEvery == 0)
                        yield return processed;
                }
                if (found) break;
            }

            row.State = found ? ValueState.HasValue : ValueState.Empty;
            yield return processed;
        }

        private List<Element> GetElements(Document doc, Category cat, bool isType)
        {
            string key = cat.Id.ToString() + (isType ? "|T" : "|I");
            if (_elementCache.TryGetValue(key, out var cached))
                return cached;

            List<Element> list;
            try
            {
                var col = new FilteredElementCollector(doc).OfCategoryId(cat.Id);
                col = isType ? col.WhereElementIsElementType() : col.WhereElementIsNotElementType();
                list = col.ToList();
            }
            catch
            {
                list = new List<Element>();
            }

            _elementCache[key] = list;
            return list;
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

        public void ClearCache() => _elementCache.Clear();
    }
}

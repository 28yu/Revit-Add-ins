using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.DuplicateInstance.Models;

namespace Tools28.Commands.DuplicateInstance.Services
{
    /// <summary>
    /// モデルの警告から「同じ場所に同一のインスタンスがあります」を探し、
    /// 重なっている要素を組ごとにまとめて一覧用の行を作る。
    /// </summary>
    public class DuplicateScanner
    {
        public List<DuplicateRow> Scan(Document doc)
        {
            var dupGuid = BuiltInFailures.OverlapFailures.DuplicateInstances.Guid;

            // 1つの警告に含まれる要素の集まりを「組」とする。
            // 3つ以上重なっていると警告が複数に分かれることがあるため、共通の要素を持つ組はまとめる。
            var uf = new UnionFind();
            foreach (var w in doc.GetWarnings())
            {
                if (w == null || w.GetFailureDefinitionId().Guid != dupGuid) continue;
                var ids = new List<ElementId>();
                ids.AddRange(w.GetFailingElements() ?? new List<ElementId>());
                ids.AddRange(w.GetAdditionalElements() ?? new List<ElementId>());
                ids = ids.Where(id => id != null && id != ElementId.InvalidElementId && doc.GetElement(id) != null)
                         .Distinct().ToList();
                if (ids.Count < 2) continue;
                for (int i = 1; i < ids.Count; i++) uf.Union(ids[0], ids[i]);
            }

            var groups = uf.Groups()
                           .Where(g => g.Count >= 2)
                           .Select(g => g.OrderBy(id => id.LongValue()).ToList())
                           .OrderBy(g => g[0].LongValue())
                           .ToList();

            var rows = new List<DuplicateRow>();
            int no = 0;
            foreach (var g in groups)
            {
                no++;
                var groupRows = g.Select(id => BuildRow(doc, doc.GetElement(id), no, g.Count)).ToList();
                ComputeDiffs(groupRows);
                rows.AddRange(groupRows);
            }
            return rows;
        }

        private DuplicateRow BuildRow(Document doc, Element e, int groupNo, int groupSize)
        {
            var row = new DuplicateRow
            {
                GroupNo = groupNo,
                GroupSize = groupSize,
                Id = e.Id,
                IdValue = e.Id.LongValue(),
                Category = e.Category?.Name ?? "",
            };

            try
            {
                if (doc.GetElement(e.GetTypeId()) is ElementType et)
                {
                    row.Family = et.FamilyName ?? "";
                    row.TypeName = et.Name ?? "";
                }
            }
            catch { }

            row.Level = SafeName(doc, e.LevelId);
            if (string.IsNullOrEmpty(row.Level))
            {
                // 柱・梁などは LevelId を持たず、参照レベルのパラメータで持つ
                foreach (var bip in new[] { BuiltInParameter.FAMILY_LEVEL_PARAM,
                                            BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                                            BuiltInParameter.SCHEDULE_LEVEL_PARAM })
                {
                    try
                    {
                        var p = e.get_Parameter(bip);
                        if (p != null && p.StorageType == StorageType.ElementId)
                        {
                            row.Level = SafeName(doc, p.AsElementId());
                            if (!string.IsNullOrEmpty(row.Level)) break;
                        }
                    }
                    catch { }
                }
            }

            try
            {
                if (doc.IsWorkshared && e.WorksetId != WorksetId.InvalidWorksetId)
                    row.Workset = doc.GetWorksetTable().GetWorkset(e.WorksetId)?.Name ?? "";
            }
            catch { }

            row.Mark = ParamText(doc, e.get_Parameter(BuiltInParameter.ALL_MODEL_MARK));
            row.Comments = ParamText(doc, e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS));
            try { row.Host = (e as FamilyInstance)?.Host?.Name ?? ""; } catch { }

            // インスタンスパラメータを全部読む（同じ名前が2つあるときは番号を付けて区別する）
            int filled = 0;
            foreach (Parameter p in e.Parameters)
            {
                if (p?.Definition == null) continue;
                string name = p.Definition.Name ?? "";
                string key = name;
                for (int n = 2; row.Parameters.ContainsKey(key); n++) key = $"{name} [{n}]";

                string val = ParamText(doc, p);
                row.Parameters[key] = val;
                if (!string.IsNullOrWhiteSpace(val)) filled++;
            }
            row.FilledCount = filled;
            return row;
        }

        /// <summary>同じ組の中で値が違うパラメータの数を、行ごとに数える。</summary>
        private static void ComputeDiffs(List<DuplicateRow> group)
        {
            var keys = group.SelectMany(r => r.Parameters.Keys).Distinct().ToList();
            foreach (var r in group)
            {
                int diff = 0;
                foreach (var k in keys)
                {
                    string mine = r.Parameters.TryGetValue(k, out var v) ? v : null;
                    if (group.Any(o => o != r && (o.Parameters.TryGetValue(k, out var ov) ? ov : null) != mine))
                        diff++;
                }
                r.DiffCount = diff;
            }
        }

        /// <summary>パラメータの値を画面に出す文字にする（値が無ければ空文字）。</summary>
        public static string ParamText(Document doc, Parameter p)
        {
            if (p == null) return "";
            try
            {
                if (!p.HasValue) return "";
                switch (p.StorageType)
                {
                    case StorageType.String:
                        return p.AsString() ?? "";
                    case StorageType.ElementId:
                        var id = p.AsElementId();
                        if (id == null || id == ElementId.InvalidElementId) return "";
                        var s = p.AsValueString();
                        if (!string.IsNullOrEmpty(s)) return s;
                        return SafeName(doc, id);
                    case StorageType.Integer:
                        return p.AsValueString() ?? p.AsInteger().ToString();
                    case StorageType.Double:
                        return p.AsValueString() ?? p.AsDouble().ToString("0.###");
                    default:
                        return "";
                }
            }
            catch
            {
                return "";
            }
        }

        private static string SafeName(Document doc, ElementId id)
        {
            try
            {
                if (id == null || id == ElementId.InvalidElementId) return "";
                return doc.GetElement(id)?.Name ?? "";
            }
            catch { return ""; }
        }

        /// <summary>共通の要素を持つ組をまとめるための簡易な Union-Find。</summary>
        private class UnionFind
        {
            private readonly Dictionary<long, long> _parent = new Dictionary<long, long>();
            private readonly Dictionary<long, ElementId> _ids = new Dictionary<long, ElementId>();

            private long Find(long x)
            {
                if (!_parent.ContainsKey(x)) _parent[x] = x;
                while (_parent[x] != x)
                {
                    _parent[x] = _parent[_parent[x]];
                    x = _parent[x];
                }
                return x;
            }

            public void Union(ElementId a, ElementId b)
            {
                long av = a.LongValue(), bv = b.LongValue();
                _ids[av] = a; _ids[bv] = b;
                long ra = Find(av), rb = Find(bv);
                if (ra != rb) _parent[ra] = rb;
            }

            public List<List<ElementId>> Groups()
            {
                return _ids.Keys.GroupBy(Find)
                           .Select(g => g.Select(k => _ids[k]).ToList())
                           .ToList();
            }
        }
    }
}

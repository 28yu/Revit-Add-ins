using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Autodesk.Revit.DB;

namespace Tools28.Commands.ParameterCleanup.Services
{
    /// <summary>
    /// 「どの要素がどのパラメータを実際に保持しているか」をドキュメント1回走査で索引化するサービス。
    ///
    /// 【なぜ必要か】
    /// ParameterBindings（＝プロジェクトパラメータのバインド）に載っていないパラメータでも、
    /// ファミリ内部で定義された共有パラメータは、そのファミリの要素／タイプ上に実在し値を持つ。
    /// バインド情報からは走査対象カテゴリが分からないため、従来は「バインドなし＝判定不能」として
    /// 値の確認をスキップしていた（＝値が入っているのに削除できてしまう事故の原因）。
    /// 本クラスはバインドに依存せず、要素側から実際のパラメータ保持状況を索引化する。
    ///
    /// 【走査量を抑える工夫】
    ///  1) インスタンスは「タイプIdごと」にグループ化する（同一タイプの要素はパラメータ構成が同じ）。
    ///     グループ化のパスでは GetTypeId()/Category を読むだけで、パラメータ列挙は行わない。
    ///  2) 要素タイプは「ファミリごと」にグループ化する（同一ファミリのタイプは構成が同じ）。
    ///  3) パラメータ列挙は各グループの代表要素1つに対してのみ行う（数千回で済む）。
    ///  4) 値の判定は、そのパラメータを持つグループの要素だけを走査する（全要素走査を回避）。
    /// </summary>
    public class ParameterUsageIndex
    {
        /// <summary>この件数ごとに反復子が yield して UI へ制御を返す</summary>
        private const int YieldEvery = 3000;

        /// <summary>パラメータ構成を共有する要素の集合</summary>
        public class UsageGroup
        {
            /// <summary>要素タイプ（ElementType）のグループなら true、インスタンスなら false</summary>
            public bool IsElementType;

            /// <summary>インスタンスグループの場合のタイプ Id</summary>
            public ElementId TypeId = ElementId.InvalidElementId;

            public ElementId CategoryId = ElementId.InvalidElementId;
            public string CategoryName = "";

            /// <summary>表示用ラベル（ファミリ名: タイプ名 等）</summary>
            public string Label = "";

            /// <summary>このグループに属する要素 Id</summary>
            public List<ElementId> Elements = new List<ElementId>();
        }

        private readonly Document _doc;
        private readonly List<UsageGroup> _groups = new List<UsageGroup>();

        // パラメータ要素Id(int) -> そのパラメータを保持するグループ
        private readonly Dictionary<int, List<UsageGroup>> _byParam = new Dictionary<int, List<UsageGroup>>();

        public bool IsBuilt { get; private set; }

        /// <summary>索引化した要素数（進捗表示用）</summary>
        public int ScannedElementCount { get; private set; }

        public ParameterUsageIndex(Document doc)
        {
            _doc = doc;
        }

        /// <summary>
        /// 索引を構築する反復子。一定件数ごとに処理済み要素数を yield する。
        /// キャンセル時は IsBuilt を立てずに yield break する。
        /// </summary>
        /// <param name="targetParamIds">索引に載せるパラメータ要素 Id（絞り込みにより列挙コストを削減）</param>
        /// <param name="targetGuids">共有パラメータ GUID -> パラメータ要素 Id（Id 照合の保険）</param>
        public IEnumerable<int> Build(HashSet<int> targetParamIds,
                                      Dictionary<Guid, int> targetGuids,
                                      CancellationToken ct)
        {
            _groups.Clear();
            _byParam.Clear();
            IsBuilt = false;
            ScannedElementCount = 0;

            int n = 0;

            // ---- フェーズ1: インスタンス要素をタイプ Id でグループ化（パラメータ列挙はしない）----
            var map = new Dictionary<string, UsageGroup>();
            var seen = new HashSet<int>();

            IList<Element> instances;
            try
            {
                instances = new FilteredElementCollector(_doc)
                    .WhereElementIsNotElementType()
                    .ToElements();
            }
            catch { instances = new List<Element>(); }

            foreach (var e in instances)
            {
                if (e == null) continue;
                if (ct.IsCancellationRequested) yield break;

                AddInstance(e, map, seen);

                n++;
                if (n % YieldEvery == 0) yield return n;
            }

            // ProjectInformation はコレクタから漏れる環境があるため明示的に追加する
            // （プロジェクトパラメータの主要なバインド先のため取りこぼし厳禁）
            try
            {
                var pi = _doc.ProjectInformation;
                if (pi != null) AddInstance(pi, map, seen);
            }
            catch { }

            // ---- フェーズ2: 要素タイプをファミリ単位でグループ化 ----
            IList<Element> types;
            try
            {
                types = new FilteredElementCollector(_doc)
                    .WhereElementIsElementType()
                    .ToElements();
            }
            catch { types = new List<Element>(); }

            foreach (var t in types)
            {
                if (t == null) continue;
                if (ct.IsCancellationRequested) yield break;

                AddType(t, map, seen);

                n++;
                if (n % YieldEvery == 0) yield return n;
            }

            // ---- フェーズ3: 各グループの代表要素だけパラメータを列挙して索引化 ----
            foreach (var g in _groups)
            {
                if (ct.IsCancellationRequested) yield break;
                if (g.Elements.Count == 0) continue;

                Element rep = null;
                try { rep = _doc.GetElement(g.Elements[0]); }
                catch { rep = null; }
                if (rep == null) continue;

                g.Label = BuildLabel(rep, g);

                ParameterSet pset = null;
                try { pset = rep.Parameters; }
                catch { pset = null; }
                if (pset == null) continue;

                foreach (Parameter p in pset)
                {
                    if (p == null) continue;

                    int pid = 0;
                    try { pid = p.Id != null ? p.Id.IntValue() : 0; }
                    catch { pid = 0; }

                    bool hit = pid > 0 && (targetParamIds == null || targetParamIds.Contains(pid));

                    // Id 照合が効かない環境への保険として GUID でも突き合わせる
                    if (!hit && targetGuids != null && targetGuids.Count > 0)
                    {
                        try
                        {
                            if (p.IsShared)
                            {
                                int mapped;
                                if (targetGuids.TryGetValue(p.GUID, out mapped))
                                {
                                    pid = mapped;
                                    hit = true;
                                }
                            }
                        }
                        catch { }
                    }

                    if (!hit) continue;

                    List<UsageGroup> list;
                    if (!_byParam.TryGetValue(pid, out list))
                    {
                        list = new List<UsageGroup>();
                        _byParam[pid] = list;
                    }
                    list.Add(g);
                }

                n++;
                if (n % 200 == 0) yield return n;
            }

            ScannedElementCount = n;
            IsBuilt = true;
            yield return n;
        }

        private void AddInstance(Element e, Dictionary<string, UsageGroup> map, HashSet<int> seen)
        {
            int eid;
            try { eid = e.Id.IntValue(); }
            catch { return; }
            if (!seen.Add(eid)) return;

            ElementId tid = ElementId.InvalidElementId;
            try { tid = e.GetTypeId() ?? ElementId.InvalidElementId; }
            catch { tid = ElementId.InvalidElementId; }

            ElementId cid = ElementId.InvalidElementId;
            string cname = "";
            try { var c = e.Category; if (c != null) { cid = c.Id; cname = c.Name ?? ""; } }
            catch { }

            // タイプを持つ要素はタイプ単位、持たない要素はカテゴリ＋クラス単位で束ねる
            string key = (tid != ElementId.InvalidElementId)
                ? "T" + tid.IntValue()
                : "C" + cid.IntValue() + "|" + e.GetType().Name;

            UsageGroup g;
            if (!map.TryGetValue(key, out g))
            {
                g = new UsageGroup
                {
                    IsElementType = false,
                    TypeId = tid,
                    CategoryId = cid,
                    CategoryName = cname
                };
                map[key] = g;
                _groups.Add(g);
            }
            g.Elements.Add(e.Id);
        }

        private void AddType(Element t, Dictionary<string, UsageGroup> map, HashSet<int> seen)
        {
            int tidInt;
            try { tidInt = t.Id.IntValue(); }
            catch { return; }
            if (!seen.Add(tidInt)) return;

            ElementId cid = ElementId.InvalidElementId;
            string cname = "";
            try { var c = t.Category; if (c != null) { cid = c.Id; cname = c.Name ?? ""; } }
            catch { }

            // 同一ファミリのタイプはパラメータ構成が同じ。システムタイプはクラス＋カテゴリで束ねる。
            string key;
            try
            {
                var fs = t as FamilySymbol;
                key = (fs != null && fs.Family != null)
                    ? "F" + fs.Family.Id.IntValue()
                    : "S" + cid.IntValue() + "|" + t.GetType().Name;
            }
            catch { key = "S" + cid.IntValue() + "|" + t.GetType().Name; }

            UsageGroup g;
            if (!map.TryGetValue(key, out g))
            {
                g = new UsageGroup
                {
                    IsElementType = true,
                    CategoryId = cid,
                    CategoryName = cname
                };
                map[key] = g;
                _groups.Add(g);
            }
            g.Elements.Add(t.Id);
        }

        private string BuildLabel(Element rep, UsageGroup g)
        {
            try
            {
                if (g.IsElementType)
                {
                    var et = rep as ElementType;
                    if (et != null && !string.IsNullOrEmpty(et.FamilyName))
                        return et.FamilyName;
                    return rep.Name ?? "";
                }

                if (g.TypeId != ElementId.InvalidElementId)
                {
                    var et = _doc.GetElement(g.TypeId) as ElementType;
                    if (et != null)
                    {
                        if (!string.IsNullOrEmpty(et.FamilyName))
                            return et.FamilyName + ": " + (et.Name ?? "");
                        return et.Name ?? "";
                    }
                }
                return g.CategoryName ?? "";
            }
            catch { return g.CategoryName ?? ""; }
        }

        /// <summary>指定パラメータを保持するグループ（存在しなければ空リスト）</summary>
        public List<UsageGroup> GetGroups(ElementId paramId)
        {
            if (paramId == null) return new List<UsageGroup>();
            List<UsageGroup> list;
            if (_byParam.TryGetValue(paramId.IntValue(), out list)) return list;
            return new List<UsageGroup>();
        }

        public void Clear()
        {
            _groups.Clear();
            _byParam.Clear();
            IsBuilt = false;
        }
    }
}

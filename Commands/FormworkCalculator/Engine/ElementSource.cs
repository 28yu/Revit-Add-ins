using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 要素の出自情報 (ホストドキュメント / リンクドキュメント / 配置トランスフォーム)。
    /// リンクモデル対応で導入。
    ///
    /// 設計方針:
    ///   - ホスト要素: SourceDoc = ホストドキュメント、Transform = Identity、IsLinked = false
    ///   - リンク要素: SourceDoc = リンクドキュメント、Transform = リンクの GetTotalTransform、IsLinked = true
    ///
    /// SurrogateId は ElementResult.ElementId 等で使う代理 ID。ホスト要素は元の ElementId.IntValue()
    /// を維持し、リンク要素は連番のマイナス値を割り当てて衝突を防ぐ。
    /// </summary>
    internal class ElementSource
    {
        public int SurrogateId;
        public Element Element;
        public Document SourceDoc;
        public Transform Transform;
        public string SourceName;            // "ホスト" or リンクファイル名
        public ElementId LinkInstanceId;     // ホストの場合は ElementId.InvalidElementId
        public bool IsLinked;
    }

    /// <summary>
    /// 全要素 (ホスト + リンク) の ElementSource を SurrogateId で索引するレジストリ。
    /// パイプラインの各段階で SurrogateId から実 Element / Document / Transform を解決するために使う。
    /// </summary>
    internal class ElementSourceRegistry
    {
        private readonly Dictionary<int, ElementSource> _byId = new Dictionary<int, ElementSource>();

        /// <summary>リンク要素用の SurrogateId カウンタ (マイナス値、ホスト要素 ID と衝突しない)。</summary>
        private int _linkedCounter = -1;

        /// <summary>
        /// ホスト要素を登録する。ElementId.IntValue() を SurrogateId としてそのまま使う。
        /// </summary>
        public ElementSource RegisterHost(Element elem, Document hostDoc)
        {
            int id = elem.Id.IntValue();
            if (_byId.TryGetValue(id, out var existing)) return existing;
            var src = new ElementSource
            {
                SurrogateId = id,
                Element = elem,
                SourceDoc = hostDoc,
                Transform = Transform.Identity,
                SourceName = HostSourceName,
                LinkInstanceId = ElementId.InvalidElementId,
                IsLinked = false,
            };
            _byId[id] = src;
            return src;
        }

        /// <summary>
        /// リンク要素を登録する。マイナス値の連番 SurrogateId を割り当てる。
        /// </summary>
        public ElementSource RegisterLinked(
            Element elem, Document linkDoc, Transform transform,
            string sourceName, ElementId linkInstanceId)
        {
            int id = _linkedCounter--;
            var src = new ElementSource
            {
                SurrogateId = id,
                Element = elem,
                SourceDoc = linkDoc,
                Transform = transform ?? Transform.Identity,
                SourceName = sourceName ?? string.Empty,
                LinkInstanceId = linkInstanceId ?? ElementId.InvalidElementId,
                IsLinked = true,
            };
            _byId[id] = src;
            return src;
        }

        /// <summary>SurrogateId から ElementSource を取得する。見つからなければ null。</summary>
        public ElementSource Get(int surrogateId)
        {
            return _byId.TryGetValue(surrogateId, out var src) ? src : null;
        }

        /// <summary>登録された全 ElementSource を返す。</summary>
        public IEnumerable<ElementSource> All() { return _byId.Values; }

        /// <summary>ホスト要素のソース名 (集計表で表示)。</summary>
        public const string HostSourceName = "ホスト";
    }
}

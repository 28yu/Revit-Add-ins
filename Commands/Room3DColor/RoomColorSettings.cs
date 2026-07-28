using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.Room3DColor
{
    /// <summary>
    /// 部屋の色分け基準
    /// </summary>
    public enum RoomColorBasis
    {
        /// <summary>部屋名ごと</summary>
        ByName,
        /// <summary>レベル（階）ごと</summary>
        ByLevel,
        /// <summary>指定パラメータの値ごと</summary>
        ByParameter,
        /// <summary>部屋ごとに個別</summary>
        PerRoom,
        /// <summary>すべて（全部屋を同色・1グループ）</summary>
        All
    }

    /// <summary>
    /// ダイアログに渡す部屋情報（色分け基準の算出用）
    /// </summary>
    public class RoomEntry
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public string LevelName { get; set; }
        /// <summary>パラメータ名 → 表示値</summary>
        public Dictionary<string, string> Parameters { get; set; }
            = new Dictionary<string, string>();
    }

    /// <summary>
    /// 色分けグループ（基準ごとにまとめた部屋の集合＋割当色）
    /// </summary>
    public class RoomColorGroup
    {
        /// <summary>グループの識別キー（内部使用）</summary>
        public string Key { get; set; }
        /// <summary>UI・凡例に表示するラベル</summary>
        public string Label { get; set; }
        /// <summary>このグループに属する部屋のId</summary>
        public List<ElementId> RoomIds { get; set; } = new List<ElementId>();
        /// <summary>割当色（ダイアログで変更可能）</summary>
        public Color Color { get; set; }
    }

    /// <summary>
    /// ダイアログの実行結果
    /// </summary>
    public class RoomColorResult
    {
        public RoomColorBasis Basis { get; set; }
        public string ParameterName { get; set; }
        public List<RoomColorGroup> Groups { get; set; } = new List<RoomColorGroup>();
        public string ViewName { get; set; }
        public string WorksetName { get; set; }
        public bool DeleteExisting { get; set; }
        public bool CreateLegend { get; set; }
    }
}

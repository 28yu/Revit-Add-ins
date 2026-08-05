using Autodesk.Revit.DB;

namespace Tools28.Commands.DwgLayerTransfer.Models
{
    /// <summary>
    /// DWG 1レイヤ分の「オブジェクトスタイル」設定。
    ///
    /// Revit で DWG レイヤの見た目は 2 段構えになっている。
    ///   1. オブジェクトスタイル（管理 &gt; オブジェクトスタイル &gt; 読み込みオブジェクト）… モデル全体の基準値
    ///   2. ビューごとの V/G 上書き … 1 に対する差分
    ///
    /// V/G は「上書きしていない項目」を持たない。つまり V/G だけを移しても、
    /// 上書きしていないレイヤの見た目はオブジェクトスタイル側で決まるため、
    /// モデル間でオブジェクトスタイルが違えば表示は一致しない。
    /// V/G だけの移行では不十分だったのはこのため。
    ///
    /// ElementId はモデル間で通用しないため、線種は名前で保持して移行先で解決する。
    /// </summary>
    public sealed class DwgObjectStyle
    {
        /// <summary>レイヤ名。空文字は DWG 本体（親カテゴリ）。</summary>
        public string LayerName { get; set; } = "";

        /// <summary>線の色。IsValid == false なら未設定。</summary>
        public Color LineColor { get; set; }

        /// <summary>投影の線の太さ。-1 なら未設定。</summary>
        public int ProjectionLineWeight { get; set; } = -1;

        /// <summary>切断の線の太さ。-1 なら未設定。</summary>
        public int CutLineWeight { get; set; } = -1;

        /// <summary>投影の線種名。null なら未設定。</summary>
        public string ProjectionLinePattern { get; set; }

        /// <summary>切断の線種名。null なら未設定。</summary>
        public string CutLinePattern { get; set; }

        public bool HasAny
            => (LineColor != null && LineColor.IsValid)
            || ProjectionLineWeight > 0
            || CutLineWeight > 0
            || ProjectionLinePattern != null
            || CutLinePattern != null;

        public string Describe()
        {
            string color = (LineColor != null && LineColor.IsValid)
                ? $"{LineColor.Red},{LineColor.Green},{LineColor.Blue}" : "-";
            return $"線色={color} 投影幅={ProjectionLineWeight} 切断幅={CutLineWeight} " +
                   $"投影線種={ProjectionLinePattern ?? "-"} 切断線種={CutLinePattern ?? "-"}";
        }
    }
}

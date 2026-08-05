using Autodesk.Revit.DB;

namespace Tools28.Commands.DwgLayerTransfer.Models
{
    /// <summary>
    /// 1ビュー内の「DWG 1レイヤ分」の表示設定スナップショット。
    ///
    /// 保持する項目は Revit の「表示/グラフィックスの上書き &gt; 読み込みカテゴリ」タブで
    /// 実際に編集できる項目に合わせている（表示/非表示・線の色/線幅/線種・ハーフトーン）。
    /// サーフェスパターンや透明度は DWG のジオメトリには効かないため意図的に対象外とし、
    /// 移行先の設定を無用にリセットしないようにしている。
    ///
    /// ElementId はモデル間で意味を持たないため、線種は「名前」で保持して
    /// 移行先で名前解決する（<see cref="SolidPatternMarker"/> は既定の実線を表す）。
    /// </summary>
    public sealed class LayerGraphicSetting
    {
        /// <summary>線種が「実線」（LinePatternElement.GetSolidPatternId）であることを表す予約名</summary>
        public const string SolidPatternMarker = "<solid>";

        /// <summary>レイヤ名。空文字は DWG 本体（親カテゴリ）の行を表す。</summary>
        public string LayerName { get; set; } = "";

        /// <summary>ビューで非表示にされているか</summary>
        public bool Hidden { get; set; }

        // --- 投影/サーフェス 線 ---
        /// <summary>投影線の色。IsValid == false なら上書きなし。</summary>
        public Color ProjectionLineColor { get; set; }

        /// <summary>投影線の線幅。-1 なら上書きなし。</summary>
        public int ProjectionLineWeight { get; set; } = -1;

        /// <summary>投影線の線種名。null なら上書きなし。</summary>
        public string ProjectionLinePattern { get; set; }

        // --- 切断 線 ---
        /// <summary>切断線の色。IsValid == false なら上書きなし。</summary>
        public Color CutLineColor { get; set; }

        /// <summary>切断線の線幅。-1 なら上書きなし。</summary>
        public int CutLineWeight { get; set; } = -1;

        /// <summary>切断線の線種名。null なら上書きなし。</summary>
        public string CutLinePattern { get; set; }

        /// <summary>ハーフトーン</summary>
        public bool Halftone { get; set; }

        /// <summary>色・線幅・線種・ハーフトーンのいずれかが上書きされているか</summary>
        public bool HasOverride
            => (ProjectionLineColor != null && ProjectionLineColor.IsValid)
            || ProjectionLineWeight > 0
            || ProjectionLinePattern != null
            || (CutLineColor != null && CutLineColor.IsValid)
            || CutLineWeight > 0
            || CutLinePattern != null
            || Halftone;

        /// <summary>既定状態から変化しているか（一覧の「設定あり」件数に使う）</summary>
        public bool HasAnySetting => Hidden || HasOverride;
    }
}

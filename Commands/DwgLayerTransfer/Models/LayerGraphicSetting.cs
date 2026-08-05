using Autodesk.Revit.DB;

namespace Tools28.Commands.DwgLayerTransfer.Models
{
    /// <summary>
    /// 1ビュー内の「DWG 1レイヤ分」の表示設定スナップショット。
    ///
    /// Revit の「表示/グラフィックスの上書き &gt; 読み込みカテゴリ」で設定できる項目を
    /// <see cref="OverrideGraphicSettings"/> の全項目ぶん保持する。
    ///
    /// ⚠️ 以前はサーフェスパターンや透明度を「DWG には効かない」と判断して除外していたが、
    /// それらに設定が入っていると移行しても見た目が変わらないため、除外はやめた。
    /// 何を移すかを勝手に間引かず、Revit が持っている項目はすべて移す。
    ///
    /// ElementId はモデル間で意味を持たないため、線種・塗潰しパターンは「名前」で保持して
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

        // --- サーフェス パターン ---
        /// <summary>サーフェス前景パターン名。null なら上書きなし。</summary>
        public string SurfaceFgPattern { get; set; }
        public Color SurfaceFgColor { get; set; }
        /// <summary>サーフェス前景パターンを表示するか（false は明示的な非表示＝上書きあり）</summary>
        public bool SurfaceFgVisible { get; set; } = true;

        public string SurfaceBgPattern { get; set; }
        public Color SurfaceBgColor { get; set; }
        public bool SurfaceBgVisible { get; set; } = true;

        // --- 切断 パターン ---
        public string CutFgPattern { get; set; }
        public Color CutFgColor { get; set; }
        public bool CutFgVisible { get; set; } = true;

        public string CutBgPattern { get; set; }
        public Color CutBgColor { get; set; }
        public bool CutBgVisible { get; set; } = true;

        /// <summary>透明度（0-100）。0 なら上書きなし。</summary>
        public int Transparency { get; set; }

        /// <summary>ハーフトーン</summary>
        public bool Halftone { get; set; }

        /// <summary>詳細レベル。Undefined なら上書きなし。</summary>
        public ViewDetailLevel DetailLevel { get; set; } = ViewDetailLevel.Undefined;

        private static bool HasColor(Color c) => c != null && c.IsValid;

        /// <summary>いずれかの項目が上書きされているか</summary>
        public bool HasOverride
            => HasColor(ProjectionLineColor)
            || ProjectionLineWeight > 0
            || ProjectionLinePattern != null
            || HasColor(CutLineColor)
            || CutLineWeight > 0
            || CutLinePattern != null
            || SurfaceFgPattern != null || HasColor(SurfaceFgColor) || !SurfaceFgVisible
            || SurfaceBgPattern != null || HasColor(SurfaceBgColor) || !SurfaceBgVisible
            || CutFgPattern != null || HasColor(CutFgColor) || !CutFgVisible
            || CutBgPattern != null || HasColor(CutBgColor) || !CutBgVisible
            || Transparency > 0
            || Halftone
            || DetailLevel != ViewDetailLevel.Undefined;

        /// <summary>既定状態から変化しているか（一覧の「設定あり」件数に使う）</summary>
        public bool HasAnySetting => Hidden || HasOverride;
    }
}

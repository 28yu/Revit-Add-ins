using System.Collections.Generic;

namespace Tools28.Commands.FormworkCalculator.Models
{
    public enum CalculationScope
    {
        EntireProject,
        CurrentView,
    }

    public enum ColorSchemeType
    {
        ByCategory,
        ByZone,
        ByFormworkType,
    }

    public class FormworkSettings
    {
        public CalculationScope Scope { get; set; } = CalculationScope.CurrentView;

        public bool GroupByCategory { get; set; } = true;
        public bool GroupByZone { get; set; }
        public string ZoneParameterName { get; set; } = string.Empty;
        public bool GroupByFormworkType { get; set; }
        public string FormworkTypeParameterName { get; set; } = string.Empty;

        public bool ExportToExcel { get; set; } = false;
        public bool CreateSchedule { get; set; } = true;
        public bool Create3DView { get; set; } = true;
        public bool CreateSheet { get; set; } = true;

        public ColorSchemeType ColorScheme { get; set; } = ColorSchemeType.ByCategory;
        public bool ShowDeductedFaces { get; set; }

        public bool UseGLDeduction { get; set; }
        public double GLElevationMeters { get; set; }

        // デバッグログ出力 (C:\temp\Formwork_debug.txt)
        // リリース時は false に変更すること
        public bool EnableDebugLog { get; set; } = true;

        // 構造柱・構造フレームから鉄骨部材 (H鋼・角形鋼管・CFT等) を自動除外する。
        // 暗黙挙動として常に true。デバッグ時のみ false にできる (UI には露出しない)。
        public bool ExcludeSteelMembers { get; set; } = true;

        // 階段カテゴリから鉄骨階段 (タイプ名・マテリアルに鉄骨キーワードを含む) を自動除外する。
        // 暗黙挙動として常に true。デバッグ時のみ false にできる (UI には露出しない)。
        public bool ExcludeSteelStairs { get; set; } = true;

        // 壁カテゴリから ALC/ECP パネル (タイプ名に "ALC" または "ECP" を含む) を自動除外する。
        // 工場製品の取付パネルで型枠不要のため。
        // 暗黙挙動として常に true。デバッグ時のみ false にできる (UI には露出しない)。
        public bool ExcludeAlcEcpPanels { get; set; } = true;

        // 壁カテゴリから LGS壁・乾式壁 (壁構造に石膏ボード層が含まれ、コンクリート層が無い壁)
        // を自動除外する。LGS = Light Gauge Steel (軽量鉄骨)。
        // 暗黙挙動として常に true。デバッグ時のみ false にできる (UI には露出しない)。
        public bool ExcludeLgsWalls { get; set; } = true;

        // 壁の天端が斜めの場合、斜面の幅（横方向射影）が以下のしきい値 (mm) 以上であれば
        // 型枠不要 (天端扱い) として控除する。0 以下にすると無効。
        public double SlopedWallTopWidthThresholdMm { get; set; } = 30.0;

        public List<string> IncludedCategories { get; set; } = new List<string>
        {
            "StructuralColumns",
            "StructuralFraming",
            "Walls",
            "Floors",
            "StructuralFoundation",
            "Stairs",
            "Roofs",
        };

        public string ExcelOutputPath { get; set; } = string.Empty;
    }
}

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
        public CalculationScope Scope { get; set; } = CalculationScope.EntireProject;

        public bool GroupByCategory { get; set; } = true;
        public bool GroupByZone { get; set; }
        public string ZoneParameterName { get; set; } = string.Empty;
        public bool GroupByFormworkType { get; set; }
        public string FormworkTypeParameterName { get; set; } = string.Empty;

        public bool ExportToExcel { get; set; } = true;
        public bool CreateSchedule { get; set; } = true;
        public bool Create3DView { get; set; } = true;

        public ColorSchemeType ColorScheme { get; set; } = ColorSchemeType.ByCategory;
        public bool ShowDeductedFaces { get; set; }

        public bool UseGLDeduction { get; set; }
        public double GLElevationMeters { get; set; }

        // デバッグログ出力 (C:\temp\Formwork_debug.txt)
        // リリース時は false に変更すること
        public bool EnableDebugLog { get; set; } = true;

        public List<string> IncludedCategories { get; set; } = new List<string>
        {
            "StructuralColumns",
            "StructuralFraming",
            "Walls",
            "Floors",
            "StructuralFoundation",
            "Stairs",
        };

        public string ExcelOutputPath { get; set; } = string.Empty;
    }
}

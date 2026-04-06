using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FireProtection
{
    public class FireProtectionParameterInfo
    {
        public string ParameterName { get; set; }
        public int DetectedCount { get; set; }
        public List<string> UniqueValues { get; set; }
    }

    public class LineStyleItem
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public override string ToString() => Name;
    }

    public class FillPatternItem
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public override string ToString() => Name;
    }

    public class FpTextNoteTypeItem
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }

        public FpTextNoteTypeItem(TextNoteType tnt)
        {
            Id = tnt.Id;
            Name = tnt.Name;
        }

        public override string ToString() => Name;
    }

    public class FireProtectionTypeEntry
    {
        public string Name { get; set; }
        public double OffsetMm { get; set; }
        // 前景
        public byte ColorR { get; set; }
        public byte ColorG { get; set; }
        public byte ColorB { get; set; }
        public bool ForegroundVisible { get; set; } = true;
        // 背景
        public byte BgColorR { get; set; } = 255;
        public byte BgColorG { get; set; } = 255;
        public byte BgColorB { get; set; } = 255;
        public bool BackgroundVisible { get; set; }
        // 柱の色設定（平面/天伏用）
        public byte ColColorR { get; set; }
        public byte ColColorG { get; set; }
        public byte ColColorB { get; set; }
        public bool ColForegroundVisible { get; set; } = true;
        public byte ColBgColorR { get; set; } = 255;
        public byte ColBgColorG { get; set; } = 255;
        public byte ColBgColorB { get; set; } = 255;
        public bool ColBackgroundVisible { get; set; }
    }

    public class FireProtectionDialogData
    {
        public string ViewName { get; set; }
        public string ViewTypeName { get; set; }
        public bool IsSectionView { get; set; }
        public int BeamCount { get; set; }
        public int ColumnCount { get; set; }
        public bool HasBeams { get; set; }
        public bool HasColumns { get; set; }
        public List<FireProtectionParameterInfo> BeamParameters { get; set; }
        public List<FireProtectionParameterInfo> ColumnParameters { get; set; }
        public List<FpTextNoteTypeItem> TextNoteTypes { get; set; }
        public List<LineStyleItem> LineStyles { get; set; }
        public List<FillPatternItem> FillPatterns { get; set; }
    }

    public class FireProtectionResult
    {
        public bool IncludeBeams { get; set; }
        public bool IncludeColumns { get; set; }
        public string SelectedParameterName { get; set; }
        public List<FireProtectionTypeEntry> Types { get; set; }
        public bool UseCommonOffset { get; set; }
        public double CommonOffsetMm { get; set; }
        public ElementId LineStyleId { get; set; }
        public ElementId FillPatternId { get; set; }
        public ElementId TextNoteTypeId { get; set; }
        public bool OverwriteExisting { get; set; }
        // 柱設定（平面/天伏ビュー用）
        public double ColumnA_mm { get; set; } = 400;
        public double ColumnB_mm { get; set; } = 150;
        // 柱の色は梁と共通（FireProtectionTypeEntryの色を使用）
        // 柱専用色が必要な場合はここに追加
    }

    public class MergeResult
    {
        public List<CurveLoop> MergedLoops { get; set; } = new List<CurveLoop>();
        public List<CurveLoop> UnmergedLoops { get; set; } = new List<CurveLoop>();
    }
}

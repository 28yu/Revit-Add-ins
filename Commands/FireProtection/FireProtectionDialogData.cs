using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FireProtection
{
    public class FireProtectionTypeEntry
    {
        public string Name { get; set; }
        public double OffsetMm { get; set; }
    }

    public class BeamTypeInfo
    {
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public string DisplayName => $"{FamilyName}: {TypeName}";
        public int Count { get; set; }
        public List<FamilyInstance> Beams { get; set; }
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

    public class FireProtectionDialogData
    {
        public string ViewName { get; set; }
        public int BeamCount { get; set; }
        public Level RefLevel { get; set; }
        public List<BeamTypeInfo> BeamTypes { get; set; }
        public List<FpTextNoteTypeItem> TextNoteTypes { get; set; }
        public List<LineStyleItem> LineStyles { get; set; }
        public List<FillPatternItem> FillPatterns { get; set; }
    }

    public class FireProtectionResult
    {
        public List<FireProtectionTypeEntry> Types { get; set; }
        public bool UseCommonOffset { get; set; }
        public double CommonOffsetMm { get; set; }
        public Dictionary<string, string> BeamTypeAssignments { get; set; }
        public ElementId LineStyleId { get; set; }
        public ElementId FillPatternId { get; set; }
        public ElementId TextNoteTypeId { get; set; }
        public bool OverwriteExisting { get; set; }
    }
}

using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamUnderLevel
{
    /// <summary>
    /// ダイアログに渡すデータ
    /// </summary>
    public class BeamUnderLevelDialogData
    {
        public string ViewName { get; set; }
        public int BeamCount { get; set; }
        public Level RefLevel { get; set; }
        public List<Level> LowerLevels { get; set; }
        public Level DefaultLowerLevel { get; set; }
        public Dictionary<string, List<FamilyInstance>> BeamsByFamily { get; set; }
        public Dictionary<string, List<ParamCandidate>> ParamCandidates { get; set; }
        public Dictionary<string, List<ParamCandidate>> TopLevelParamCandidates { get; set; }
        public Dictionary<string, List<string>> AdditionalLevelParams { get; set; }
        public List<TextNoteTypeItem> TextNoteTypes { get; set; }
    }

    /// <summary>
    /// TextNoteType表示用ラッパー
    /// </summary>
    public class TextNoteTypeItem
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }

        public TextNoteTypeItem(TextNoteType type)
        {
            Id = type.Id;
            Name = type.Name;
        }

        public override string ToString()
        {
            return Name;
        }
    }

    /// <summary>
    /// レベル表示用ラッパー
    /// </summary>
    public class LevelItem
    {
        public Level Level { get; set; }
        public string DisplayName { get; set; }

        public LevelItem(Level level)
        {
            Level = level;
            double elevMm = BeamCalculator.FeetToMm(level.Elevation);
            DisplayName = $"{level.Name} ({elevMm:+0;-0}mm)";
        }

        public override string ToString()
        {
            return DisplayName;
        }
    }
}

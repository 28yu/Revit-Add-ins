using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamTopLevel
{
    /// <summary>
    /// ダイアログに渡すデータ
    /// </summary>
    public class BeamTopLevelDialogData
    {
        public string ViewName { get; set; }
        public int BeamCount { get; set; }
        public Level RefLevel { get; set; }
        public Dictionary<string, List<FamilyInstance>> BeamsByFamily { get; set; }
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
}

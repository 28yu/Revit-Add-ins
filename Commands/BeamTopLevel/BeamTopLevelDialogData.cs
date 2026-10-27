using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.BeamCommon;

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
        /// <summary>構造フレームタグのタイプ（梁ラベルを梁タグで付ける場合）</summary>
        public List<BeamTagTypeItem> TagTypes { get; set; }
        /// <summary>表記する文字に対応する凡例が既にあるか</summary>
        public Func<string, bool> LegendExists { get; set; }
        /// <summary>対象の梁の水平・斜めごとの本数</summary>
        public BeamSlopeCounts SlopeCounts { get; set; }
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

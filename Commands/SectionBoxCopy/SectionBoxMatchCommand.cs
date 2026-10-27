using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Tools28.Commands.ViewCopy;
using Tools28.Localization;

namespace Tools28.Commands.SectionBoxCopy
{
    /// <summary>
    /// 切断ボックス合わせ（旧 切断ボックス コピー／ペーストを1ボタンにまとめたもの）。
    /// 流れは <see cref="View3DMatchCommandBase"/>。コピー元で切断ボックスが OFF なら、貼ると OFF になる。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class SectionBoxMatchCommand : View3DMatchCommandBase
    {
        protected override string KeyPrefix => "SbMatch";

        protected override bool HasCopied => SectionBoxClipboard.HasCopiedData;

        protected override string BadgeIcon => "sectionbox_paste.png";

        protected override string SourceViewName => SectionBoxClipboard.SourceViewName;

        protected override string SourceDescription =>
            SectionBoxClipboard.IsSectionBoxActive
                ? SectionBoxClipboard.SourceViewName
                : string.Format(Loc.S("SbMatch.SourceOff"), SectionBoxClipboard.SourceViewName);

        protected override string CopyFrom(View3D view)
        {
            return SectionBoxClipboard.CopyFrom(view) ? null : Loc.S("SectionBox.NoRange");
        }

        protected override void ApplyTo(View3D view) => SectionBoxClipboard.ApplyTo(view);

        protected override void ClearCopy() => SectionBoxClipboard.Clear();
    }
}

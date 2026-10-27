using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Tools28.Commands.ViewCopy;
using Tools28.Localization;

namespace Tools28.Commands.SectionBoxCopy
{
    /// <summary>
    /// 切断ボックス合わせ（旧 切断ボックス コピー／ペーストを1ボタンにまとめたもの）。
    /// 流れは <see cref="ViewSettingMatchCommandBase"/>。コピー元で切断ボックスが OFF なら、貼ると OFF になる。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class SectionBoxMatchCommand : ViewSettingMatchCommandBase
    {
        protected override string KeyPrefix => "SbMatch";

        protected override bool HasCopied => SectionBoxClipboard.HasCopiedData;

        protected override string BadgeIcon => "sectionbox_paste.png";

        protected override string SourceViewName => SectionBoxClipboard.SourceViewName;

        protected override string SourceDescription =>
            SectionBoxClipboard.IsSectionBoxActive
                ? SectionBoxClipboard.SourceViewName
                : string.Format(Loc.S("SbMatch.SourceOff"), SectionBoxClipboard.SourceViewName);

        protected override string CopyFrom(View view)
        {
            return SectionBoxClipboard.CopyFrom((View3D)view) ? null : Loc.S("SectionBox.NoRange");
        }

        protected override void ApplyTo(View view) => SectionBoxClipboard.ApplyTo((View3D)view);

        protected override void ClearCopy() => SectionBoxClipboard.Clear();
    }
}

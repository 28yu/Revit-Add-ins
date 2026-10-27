using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Tools28.Commands.ViewCopy;
using Tools28.Localization;

namespace Tools28.Commands.CropBoxCopy
{
    /// <summary>
    /// トリミング領域合わせ（旧 トリミング領域 コピー／ペーストを1ボタンにまとめたもの）。
    /// 流れは <see cref="ViewSettingMatchCommandBase"/>。主に3Dビュー以外で使うので、
    /// 視点・切断ボックスとの組み合わせ（どちらのボタンでも両方貼る）には入れない。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class CropBoxMatchCommand : ViewSettingMatchCommandBase
    {
        protected override string KeyPrefix => "CbMatch";

        protected override string BadgeIcon => "cropbox_copy.png";

        protected override bool CombinesWithOthers => false;

        protected override bool HasCopied => CropBoxClipboard.HasCopiedData;

        protected override string SourceViewName => CropBoxClipboard.SourceName;

        protected override string SourceDescription =>
            CropBoxClipboard.CopiedData != null && !CropBoxClipboard.CopiedData.IsActive
                ? string.Format(Loc.S("CbMatch.SourceOff"), CropBoxClipboard.SourceName)
                : CropBoxClipboard.SourceName;

        protected override bool Supports(View view) => CropBoxClipboard.IsSupported(view);

        protected override string NeedViewMessage(View activeView)
        {
            return activeView == null
                ? Loc.S("Common.NoActiveView")
                : string.Format(Loc.S("CropBox.UnsupportedViewType"), CropBoxClipboard.GetViewTypeName(activeView.ViewType));
        }

        protected override string CopyFrom(View view)
        {
            CropBoxClipboard.CopyFrom(view);
            return null;
        }

        protected override void ApplyTo(View view) => CropBoxClipboard.ApplyTo(view);

        protected override void ClearCopy() => CropBoxClipboard.Clear();
    }
}

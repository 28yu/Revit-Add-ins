using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;

namespace Tools28.Commands.ViewCopy
{
    /// <summary>
    /// 視点合わせ（旧 視点コピー／視点ペーストを1ボタンにまとめたもの）。流れは <see cref="ViewSettingMatchCommandBase"/>。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ViewMatchCommand : ViewSettingMatchCommandBase
    {
        protected override string KeyPrefix => "ViewMatch";

        protected override bool HasCopied => ViewOrientationClipboard.HasCopiedOrientation;

        protected override string BadgeIcon => "view_paste.png";

        protected override string SourceViewName => ViewOrientationClipboard.SourceViewName;

        protected override string SourceDescription => ViewOrientationClipboard.SourceViewName;

        protected override string CopyFrom(View view)
        {
            ViewOrientationClipboard.CopyFrom((View3D)view);
            return null;
        }

        protected override void ApplyTo(View view) => ViewOrientationClipboard.ApplyTo((View3D)view);

        protected override void ClearCopy() => ViewOrientationClipboard.Clear();
    }
}

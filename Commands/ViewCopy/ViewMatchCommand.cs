using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;

namespace Tools28.Commands.ViewCopy
{
    /// <summary>
    /// 視点合わせ（旧 視点コピー／視点ペーストを1ボタンにまとめたもの）。流れは <see cref="View3DMatchCommandBase"/>。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ViewMatchCommand : View3DMatchCommandBase
    {
        protected override string KeyPrefix => "ViewMatch";

        protected override bool HasCopied => ViewOrientationClipboard.HasCopiedOrientation;

        protected override string SourceDescription => ViewOrientationClipboard.SourceViewName;

        protected override string CopyFrom(View3D view)
        {
            ViewOrientationClipboard.CopyFrom(view);
            return null;
        }

        protected override void ApplyTo(View3D view) => ViewOrientationClipboard.ApplyTo(view);
    }
}

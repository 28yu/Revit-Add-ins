using Autodesk.Revit.DB;

namespace Tools28.Commands.ViewCopy
{
    // 視点情報を保存するための静的クラス（「視点合わせ」ボタンが使う）
    public static class ViewOrientationClipboard
    {
        public static ViewOrientation3D CopiedOrientation { get; set; }
        public static string SourceViewName { get; set; }
        public static bool HasCopiedOrientation => CopiedOrientation != null;

        public static void Clear()
        {
            CopiedOrientation = null;
            SourceViewName = null;
        }

        /// <summary>3Dビューの視点を覚える。</summary>
        public static void CopyFrom(View3D view)
        {
            CopiedOrientation = view.GetOrientation();
            SourceViewName = view.Name;
        }

        /// <summary>覚えた視点を3Dビューに当てる（トランザクション内で呼ぶ）。</summary>
        public static void ApplyTo(View3D view)
        {
            view.SetOrientation(CopiedOrientation);
        }
    }
}

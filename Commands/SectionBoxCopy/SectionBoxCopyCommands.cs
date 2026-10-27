using Autodesk.Revit.DB;

namespace Tools28.Commands.SectionBoxCopy
{
    // 切断ボックス情報を保持する静的クラス
    public static class SectionBoxClipboard
    {
        public static BoundingBoxXYZ CopiedSectionBox { get; set; }
        public static string SourceViewName { get; set; }

        /// <summary>コピー元で切断ボックスが有効だったか。</summary>
        public static bool IsSectionBoxActive { get; set; }

        /// <summary>
        /// コピーを1度でも実行したか。
        /// 「まだコピーしていない」と「コピー元が切断ボックスOFFだった」は
        /// どちらも CopiedSectionBox == null になるため、この独立したフラグで区別する。
        /// （旧実装は !IsSectionBoxActive を有効なコピーとみなしていたため、
        ///   起動直後にペーストすると対象ビューの切断ボックスが黙って解除されていた）
        /// </summary>
        public static bool HasCopied { get; set; }

        public static bool HasCopiedData => HasCopied;

        public static void Clear()
        {
            CopiedSectionBox = null;
            SourceViewName = null;
            IsSectionBoxActive = false;
            HasCopied = false;
        }

        /// <summary>
        /// 3Dビューの切断ボックスを覚える。切断ボックスが OFF のビューは「OFF」として覚える。
        /// 範囲を取得できなかったときは false。
        /// </summary>
        public static bool CopyFrom(View3D view)
        {
            BoundingBoxXYZ box = null;
            bool active = view.IsSectionBoxActive;
            if (active)
            {
                box = view.GetSectionBox();
                if (box == null) return false;
            }
            CopiedSectionBox = box;
            SourceViewName = view.Name;
            IsSectionBoxActive = active;
            HasCopied = true;
            return true;
        }

        /// <summary>覚えた切断ボックスを3Dビューに当てる（トランザクション内で呼ぶ）。</summary>
        public static void ApplyTo(View3D view)
        {
            if (IsSectionBoxActive && CopiedSectionBox != null)
            {
                // 切断ボックスを有効にして範囲を設定
                view.SetSectionBox(CopiedSectionBox);
            }
            else
            {
                // 切断ボックスを無効にする。
                // SetSectionBox(null) は Revit のバージョンによって例外になるため、
                // 明示的にプロパティで OFF にする。
                view.IsSectionBoxActive = false;
            }
        }
    }
}

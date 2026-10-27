using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.CropBoxCopy
{
    // トリミング領域データを保持するクラス
    public class CropBoxData
    {
        public bool IsActive { get; set; }
        public bool IsVisible { get; set; }
        public BoundingBoxXYZ CropBox { get; set; }
        public string SourceViewName { get; set; }
        public ViewType SourceViewType { get; set; }
        public string ViewTypeName { get; set; }
    }

    // 静的クリップボードクラス（「トリミング領域合わせ」ボタンが使う）
    public static class CropBoxClipboard
    {
        public static CropBoxData CopiedData { get; set; }
        public static string SourceName { get; set; }
        public static bool HasCopiedData => CopiedData != null;

        public static void Clear()
        {
            CopiedData = null;
            SourceName = null;
        }

        /// <summary>トリミング領域を扱えるビューか（平面・天井伏・断面・立面・詳細・3D）。</summary>
        public static bool IsSupported(View view)
        {
            ViewType viewType = view.ViewType;
            return viewType == ViewType.FloorPlan ||
                   viewType == ViewType.CeilingPlan ||
                   viewType == ViewType.Section ||
                   viewType == ViewType.Elevation ||
                   viewType == ViewType.Detail ||
                   viewType == ViewType.ThreeD;
        }

        public static string GetViewTypeName(ViewType viewType)
        {
            switch (viewType)
            {
                case ViewType.FloorPlan: return Loc.S("ViewType.FloorPlan");
                case ViewType.CeilingPlan: return Loc.S("ViewType.CeilingPlan");
                case ViewType.Section: return Loc.S("ViewType.Section");
                case ViewType.Elevation: return Loc.S("ViewType.Elevation");
                case ViewType.Detail: return Loc.S("ViewType.Detail");
                case ViewType.ThreeD: return Loc.S("ViewType.ThreeD");
                default: return viewType.ToString();
            }
        }

        /// <summary>ビューのトリミング領域（ON／OFF と表示の ON／OFF を含む）を覚える。</summary>
        public static void CopyFrom(View view)
        {
            CopiedData = new CropBoxData
            {
                IsActive = view.CropBoxActive,
                IsVisible = view.CropBoxActive ? view.CropBoxVisible : false,
                CropBox = view.CropBoxActive ? view.CropBox : null,
                SourceViewName = view.Name,
                SourceViewType = view.ViewType,
                ViewTypeName = GetViewTypeName(view.ViewType)
            };
            SourceName = view.Name;
        }

        /// <summary>覚えたトリミング領域をビューに当てる（トランザクション内で呼ぶ）。</summary>
        public static void ApplyTo(View view)
        {
            CropBoxData copied = CopiedData;
            if (copied.IsActive)
            {
                // コピー元がトリミングON → ペースト先もONにして領域・表示設定を写す
                view.CropBoxActive = true;
                if (copied.CropBox != null)
                    view.CropBox = copied.CropBox;
                view.CropBoxVisible = copied.IsVisible;
            }
            else
            {
                // コピー元がトリミングOFF → ペースト先もOFFにする
                view.CropBoxActive = false;
            }
        }
    }
}

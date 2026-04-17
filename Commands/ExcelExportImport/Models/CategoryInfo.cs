using Autodesk.Revit.DB;

namespace Tools28.Commands.ExcelExportImport.Models
{
    /// <summary>
    /// カテゴリ情報を保持するモデル
    /// </summary>
    public class CategoryInfo
    {
        public BuiltInCategory BuiltInCategory { get; set; }
        public string Name { get; set; }
        public bool IsChecked { get; set; }
        public int ElementCount { get; set; }

        /// <summary>表示用ラベル (「カテゴリ名 (件数)」)</summary>
        public string DisplayLabel => $"{Name} ({ElementCount})";

        public CategoryInfo(BuiltInCategory builtInCategory, string name, int elementCount)
        {
            BuiltInCategory = builtInCategory;
            Name = name;
            ElementCount = elementCount;
            IsChecked = false;
        }

        public override string ToString()
        {
            return $"{Name} ({ElementCount})";
        }
    }
}

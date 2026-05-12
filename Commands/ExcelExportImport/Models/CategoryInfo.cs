using Autodesk.Revit.DB;
using Tools28.Commands.ExcelExportImport.Services;

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

        /// <summary>表示用ラベル（Tools28言語設定に応じてローカライズ）</summary>
        public string DisplayLabel
        {
            get
            {
                string localizedName = CategoryLocalizer.GetLocalizedName(BuiltInCategory, Name);
                return $"{localizedName} ({ElementCount})";
            }
        }

        public CategoryInfo(BuiltInCategory builtInCategory, string name, int elementCount)
        {
            BuiltInCategory = builtInCategory;
            Name = name;
            ElementCount = elementCount;
            IsChecked = false;
        }

        public override string ToString()
        {
            return DisplayLabel;
        }
    }
}

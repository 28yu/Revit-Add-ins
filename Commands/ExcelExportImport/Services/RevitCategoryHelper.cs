using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.ExcelExportImport.Models;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// Revitカテゴリの取得ヘルパー
    /// </summary>
    public static class RevitCategoryHelper
    {
        /// <summary>
        /// ドキュメント内に存在する要素を持つカテゴリ一覧を取得
        /// </summary>
        public static List<CategoryInfo> GetCategoriesWithElements(Document doc)
        {
            var result = new List<CategoryInfo>();
            var categories = doc.Settings.Categories;

            foreach (Category cat in categories)
            {
                if (cat.CategoryType != CategoryType.Model && cat.CategoryType != CategoryType.AnalyticalModel)
                    continue;

                // サブカテゴリや不要なカテゴリをスキップ
                if (cat.Parent != null)
                    continue;

                var builtInCat = (BuiltInCategory)cat.Id.IntegerValue;

                // このカテゴリに属する要素数をカウント
                int count;
                try
                {
                    count = new FilteredElementCollector(doc)
                        .OfCategory(builtInCat)
                        .WhereElementIsNotElementType()
                        .GetElementCount();
                }
                catch
                {
                    continue;
                }

                if (count > 0)
                {
                    result.Add(new CategoryInfo(builtInCat, cat.Name, count));
                }
            }

            return result.OrderBy(c => c.Name).ToList();
        }

        /// <summary>
        /// 指定カテゴリの全要素を取得
        /// </summary>
        public static List<Element> GetElementsByCategory(Document doc, BuiltInCategory category)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .ToList();
        }
    }
}

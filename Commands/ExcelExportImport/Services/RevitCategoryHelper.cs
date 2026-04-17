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
        /// ドキュメント内（または指定スコープ内）に存在する要素を持つカテゴリ一覧を取得
        /// </summary>
        public static List<CategoryInfo> GetCategoriesWithElements(
            Document doc,
            ExportScope scope = ExportScope.EntireProject,
            View activeView = null,
            ICollection<ElementId> selectionIds = null)
        {
            var result = new List<CategoryInfo>();
            var categories = doc.Settings.Categories;

            // Selection スコープの場合、選択要素のカテゴリIDを先に集計
            HashSet<long> selectedCatIds = null;
            if (scope == ExportScope.Selection && selectionIds != null && selectionIds.Count > 0)
            {
                selectedCatIds = new HashSet<long>();
                foreach (var id in selectionIds)
                {
                    var e = doc.GetElement(id);
                    if (e?.Category == null) continue;
#if REVIT2026
                    selectedCatIds.Add(e.Category.Id.Value);
#else
                    selectedCatIds.Add(e.Category.Id.IntegerValue);
#endif
                }
            }

            foreach (Category cat in categories)
            {
                // モデル・解析・注釈（通り芯/レベル等）を対象にする
                if (cat.CategoryType != CategoryType.Model
                    && cat.CategoryType != CategoryType.AnalyticalModel
                    && cat.CategoryType != CategoryType.Annotation)
                    continue;

                // サブカテゴリは対象外
                if (cat.Parent != null)
                    continue;

#if REVIT2026
                var builtInCat = (BuiltInCategory)cat.Id.Value;
                long catIdValue = cat.Id.Value;
#else
                var builtInCat = (BuiltInCategory)cat.Id.IntegerValue;
                long catIdValue = cat.Id.IntegerValue;
#endif

                // 注釈カテゴリはパラメータエクスポートに有用なものだけ残す
                if (cat.CategoryType == CategoryType.Annotation
                    && !IsUsefulAnnotationCategory(builtInCat))
                    continue;

                int count;
                try
                {
                    count = GetElements(doc, builtInCat, scope, activeView, selectionIds).Count;
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
        /// 指定カテゴリの要素を取得（スコープ対応）
        /// </summary>
        public static List<Element> GetElementsByCategory(
            Document doc,
            BuiltInCategory category,
            ExportScope scope = ExportScope.EntireProject,
            View activeView = null,
            ICollection<ElementId> selectionIds = null)
        {
            return GetElements(doc, category, scope, activeView, selectionIds);
        }

        private static List<Element> GetElements(
            Document doc,
            BuiltInCategory category,
            ExportScope scope,
            View activeView,
            ICollection<ElementId> selectionIds)
        {
            FilteredElementCollector collector;

            switch (scope)
            {
                case ExportScope.ActiveView:
                    if (activeView == null)
                        return new List<Element>();
                    collector = new FilteredElementCollector(doc, activeView.Id);
                    break;

                case ExportScope.Selection:
                    if (selectionIds == null || selectionIds.Count == 0)
                        return new List<Element>();
                    collector = new FilteredElementCollector(doc, selectionIds);
                    break;

                case ExportScope.EntireProject:
                default:
                    collector = new FilteredElementCollector(doc);
                    break;
            }

            return collector
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .ToList();
        }

        /// <summary>
        /// パラメータエクスポートに有用な注釈カテゴリか判定
        /// </summary>
        private static bool IsUsefulAnnotationCategory(BuiltInCategory cat)
        {
            switch (cat)
            {
                case BuiltInCategory.OST_Grids:
                case BuiltInCategory.OST_Levels:
                case BuiltInCategory.OST_Sheets:
                case BuiltInCategory.OST_Views:
                case BuiltInCategory.OST_Viewports:
                case BuiltInCategory.OST_TextNotes:
                case BuiltInCategory.OST_GenericAnnotation:
                case BuiltInCategory.OST_RevisionClouds:
                case BuiltInCategory.OST_ScheduleGraphics:
                    return true;
                default:
                    return false;
            }
        }
    }
}

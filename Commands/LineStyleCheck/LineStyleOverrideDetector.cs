using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.LineStyleCheck
{
    /// <summary>
    /// ラインワークで線種が変更された要素の情報
    /// </summary>
    public class OverriddenElementInfo
    {
        public ElementId ElementId { get; set; }
        public string CategoryName { get; set; }
        public string ElementName { get; set; }
        public int OverriddenEdgeCount { get; set; }
        public List<string> OverriddenStyleNames { get; set; } = new List<string>();
        public OverrideGraphicSettings OriginalOverrides { get; set; }
    }

    /// <summary>
    /// ビュー内の要素のラインワーク変更を検出する
    /// </summary>
    public class LineStyleOverrideDetector
    {
        /// <summary>
        /// ビュー内でラインワーク変更されている要素を検出する
        /// </summary>
        public List<OverriddenElementInfo> FindOverriddenElements(Document doc, View view)
        {
            var result = new List<OverriddenElementInfo>();

            // ビュー内の全要素を取得（要素タイプは除外）
            var collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType();

            // Lines カテゴリ（ラインワークの変更先線種が属する）のサブカテゴリIDを収集
            HashSet<ElementId> linesCategoryStyleIds = GetLinesCategoryStyleIds(doc);

            foreach (Element elem in collector)
            {
                if (elem.Category == null) continue;

                // カテゴリが Lines 自体の場合はスキップ（詳細線やモデル線分等）
                if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Lines) continue;
                if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_SketchLines) continue;

                try
                {
                    var info = CheckElementForOverrides(doc, view, elem, linesCategoryStyleIds);
                    if (info != null)
                    {
                        result.Add(info);
                    }
                }
                catch
                {
                    // ジオメトリ取得に失敗する要素はスキップ
                }
            }

            return result;
        }

        /// <summary>
        /// Lines カテゴリの全サブカテゴリの GraphicsStyle ID を収集する
        /// </summary>
        private HashSet<ElementId> GetLinesCategoryStyleIds(Document doc)
        {
            var styleIds = new HashSet<ElementId>();

            Category linesCat = doc.Settings.Categories
                .get_Item(BuiltInCategory.OST_Lines);
            if (linesCat == null) return styleIds;

            foreach (Category subCat in linesCat.SubCategories)
            {
                GraphicsStyle projStyle =
                    subCat.GetGraphicsStyle(GraphicsStyleType.Projection);
                if (projStyle != null)
                    styleIds.Add(projStyle.Id);

                GraphicsStyle cutStyle =
                    subCat.GetGraphicsStyle(GraphicsStyleType.Cut);
                if (cutStyle != null)
                    styleIds.Add(cutStyle.Id);
            }

            // Lines カテゴリ自体のスタイルも追加
            GraphicsStyle linesProjStyle =
                linesCat.GetGraphicsStyle(GraphicsStyleType.Projection);
            if (linesProjStyle != null)
                styleIds.Add(linesProjStyle.Id);

            return styleIds;
        }

        /// <summary>
        /// 要素のカテゴリおよびサブカテゴリの GraphicsStyle ID を収集する
        /// </summary>
        private HashSet<ElementId> GetCategoryStyleIds(Category category)
        {
            var styleIds = new HashSet<ElementId>();
            if (category == null) return styleIds;

            GraphicsStyle projStyle =
                category.GetGraphicsStyle(GraphicsStyleType.Projection);
            if (projStyle != null)
                styleIds.Add(projStyle.Id);

            GraphicsStyle cutStyle =
                category.GetGraphicsStyle(GraphicsStyleType.Cut);
            if (cutStyle != null)
                styleIds.Add(cutStyle.Id);

            // サブカテゴリのスタイルも追加
            if (category.SubCategories != null)
            {
                foreach (Category subCat in category.SubCategories)
                {
                    GraphicsStyle subProjStyle =
                        subCat.GetGraphicsStyle(GraphicsStyleType.Projection);
                    if (subProjStyle != null)
                        styleIds.Add(subProjStyle.Id);

                    GraphicsStyle subCutStyle =
                        subCat.GetGraphicsStyle(GraphicsStyleType.Cut);
                    if (subCutStyle != null)
                        styleIds.Add(subCutStyle.Id);
                }
            }

            return styleIds;
        }

        /// <summary>
        /// 個別要素のラインワーク変更を検出する
        /// </summary>
        private OverriddenElementInfo CheckElementForOverrides(
            Document doc, View view, Element elem,
            HashSet<ElementId> linesCategoryStyleIds)
        {
            // 要素カテゴリの既定スタイルIDを取得
            HashSet<ElementId> categoryStyleIds = GetCategoryStyleIds(elem.Category);

            // ビュー固有のジオメトリを取得
            Options geomOptions = new Options
            {
                View = view,
                ComputeReferences = true
            };

            GeometryElement geomElem = elem.get_Geometry(geomOptions);
            if (geomElem == null) return null;

            int overriddenEdgeCount = 0;
            var overriddenStyleNames = new HashSet<string>();

            CheckGeometry(doc, geomElem, categoryStyleIds, linesCategoryStyleIds,
                ref overriddenEdgeCount, overriddenStyleNames);

            if (overriddenEdgeCount > 0)
            {
                // 現在のVGオーバーライドを保存
                OverrideGraphicSettings originalOgs =
                    view.GetElementOverrides(elem.Id);

                return new OverriddenElementInfo
                {
                    ElementId = elem.Id,
                    CategoryName = elem.Category?.Name ?? "不明",
                    ElementName = elem.Name ?? "",
                    OverriddenEdgeCount = overriddenEdgeCount,
                    OverriddenStyleNames = overriddenStyleNames.ToList(),
                    OriginalOverrides = originalOgs
                };
            }

            return null;
        }

        /// <summary>
        /// ジオメトリを再帰的に検査する
        /// </summary>
        private void CheckGeometry(
            Document doc,
            GeometryElement geomElem,
            HashSet<ElementId> categoryStyleIds,
            HashSet<ElementId> linesCategoryStyleIds,
            ref int overriddenEdgeCount,
            HashSet<string> overriddenStyleNames)
        {
            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid && solid.Faces.Size > 0)
                {
                    foreach (Edge edge in solid.Edges)
                    {
                        ElementId styleId = edge.GraphicsStyleId;
                        if (styleId == null || styleId == ElementId.InvalidElementId)
                            continue;

                        // エッジのスタイルが「Lines」カテゴリに属し、
                        // かつ要素自身のカテゴリスタイルでなければ、
                        // ラインワークで変更されたと判定
                        if (linesCategoryStyleIds.Contains(styleId)
                            && !categoryStyleIds.Contains(styleId))
                        {
                            overriddenEdgeCount++;
                            GraphicsStyle style =
                                doc.GetElement(styleId) as GraphicsStyle;
                            if (style != null)
                            {
                                overriddenStyleNames.Add(
                                    style.GraphicsStyleCategory?.Name ?? style.Name);
                            }
                        }
                    }
                }
                else if (geomObj is GeometryInstance instance)
                {
                    GeometryElement instanceGeom = instance.GetInstanceGeometry();
                    if (instanceGeom != null)
                    {
                        CheckGeometry(doc, instanceGeom, categoryStyleIds,
                            linesCategoryStyleIds, ref overriddenEdgeCount,
                            overriddenStyleNames);
                    }
                }
            }
        }
    }
}

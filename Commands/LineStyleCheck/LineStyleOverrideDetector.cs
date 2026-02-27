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

            var collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType();

            HashSet<ElementId> linesCategoryStyleIds = GetLinesCategoryStyleIds(doc);

            foreach (Element elem in collector)
            {
                if (elem.Category == null) continue;

                // ラインワークはモデル要素のエッジのみ変更可能
                if (elem.Category.CategoryType != CategoryType.Model) continue;

                // Lines カテゴリ自体の要素はスキップ
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
        /// 非ビュージオメトリからエッジの Reference → GraphicsStyleId マップを構築する
        /// </summary>
        private void BuildEdgeStyleMap(
            Document doc,
            GeometryElement geomElem,
            Dictionary<string, ElementId> styleMap)
        {
            if (geomElem == null) return;

            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid && solid.Faces.Size > 0)
                {
                    foreach (Edge edge in solid.Edges)
                    {
                        Reference edgeRef = edge.Reference;
                        if (edgeRef == null) continue;

                        try
                        {
                            string key = edgeRef.ConvertToStableRepresentation(doc);
                            ElementId styleId = edge.GraphicsStyleId;
                            styleMap[key] = styleId ?? ElementId.InvalidElementId;
                        }
                        catch
                        {
                            // Reference 変換に失敗するエッジはスキップ
                        }
                    }
                }
                else if (geomObj is GeometryInstance instance)
                {
                    GeometryElement instanceGeom = instance.GetInstanceGeometry();
                    if (instanceGeom != null)
                        BuildEdgeStyleMap(doc, instanceGeom, styleMap);
                }
            }
        }

        /// <summary>
        /// 個別要素のラインワーク変更を検出する
        ///
        /// 検出方式:
        /// 1. 非ビュージオメトリ（ラインワーク変更なし）の各エッジの Reference と StyleId を記録
        /// 2. ビュージオメトリ（ラインワーク変更あり）の各エッジと比較
        /// 3. 同一エッジ（Reference 一致）でスタイルが変わっていれば → ラインワーク変更
        /// 4. ビューにしか存在しないエッジ（断面カット等）→ カテゴリスタイル比較でフォールバック
        /// </summary>
        private OverriddenElementInfo CheckElementForOverrides(
            Document doc, View view, Element elem,
            HashSet<ElementId> linesCategoryStyleIds)
        {
            // カテゴリスタイル（フォールバック判定用）
            HashSet<ElementId> categoryStyleIds = GetCategoryStyleIds(elem.Category);

            // 非ビュージオメトリから Reference → StyleId マップを構築
            var defaultStyleByRef = new Dictionary<string, ElementId>();
            try
            {
                Options defaultOptions = new Options
                {
                    ComputeReferences = true
                };
                GeometryElement defaultGeom = elem.get_Geometry(defaultOptions);
                BuildEdgeStyleMap(doc, defaultGeom, defaultStyleByRef);
            }
            catch
            {
                // 非ビュージオメトリ取得失敗時はフォールバックのみで続行
            }

            // ビュー固有のジオメトリを取得
            Options viewOptions = new Options
            {
                View = view,
                ComputeReferences = true
            };

            GeometryElement viewGeom = elem.get_Geometry(viewOptions);
            if (viewGeom == null) return null;

            int overriddenEdgeCount = 0;
            var overriddenStyleNames = new HashSet<string>();

            CheckGeometryWithRefComparison(
                doc, viewGeom, defaultStyleByRef,
                categoryStyleIds, linesCategoryStyleIds,
                ref overriddenEdgeCount, overriddenStyleNames);

            if (overriddenEdgeCount > 0)
            {
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
        /// ビュージオメトリのエッジを非ビュージオメトリと比較して検査する
        /// </summary>
        private void CheckGeometryWithRefComparison(
            Document doc,
            GeometryElement geomElem,
            Dictionary<string, ElementId> defaultStyleByRef,
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
                        ElementId viewStyleId = edge.GraphicsStyleId;
                        if (viewStyleId == null || viewStyleId == ElementId.InvalidElementId)
                            continue;

                        // Lines カテゴリに属するスタイルのみ検査対象
                        if (!linesCategoryStyleIds.Contains(viewStyleId))
                            continue;

                        bool isOverride = false;
                        Reference edgeRef = edge.Reference;
                        string refKey = null;

                        if (edgeRef != null)
                        {
                            try
                            {
                                refKey = edgeRef.ConvertToStableRepresentation(doc);
                            }
                            catch
                            {
                                refKey = null;
                            }
                        }

                        if (refKey != null && defaultStyleByRef.TryGetValue(refKey, out ElementId defaultStyleId))
                        {
                            // 同一エッジが非ビューにも存在 → スタイル変化で判定
                            isOverride = (viewStyleId != defaultStyleId);
                        }
                        else
                        {
                            // ビューにしか存在しないエッジ（断面カット等）
                            // → カテゴリスタイルとの比較でフォールバック
                            isOverride = !categoryStyleIds.Contains(viewStyleId);
                        }

                        if (isOverride)
                        {
                            overriddenEdgeCount++;
                            GraphicsStyle style =
                                doc.GetElement(viewStyleId) as GraphicsStyle;
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
                        CheckGeometryWithRefComparison(
                            doc, instanceGeom, defaultStyleByRef,
                            categoryStyleIds, linesCategoryStyleIds,
                            ref overriddenEdgeCount, overriddenStyleNames);
                    }
                }
            }
        }
    }
}

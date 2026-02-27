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
    ///
    /// 検出方式: 新規ビュー比較 + ビュー所有要素スキャン
    ///   方式A: ViewPlan.Create() で新規ビューを作成し、エッジスタイルを比較
    ///   方式B: OwnedByView でビュー所有要素を検索してラインワーク情報を探す
    ///   診断: エッジスタイルの実際の値をダンプして、GraphicsStyleId の挙動を確認
    /// </summary>
    public class LineStyleOverrideDetector
    {
        private System.Text.StringBuilder _log;

        /// <summary>
        /// ビュー内でラインワーク変更されている要素を検出する
        /// </summary>
        public List<OverriddenElementInfo> FindOverriddenElements(Document doc, View view)
        {
            var result = new List<OverriddenElementInfo>();
            _log = new System.Text.StringBuilder();
            _log.AppendLine("=== LineStyleCheck Debug Log v7 (Diagnostic) ===");
            _log.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            _log.AppendLine($"View: {view.Name} (Type: {view.ViewType}, Id: {view.Id})");

            var sw = System.Diagnostics.Stopwatch.StartNew();

            // === 方式B: ビュー所有要素スキャン ===
            _log.AppendLine("");
            _log.AppendLine("=== Phase B: View-Owned Elements Scan ===");
            ScanViewOwnedElements(doc, view, result);

            // === 方式A: 新規ビュー比較 (診断用) ===
            _log.AppendLine("");
            _log.AppendLine("=== Phase A: Fresh View Comparison (Diagnostic) ===");
            DiagnosticFreshViewComparison(doc, view);

            sw.Stop();
            _log.AppendLine("");
            _log.AppendLine($"=== Result: {result.Count} elements with overrides ===");
            _log.AppendLine($"Elapsed: {sw.ElapsedMilliseconds}ms");
            WriteDebugLog();

            return result;
        }

        /// <summary>
        /// ビュー所有要素をスキャンしてラインワークオーバーライド情報を検出する
        /// </summary>
        private void ScanViewOwnedElements(Document doc, View view,
            List<OverriddenElementInfo> result)
        {
            try
            {
                // ビューが所有する全要素を取得
                var ownedCollector = new FilteredElementCollector(doc)
                    .OwnedByView(view.Id);

                int totalOwned = 0;
                var categoryCount = new Dictionary<string, int>();
                var classCount = new Dictionary<string, int>();

                // ラインワーク関連の要素を探す
                var lineworkElements = new List<Element>();

                foreach (Element elem in ownedCollector)
                {
                    totalOwned++;

                    string catName = elem.Category?.Name ?? "(no category)";
                    string className = elem.GetType().Name;
                    string key = $"{catName} [{className}]";

                    if (!categoryCount.ContainsKey(key))
                        categoryCount[key] = 0;
                    categoryCount[key]++;

                    if (!classCount.ContainsKey(className))
                        classCount[className] = 0;
                    classCount[className]++;

                    // ラインワーク関連と思われる要素を収集
                    // BuiltInCategory.OST_Lines, OST_SketchLines, or any Line-related
                    bool isLineRelated = false;
                    if (elem.Category != null)
                    {
                        int catId = elem.Category.Id.IntegerValue;
                        if (catId == (int)BuiltInCategory.OST_Lines
                            || catId == (int)BuiltInCategory.OST_LinesHiddenLines
                            || catId == (int)BuiltInCategory.OST_SketchLines
                            || catId == (int)BuiltInCategory.OST_MEPSpaceSeparationLines
                            || catId == (int)BuiltInCategory.OST_RoomSeparationLines
                            || catId == (int)BuiltInCategory.OST_AreaSchemeLines)
                        {
                            isLineRelated = true;
                        }
                    }

                    // CurveElement や GenericForm なども調査
                    if (isLineRelated || elem is CurveElement)
                    {
                        lineworkElements.Add(elem);
                    }
                }

                _log.AppendLine($"Total view-owned elements: {totalOwned}");
                _log.AppendLine("--- By Category+Class ---");
                foreach (var kvp in categoryCount.OrderByDescending(x => x.Value))
                {
                    _log.AppendLine($"  {kvp.Key}: {kvp.Value}");
                }
                _log.AppendLine("--- By Class ---");
                foreach (var kvp in classCount.OrderByDescending(x => x.Value))
                {
                    _log.AppendLine($"  {kvp.Key}: {kvp.Value}");
                }

                // ラインワーク要素の詳細をダンプ
                _log.AppendLine($"--- Line-Related Elements: {lineworkElements.Count} ---");
                int dumpCount = 0;
                foreach (var elem in lineworkElements)
                {
                    if (dumpCount >= 30) // 最大30個
                    {
                        _log.AppendLine("  ... (truncated)");
                        break;
                    }

                    string catName = elem.Category?.Name ?? "(null)";
                    string className = elem.GetType().Name;
                    _log.AppendLine($"  [{elem.Id}] {catName} / {className} / Name=\"{elem.Name}\"");

                    // GraphicsStyle を確認
                    if (elem is CurveElement curveElem)
                    {
                        try
                        {
                            var lineStyle = curveElem.LineStyle;
                            _log.AppendLine($"    LineStyle: {lineStyle?.Name ?? "(null)"} [{lineStyle?.Id}]");
                        }
                        catch (Exception ex)
                        {
                            _log.AppendLine($"    LineStyle error: {ex.Message}");
                        }
                    }

                    // パラメータをダンプ
                    DumpElementParameters(elem);

                    dumpCount++;
                }

                // ラインワーク要素からホスト要素を特定
                _log.AppendLine("");
                _log.AppendLine("--- Attempting to find linework host elements ---");
                var hostElementIds = new HashSet<ElementId>();
                foreach (var elem in lineworkElements)
                {
                    // ParameterSet から HOST_ID などを探す
                    try
                    {
                        foreach (Parameter param in elem.Parameters)
                        {
                            if (param.StorageType == StorageType.ElementId)
                            {
                                ElementId refId = param.AsElementId();
                                if (refId != null
                                    && refId != ElementId.InvalidElementId
                                    && refId.IntegerValue > 0)
                                {
                                    Element refElem = doc.GetElement(refId);
                                    if (refElem != null
                                        && refElem.Category != null
                                        && refElem.Category.CategoryType == CategoryType.Model)
                                    {
                                        if (hostElementIds.Add(refId))
                                        {
                                            _log.AppendLine($"  Potential host: [{refId}] " +
                                                $"{refElem.Category.Name} \"{refElem.Name}\" " +
                                                $"(via param: {param.Definition.Name})");
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                }

                // ビューが所有する要素の中で、カテゴリが "Lines" で
                // かつ通常の詳細線分やモデル線分ではないものを検出
                _log.AppendLine("");
                _log.AppendLine("--- All owned elements full dump (first 50) ---");
                var allOwnedCollector2 = new FilteredElementCollector(doc)
                    .OwnedByView(view.Id);
                int dumpCount2 = 0;
                foreach (Element elem in allOwnedCollector2)
                {
                    if (dumpCount2 >= 50) break;
                    string catName = elem.Category?.Name ?? "(null)";
                    int catId = elem.Category?.Id?.IntegerValue ?? 0;
                    string className = elem.GetType().Name;
                    _log.AppendLine($"  [{elem.Id}] cat={catName} catId={catId} " +
                        $"class={className} name=\"{elem.Name}\"");
                    dumpCount2++;
                }
            }
            catch (Exception ex)
            {
                _log.AppendLine($"ScanViewOwnedElements ERROR: {ex.Message}");
                _log.AppendLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// 要素のパラメータをダンプする（診断用）
        /// </summary>
        private void DumpElementParameters(Element elem)
        {
            try
            {
                foreach (Parameter param in elem.Parameters)
                {
                    string val = "(unreadable)";
                    try
                    {
                        switch (param.StorageType)
                        {
                            case StorageType.String:
                                val = param.AsString() ?? "(null)";
                                break;
                            case StorageType.Integer:
                                val = param.AsInteger().ToString();
                                break;
                            case StorageType.Double:
                                val = param.AsDouble().ToString("F4");
                                break;
                            case StorageType.ElementId:
                                val = param.AsElementId()?.ToString() ?? "(null)";
                                break;
                        }
                    }
                    catch { }

                    _log.AppendLine($"    P: {param.Definition.Name} = {val} ({param.StorageType})");
                }
            }
            catch { }
        }

        /// <summary>
        /// 新規ビューとの比較による診断ログ（エッジスタイル値をダンプ）
        /// </summary>
        private void DiagnosticFreshViewComparison(Document doc, View view)
        {
            using (TransactionGroup tg = new TransactionGroup(doc, "診断用比較"))
            {
                tg.Start();

                View cleanView = null;
                using (Transaction trans = new Transaction(doc, "一時ビュー作成"))
                {
                    trans.Start();
                    cleanView = CreateCleanView(doc, view);
                    trans.Commit();
                }

                if (cleanView != null)
                {
                    DiagnosticEdgeComparison(doc, view, cleanView);
                }
                else
                {
                    _log.AppendLine("Could not create clean view for diagnostics.");
                }

                tg.RollBack();
            }
        }

        /// <summary>
        /// 最初の数要素のエッジスタイルを詳細にダンプする（診断用）
        /// </summary>
        private void DiagnosticEdgeComparison(Document doc, View origView, View cleanView)
        {
            var collector = new FilteredElementCollector(doc, origView.Id)
                .WhereElementIsNotElementType();

            int logged = 0;

            foreach (Element elem in collector)
            {
                if (logged >= 5) break; // 最初の5要素だけ

                if (elem.Category == null) continue;
                if (elem.Category.CategoryType != CategoryType.Model) continue;
                if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Lines) continue;
                if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_SketchLines) continue;

                Options origOpts = new Options { View = origView };
                GeometryElement origGeom = elem.get_Geometry(origOpts);
                if (origGeom == null) continue;

                Options cleanOpts = new Options { View = cleanView };
                GeometryElement cleanGeom = elem.get_Geometry(cleanOpts);
                if (cleanGeom == null) continue;

                Options noViewOpts = new Options();
                GeometryElement noViewGeom = elem.get_Geometry(noViewOpts);

                _log.AppendLine($"");
                _log.AppendLine($"--- Element: {elem.Category.Name} \"{elem.Name}\" [{elem.Id}] ---");

                // エッジスタイルをダンプ
                DumpEdgeStyles(doc, "OrigView", origGeom);
                DumpEdgeStyles(doc, "CleanView", cleanGeom);
                if (noViewGeom != null)
                {
                    DumpEdgeStyles(doc, "NoView", noViewGeom);
                }

                logged++;
            }

            _log.AppendLine($"");
            _log.AppendLine($"Diagnostic elements logged: {logged}");
        }

        /// <summary>
        /// ジオメトリのエッジスタイルをダンプする
        /// </summary>
        private void DumpEdgeStyles(Document doc, string label, GeometryElement geom)
        {
            _log.AppendLine($"  [{label}] Edge styles:");
            int edgeIndex = 0;

            foreach (GeometryObject gObj in geom)
            {
                if (gObj is Solid solid && solid.Faces.Size > 0)
                {
                    foreach (Edge edge in solid.Edges)
                    {
                        if (edgeIndex >= 10) // 最初の10エッジだけ
                        {
                            _log.AppendLine($"    ... (truncated at {solid.Edges.Size} edges)");
                            break;
                        }

                        ElementId styleId = edge.GraphicsStyleId;
                        string styleName = "(null)";
                        string catName = "(null)";

                        if (styleId != null && styleId != ElementId.InvalidElementId)
                        {
                            GraphicsStyle gs = doc.GetElement(styleId) as GraphicsStyle;
                            if (gs != null)
                            {
                                styleName = gs.Name;
                                catName = gs.GraphicsStyleCategory?.Name ?? "(no cat)";
                            }
                        }

                        // エッジの位置も出力（同一エッジの特定用）
                        XYZ midpoint = null;
                        try
                        {
                            var curve = edge.AsCurve();
                            midpoint = curve.Evaluate(0.5, true);
                        }
                        catch { }

                        string posStr = midpoint != null
                            ? $"({midpoint.X:F2},{midpoint.Y:F2},{midpoint.Z:F2})"
                            : "(?)";

                        _log.AppendLine($"    E{edgeIndex}: id={styleId} " +
                            $"style=\"{styleName}\" cat=\"{catName}\" pos={posStr}");

                        edgeIndex++;
                    }

                    if (edgeIndex >= 10) break;
                }
                else if (gObj is GeometryInstance inst)
                {
                    GeometryElement instGeom = inst.GetInstanceGeometry();
                    if (instGeom != null)
                    {
                        _log.AppendLine($"    (GeometryInstance - nested)");
                        DumpEdgeStyles(doc, $"{label}/Inst", instGeom);
                    }
                }
            }

            if (edgeIndex == 0)
            {
                _log.AppendLine($"    (no edges found)");
            }
        }

        /// <summary>
        /// ラインワーク変更を一切持たないクリーンなビューを新規作成する
        /// </summary>
        private View CreateCleanView(Document doc, View view)
        {
            try
            {
                if (view is ViewPlan viewPlan)
                {
                    return CreateCleanPlanView(doc, viewPlan);
                }
                else if (view is ViewSection viewSection)
                {
                    return CreateCleanSectionView(doc, viewSection);
                }
                else if (view is View3D view3D)
                {
                    return CreateClean3DView(doc, view3D);
                }

                _log.AppendLine($"Unsupported view type: {view.ViewType}");
                return null;
            }
            catch (Exception ex)
            {
                _log.AppendLine($"CreateCleanView failed: {ex.Message}");
                return null;
            }
        }

        private View CreateCleanPlanView(Document doc, ViewPlan viewPlan)
        {
            Level level = viewPlan.GenLevel;
            if (level == null)
            {
                _log.AppendLine("ERROR: ViewPlan has no GenLevel");
                return null;
            }

            ViewFamily family;
            switch (viewPlan.ViewType)
            {
                case ViewType.CeilingPlan:
                    family = ViewFamily.CeilingPlan;
                    break;
                case ViewType.AreaPlan:
                    family = ViewFamily.AreaPlan;
                    break;
                default:
                    family = ViewFamily.FloorPlan;
                    break;
            }

            ViewFamilyType vft = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(x => x.ViewFamily == family);

            if (vft == null)
            {
                _log.AppendLine($"ERROR: ViewFamilyType not found for {family}");
                return null;
            }

            ViewPlan freshView = ViewPlan.Create(doc, vft.Id, level.Id);
            CopyViewSettings(viewPlan, freshView);

            try
            {
                PlanViewRange origRange = viewPlan.GetViewRange();
                freshView.SetViewRange(origRange);
            }
            catch { }

            _log.AppendLine($"Clean PlanView created: {freshView.Name} [{freshView.Id}]");
            return freshView;
        }

        private View CreateCleanSectionView(Document doc, ViewSection viewSection)
        {
            ViewFamilyType vft = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(x => x.ViewFamily == ViewFamily.Section);

            if (vft == null)
            {
                _log.AppendLine("ERROR: Section ViewFamilyType not found");
                return null;
            }

            BoundingBoxXYZ sectionBox = viewSection.CropBox;
            ViewSection freshView = ViewSection.CreateSection(doc, vft.Id, sectionBox);
            CopyViewSettings(viewSection, freshView);

            _log.AppendLine($"Clean SectionView created: {freshView.Name} [{freshView.Id}]");
            return freshView;
        }

        private View CreateClean3DView(Document doc, View3D view3D)
        {
            ViewFamilyType vft = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);

            if (vft == null)
            {
                _log.AppendLine("ERROR: 3D ViewFamilyType not found");
                return null;
            }

            View3D freshView = View3D.CreateIsometric(doc, vft.Id);
            CopyViewSettings(view3D, freshView);

            if (view3D.IsSectionBoxActive)
            {
                try
                {
                    freshView.IsSectionBoxActive = true;
                    freshView.SetSectionBox(view3D.GetSectionBox());
                }
                catch { }
            }

            _log.AppendLine($"Clean 3DView created: {freshView.Name} [{freshView.Id}]");
            return freshView;
        }

        private void CopyViewSettings(View source, View target)
        {
            try { target.Scale = source.Scale; } catch { }
            try { target.DetailLevel = source.DetailLevel; } catch { }

            try
            {
                target.CropBoxActive = source.CropBoxActive;
                if (source.CropBoxActive)
                    target.CropBox = source.CropBox;
            }
            catch { }

            try
            {
                Parameter srcPhase = source.get_Parameter(BuiltInParameter.VIEW_PHASE);
                Parameter tgtPhase = target.get_Parameter(BuiltInParameter.VIEW_PHASE);
                if (srcPhase != null && tgtPhase != null && !tgtPhase.IsReadOnly)
                    tgtPhase.Set(srcPhase.AsElementId());
            }
            catch { }
        }

        private void WriteDebugLog()
        {
            try
            {
                string dir = @"C:\temp";
                if (!System.IO.Directory.Exists(dir))
                    System.IO.Directory.CreateDirectory(dir);

                System.IO.File.WriteAllText(
                    System.IO.Path.Combine(dir, "Tools28_LineStyleCheck_debug.txt"),
                    _log.ToString());
            }
            catch
            {
                // ログ出力失敗は無視
            }
        }
    }
}

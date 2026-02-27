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
    /// 検出方式: IExportContext2D による2Dカスタムエクスポート比較
    ///   1. CustomExporter + IExportContext2D でオリジナルビューの表示ジオメトリを出力
    ///   2. ViewPlan.Create() でクリーンビューを作成し、同じく2Dエクスポート
    ///   3. 各要素のエクスポートデータ（エッジ数、カーブ数、線分数）を比較
    ///   4. 差異がある要素 = ラインワークによる変更
    ///
    /// IExportContext2D はレンダリングパイプラインを使用するため、
    /// 通常のジオメトリ API (Edge.GraphicsStyleId) では取得できない
    /// ラインワークオーバーライドの影響が反映される可能性がある。
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
            _log.AppendLine("=== LineStyleCheck Debug Log v8 (IExportContext2D) ===");
            _log.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            _log.AppendLine($"View: {view.Name} (Type: {view.ViewType}, Id: {view.Id})");

            var sw = System.Diagnostics.Stopwatch.StartNew();

            // === IExportContext2D 比較方式 ===
            _log.AppendLine("");
            _log.AppendLine("=== Phase C: IExportContext2D Comparison ===");
            DetectVia2DExport(doc, view, result);

            sw.Stop();
            _log.AppendLine("");
            _log.AppendLine($"=== Result: {result.Count} elements with overrides ===");
            _log.AppendLine($"Elapsed: {sw.ElapsedMilliseconds}ms");
            WriteDebugLog();

            return result;
        }

        /// <summary>
        /// IExportContext2D を使用してラインワーク変更を検出する
        /// </summary>
        private void DetectVia2DExport(Document doc, View view,
            List<OverriddenElementInfo> result)
        {
            try
            {
                // 2Dエクスポート対応ビューかチェック
                if (!Is2DExportableView(view))
                {
                    _log.AppendLine($"View type {view.ViewType} does not support 2D export.");
                    _log.AppendLine("Falling back to geometry comparison.");
                    FallbackGeometryComparison(doc, view, result);
                    return;
                }

                // オリジナルビューの2Dエクスポート
                _log.AppendLine("Exporting original view...");
                var origData = Export2DView(doc, view);
                _log.AppendLine($"  Original: {origData.Count} elements exported");

                // クリーンビューの作成とエクスポート
                Dictionary<int, ElementExportData> cleanData = null;

                using (TransactionGroup tg = new TransactionGroup(doc, "ラインワーク検出"))
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
                        _log.AppendLine("Exporting clean view...");
                        cleanData = Export2DView(doc, cleanView);
                        _log.AppendLine($"  Clean: {cleanData.Count} elements exported");
                    }
                    else
                    {
                        _log.AppendLine("ERROR: Could not create clean view.");
                    }

                    tg.RollBack();
                }

                if (cleanData == null)
                {
                    _log.AppendLine("Clean view export failed. Cannot compare.");
                    return;
                }

                // 比較
                _log.AppendLine("");
                _log.AppendLine("=== Comparing export data ===");
                CompareExportData(doc, view, origData, cleanData, result);
            }
            catch (Exception ex)
            {
                _log.AppendLine($"DetectVia2DExport ERROR: {ex.Message}");
                _log.AppendLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// 2Dエクスポート対応ビューか判定
        /// </summary>
        private bool Is2DExportableView(View view)
        {
            return view.ViewType == ViewType.FloorPlan
                || view.ViewType == ViewType.CeilingPlan
                || view.ViewType == ViewType.Section
                || view.ViewType == ViewType.Elevation
                || view.ViewType == ViewType.Detail
                || view.ViewType == ViewType.AreaPlan
                || view.ViewType == ViewType.EngineeringPlan;
        }

        /// <summary>
        /// IExportContext2D を使用してビューの2Dジオメトリをエクスポートする
        /// </summary>
        private Dictionary<int, ElementExportData> Export2DView(Document doc, View view)
        {
            var context = new LineworkExportContext(doc, _log);
            using (var exporter = new CustomExporter(doc, context))
            {
                exporter.IncludeGeometricObjects = true;
                exporter.Export2DIncludingAnnotationObjects = false;
                exporter.Export2DGeometricObjectsIncludingPatternLines = false;
                exporter.ShouldStopOnError = false;
                exporter.Export(view);
            }
            return context.Elements;
        }

        /// <summary>
        /// 2つのエクスポートデータを比較してラインワーク変更を検出する
        /// </summary>
        private void CompareExportData(
            Document doc, View view,
            Dictionary<int, ElementExportData> origData,
            Dictionary<int, ElementExportData> cleanData,
            List<OverriddenElementInfo> result)
        {
            int comparedCount = 0;
            int diffCount = 0;

            // 全要素の詳細ログ（最初の20要素）
            _log.AppendLine("--- Per-element export comparison (first 20 with data) ---");
            int loggedCount = 0;

            foreach (var kvp in origData)
            {
                int elemId = kvp.Key;
                var origEd = kvp.Value;

                // モデル要素のみ比較
                Element elem = doc.GetElement(new ElementId(elemId));
                if (elem == null) continue;
                if (elem.Category == null) continue;
                if (elem.Category.CategoryType != CategoryType.Model) continue;

                comparedCount++;

                // クリーンビューのデータと比較
                ElementExportData cleanEd = null;
                if (cleanData.ContainsKey(elemId))
                    cleanEd = cleanData[elemId];

                bool hasDifference = false;
                string diffDesc = "";

                if (cleanEd == null)
                {
                    // クリーンビューに存在しない → 比較不可
                    if (loggedCount < 20)
                    {
                        _log.AppendLine($"  [{elemId}] {elem.Category.Name} \"{elem.Name}\" " +
                            $"- NOT in clean view (orig: C={origEd.CurveCount} E={origEd.FaceEdgeCount} " +
                            $"S={origEd.SilhouetteCount} L={origEd.LineSegmentCount} P={origEd.PolylineSegmentCount})");
                        loggedCount++;
                    }
                    continue;
                }

                // 各カウントを比較
                if (origEd.CurveCount != cleanEd.CurveCount)
                {
                    hasDifference = true;
                    diffDesc += $"Curve:{origEd.CurveCount}→{cleanEd.CurveCount} ";
                }
                if (origEd.FaceEdgeCount != cleanEd.FaceEdgeCount)
                {
                    hasDifference = true;
                    diffDesc += $"FaceEdge:{origEd.FaceEdgeCount}→{cleanEd.FaceEdgeCount} ";
                }
                if (origEd.SilhouetteCount != cleanEd.SilhouetteCount)
                {
                    hasDifference = true;
                    diffDesc += $"Silhouette:{origEd.SilhouetteCount}→{cleanEd.SilhouetteCount} ";
                }
                if (origEd.LineSegmentCount != cleanEd.LineSegmentCount)
                {
                    hasDifference = true;
                    diffDesc += $"LineSeg:{origEd.LineSegmentCount}→{cleanEd.LineSegmentCount} ";
                }
                if (origEd.PolylineSegmentCount != cleanEd.PolylineSegmentCount)
                {
                    hasDifference = true;
                    diffDesc += $"PolySeg:{origEd.PolylineSegmentCount}→{cleanEd.PolylineSegmentCount} ";
                }

                // エッジスタイルの比較
                if (origEd.EdgeStyleIds.Count > 0 || (cleanEd != null && cleanEd.EdgeStyleIds.Count > 0))
                {
                    string origStyles = string.Join(",", origEd.EdgeStyleIds.OrderBy(x => x));
                    string cleanStyles = string.Join(",", cleanEd.EdgeStyleIds.OrderBy(x => x));
                    if (origStyles != cleanStyles)
                    {
                        hasDifference = true;
                        diffDesc += $"EdgeStyles differ ";
                    }
                }

                if (loggedCount < 20)
                {
                    string status = hasDifference ? "*** DIFF ***" : "same";
                    _log.AppendLine($"  [{elemId}] {elem.Category.Name} \"{elem.Name}\" - {status}");
                    _log.AppendLine($"    Orig:  C={origEd.CurveCount} E={origEd.FaceEdgeCount} " +
                        $"S={origEd.SilhouetteCount} L={origEd.LineSegmentCount} P={origEd.PolylineSegmentCount} " +
                        $"Styles=[{string.Join(",", origEd.EdgeStyleIds.Take(10))}]");
                    _log.AppendLine($"    Clean: C={cleanEd.CurveCount} E={cleanEd.FaceEdgeCount} " +
                        $"S={cleanEd.SilhouetteCount} L={cleanEd.LineSegmentCount} P={cleanEd.PolylineSegmentCount} " +
                        $"Styles=[{string.Join(",", cleanEd.EdgeStyleIds.Take(10))}]");
                    loggedCount++;
                }

                if (hasDifference)
                {
                    diffCount++;

                    OverrideGraphicSettings originalOgs =
                        view.GetElementOverrides(elem.Id);

                    result.Add(new OverriddenElementInfo
                    {
                        ElementId = elem.Id,
                        CategoryName = elem.Category?.Name ?? "不明",
                        ElementName = elem.Name ?? "",
                        OverriddenEdgeCount = 1,
                        OverriddenStyleNames = new List<string> { diffDesc.Trim() },
                        OriginalOverrides = originalOgs
                    });

                    _log.AppendLine($"  *** HIT: [{elemId}] {elem.Category.Name} \"{elem.Name}\" - {diffDesc}");
                }
            }

            // クリーンビューにのみ存在する要素をチェック
            int cleanOnlyCount = 0;
            foreach (var kvp in cleanData)
            {
                if (!origData.ContainsKey(kvp.Key))
                {
                    cleanOnlyCount++;
                    if (cleanOnlyCount <= 5)
                    {
                        Element elem = doc.GetElement(new ElementId(kvp.Key));
                        string name = elem != null
                            ? $"{elem.Category?.Name ?? "?"} \"{elem.Name}\""
                            : "?";
                        _log.AppendLine($"  Clean-only: [{kvp.Key}] {name}");
                    }
                }
            }

            _log.AppendLine($"");
            _log.AppendLine($"Model elements compared: {comparedCount}");
            _log.AppendLine($"Elements with differences: {diffCount}");
            _log.AppendLine($"Clean-only elements: {cleanOnlyCount}");
        }

        /// <summary>
        /// 2Dエクスポート非対応ビュー向けのフォールバック
        /// </summary>
        private void FallbackGeometryComparison(Document doc, View view,
            List<OverriddenElementInfo> result)
        {
            _log.AppendLine("Fallback: 3D views are not supported by IExportContext2D.");
            _log.AppendLine("Linework detection for 3D views is not currently possible.");
        }

        #region View Creation

        private View CreateCleanView(Document doc, View view)
        {
            try
            {
                if (view is ViewPlan viewPlan)
                    return CreateCleanPlanView(doc, viewPlan);
                else if (view is ViewSection viewSection)
                    return CreateCleanSectionView(doc, viewSection);

                _log.AppendLine($"Unsupported view type for clean view: {view.ViewType}");
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

        #endregion

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
            catch { }
        }
    }

    /// <summary>
    /// 要素ごとの2Dエクスポートデータ
    /// </summary>
    public class ElementExportData
    {
        public int CurveCount;
        public int FaceEdgeCount;
        public int SilhouetteCount;
        public int LineSegmentCount;
        public int PolylineSegmentCount;
        public List<int> EdgeStyleIds = new List<int>();
    }

    /// <summary>
    /// IExportContext2D 実装：2Dビューのジオメトリデータを要素別に収集する
    ///
    /// Revit の CustomExporter + IExportContext2D は、ビューの表示内容を
    /// レンダリングパイプラインを通じて出力する。これにより、通常のジオメトリAPI
    /// では取得できないラインワークオーバーライドの影響が反映される可能性がある。
    /// </summary>
    class LineworkExportContext : IExportContext2D
    {
        private readonly Document _doc;
        private readonly System.Text.StringBuilder _log;
        private ElementId _currentElementId = ElementId.InvalidElementId;
        private int _elementCount = 0;

        public Dictionary<int, ElementExportData> Elements { get; }
            = new Dictionary<int, ElementExportData>();

        public LineworkExportContext(Document doc, System.Text.StringBuilder log)
        {
            _doc = doc;
            _log = log;
        }

        #region IExportContextBase

        public bool Start()
        {
            return true;
        }

        public void Finish()
        {
            _log.AppendLine($"  Export finished: {_elementCount} elements processed");
        }

        public bool IsCanceled()
        {
            return false;
        }

        #endregion

        #region View/Element lifecycle

        public RenderNodeAction OnViewBegin(ViewNode node)
        {
            return RenderNodeAction.Proceed;
        }

        public void OnViewEnd(ElementId elementId)
        {
        }

        public RenderNodeAction OnElementBegin2D(ElementNode node)
        {
            _currentElementId = node.ElementId;
            _elementCount++;

            int id = node.ElementId.IntegerValue;
            if (!Elements.ContainsKey(id))
                Elements[id] = new ElementExportData();

            return RenderNodeAction.Proceed;
        }

        public void OnElementEnd2D(ElementNode node)
        {
            _currentElementId = ElementId.InvalidElementId;
        }

        // 3D element callbacks (required by interface but not used for 2D)
        public RenderNodeAction OnElementBegin(ElementId elementId)
        {
            return RenderNodeAction.Skip;
        }

        public void OnElementEnd(ElementId elementId)
        {
        }

        #endregion

        #region Instance/Link

        public RenderNodeAction OnInstanceBegin(InstanceNode node)
        {
            return RenderNodeAction.Proceed;
        }

        public void OnInstanceEnd(InstanceNode node)
        {
        }

        public RenderNodeAction OnLinkBegin(LinkNode node)
        {
            return RenderNodeAction.Proceed;
        }

        public void OnLinkEnd(LinkNode node)
        {
        }

        #endregion

        #region Face

        public RenderNodeAction OnFaceBegin(FaceNode node)
        {
            return RenderNodeAction.Proceed;
        }

        public void OnFaceEnd(FaceNode node)
        {
        }

        #endregion

        #region 2D Geometry callbacks

        public RenderNodeAction OnCurve(CurveNode node)
        {
            RecordForCurrentElement(d => d.CurveCount++);
            return RenderNodeAction.Proceed;
        }

        public RenderNodeAction OnPolyline(PolylineNode node)
        {
            return RenderNodeAction.Proceed;
        }

        public RenderNodeAction OnFaceEdge2D(FaceEdgeNode node)
        {
            RecordForCurrentElement(d =>
            {
                d.FaceEdgeCount++;

                // FaceEdge の GraphicsStyleId を取得してみる
                try
                {
                    Edge edge = node.GetFaceEdge();
                    if (edge != null)
                    {
                        ElementId styleId = edge.GraphicsStyleId;
                        if (styleId != null)
                            d.EdgeStyleIds.Add(styleId.IntegerValue);
                    }
                }
                catch { }
            });

            return RenderNodeAction.Proceed;
        }

        public RenderNodeAction OnFaceSilhouette2D(FaceSilhouetteNode node)
        {
            RecordForCurrentElement(d => d.SilhouetteCount++);
            return RenderNodeAction.Proceed;
        }

        public void OnLineSegment(LineSegment segment)
        {
            RecordForCurrentElement(d => d.LineSegmentCount++);
        }

        public void OnPolylineSegments(PolylineSegments segments)
        {
            RecordForCurrentElement(d => d.PolylineSegmentCount++);
        }

        public void OnText(TextNode node)
        {
            // テキストは無視
        }

        #endregion

        #region 3D callbacks (required by interface, not used)

        public void OnLight(LightNode node) { }
        public void OnRPC(RPCNode node) { }
        public void OnMaterial(MaterialNode node) { }
        public void OnPolymesh(PolymeshTopology node) { }

        #endregion

        private void RecordForCurrentElement(Action<ElementExportData> action)
        {
            if (_currentElementId != ElementId.InvalidElementId)
            {
                int id = _currentElementId.IntegerValue;
                if (Elements.ContainsKey(id))
                    action(Elements[id]);
            }
        }
    }
}

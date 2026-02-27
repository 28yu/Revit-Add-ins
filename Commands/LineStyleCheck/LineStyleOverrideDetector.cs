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
    /// 検出方式: 新規ビュー比較
    ///   1. ViewPlan.Create() でラインワーク変更が一切ないクリーンなビューを新規作成
    ///      （View.Duplicate はラインワーク変更を保持するため使用不可）
    ///   2. 同一要素のジオメトリを両ビューから取得
    ///   3. 同一位置のエッジのスタイルを1対1で比較
    ///   4. スタイルが異なるエッジ = ラインワークによる変更
    ///   5. TransactionGroup.RollBack() で一時ビューを自動削除
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
            _log.AppendLine("=== LineStyleCheck Debug Log ===");
            _log.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            _log.AppendLine($"View: {view.Name} (Type: {view.ViewType})");

            var sw = System.Diagnostics.Stopwatch.StartNew();

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
                    DetectOverrides(doc, view, cleanView, result);
                }
                else
                {
                    _log.AppendLine("ERROR: Could not create clean view. Detection aborted.");
                }

                // RollBack で一時ビューを自動削除（ドキュメントに変更を残さない）
                tg.RollBack();
            }

            sw.Stop();
            _log.AppendLine($"Detected: {result.Count} elements with overrides");
            _log.AppendLine($"Elapsed: {sw.ElapsedMilliseconds}ms");
            WriteDebugLog();

            return result;
        }

        /// <summary>
        /// ラインワーク変更を一切持たないクリーンなビューを新規作成する
        ///
        /// View.Duplicate(Duplicate) はラインワーク変更を保持するため使用不可。
        /// ViewPlan.Create() で完全に新規のビューを作成し、
        /// オリジナルビューの設定（ビュー範囲、スケール、詳細レベル等）をコピーする。
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

        /// <summary>
        /// クリーンな平面図ビューを作成する（FloorPlan / CeilingPlan / AreaPlan 対応）
        /// </summary>
        private View CreateCleanPlanView(Document doc, ViewPlan viewPlan)
        {
            Level level = viewPlan.GenLevel;
            if (level == null)
            {
                _log.AppendLine("ERROR: ViewPlan has no GenLevel");
                return null;
            }

            // ビュータイプに合った ViewFamilyType を取得
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

            // オリジナルの設定をコピー
            CopyViewSettings(viewPlan, freshView);

            // ビュー範囲をコピー（平面図固有）
            try
            {
                PlanViewRange origRange = viewPlan.GetViewRange();
                freshView.SetViewRange(origRange);
                _log.AppendLine("View range copied");
            }
            catch (Exception ex)
            {
                _log.AppendLine($"View range copy failed: {ex.Message}");
            }

            _log.AppendLine($"Clean PlanView created: {freshView.Name} [{freshView.Id}]");
            return freshView;
        }

        /// <summary>
        /// クリーンな断面図ビューを作成する（Section / Elevation 対応）
        /// </summary>
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

            // オリジナルの CropBox（断面の位置・方向・範囲を含む）を使用して断面図を作成
            BoundingBoxXYZ sectionBox = viewSection.CropBox;
            ViewSection freshView = ViewSection.CreateSection(doc, vft.Id, sectionBox);

            CopyViewSettings(viewSection, freshView);

            _log.AppendLine($"Clean SectionView created: {freshView.Name} [{freshView.Id}]");
            return freshView;
        }

        /// <summary>
        /// クリーンな3Dビューを作成する
        /// </summary>
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

            // セクションボックスをコピー
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

        /// <summary>
        /// ビューの共通設定をコピーする
        /// </summary>
        private void CopyViewSettings(View source, View target)
        {
            try { target.Scale = source.Scale; } catch { }
            try { target.DetailLevel = source.DetailLevel; } catch { }

            // トリミング領域をコピー
            try
            {
                target.CropBoxActive = source.CropBoxActive;
                if (source.CropBoxActive)
                    target.CropBox = source.CropBox;
            }
            catch { }

            // フェーズをコピー
            try
            {
                Parameter srcPhase = source.get_Parameter(BuiltInParameter.VIEW_PHASE);
                Parameter tgtPhase = target.get_Parameter(BuiltInParameter.VIEW_PHASE);
                if (srcPhase != null && tgtPhase != null && !tgtPhase.IsReadOnly)
                    tgtPhase.Set(srcPhase.AsElementId());
            }
            catch { }
        }

        /// <summary>
        /// オリジナルビューとクリーンビューを比較してラインワーク変更を検出する
        /// </summary>
        private void DetectOverrides(
            Document doc, View origView, View cleanView,
            List<OverriddenElementInfo> result)
        {
            var collector = new FilteredElementCollector(doc, origView.Id)
                .WhereElementIsNotElementType();

            int totalCount = 0;
            int checkedCount = 0;

            foreach (Element elem in collector)
            {
                totalCount++;
                if (elem.Category == null) continue;
                if (elem.Category.CategoryType != CategoryType.Model) continue;
                if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Lines) continue;
                if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_SketchLines) continue;

                try
                {
                    // クリーンビュー（ラインワークなし）でジオメトリ取得
                    Options cleanOpts = new Options { View = cleanView };
                    GeometryElement cleanGeom = elem.get_Geometry(cleanOpts);
                    if (cleanGeom == null) continue;

                    // オリジナルビュー（ラインワークあり）でジオメトリ取得
                    Options origOpts = new Options { View = origView };
                    GeometryElement origGeom = elem.get_Geometry(origOpts);
                    if (origGeom == null) continue;

                    checkedCount++;

                    int overriddenEdgeCount = 0;
                    var overriddenStyleNames = new HashSet<string>();

                    CompareGeometry(doc, origGeom, cleanGeom,
                        ref overriddenEdgeCount, overriddenStyleNames);

                    if (overriddenEdgeCount > 0)
                    {
                        OverrideGraphicSettings originalOgs =
                            origView.GetElementOverrides(elem.Id);

                        result.Add(new OverriddenElementInfo
                        {
                            ElementId = elem.Id,
                            CategoryName = elem.Category?.Name ?? "不明",
                            ElementName = elem.Name ?? "",
                            OverriddenEdgeCount = overriddenEdgeCount,
                            OverriddenStyleNames = overriddenStyleNames.ToList(),
                            OriginalOverrides = originalOgs
                        });

                        _log.AppendLine($"  HIT: {elem.Category.Name} \"{elem.Name}\" [{elem.Id}] " +
                            $"- {overriddenEdgeCount} edges → {string.Join(", ", overriddenStyleNames)}");
                    }
                }
                catch (Exception ex)
                {
                    _log.AppendLine($"  ERR: {elem.Category?.Name} [{elem.Id}]: {ex.Message}");
                }
            }

            _log.AppendLine($"Elements in view: {totalCount}");
            _log.AppendLine($"Model elements checked: {checkedCount}");
        }

        /// <summary>
        /// 2つのジオメトリのエッジスタイルを1対1で比較する
        /// </summary>
        private void CompareGeometry(
            Document doc,
            GeometryElement origGeom, GeometryElement cleanGeom,
            ref int overriddenEdgeCount,
            HashSet<string> overriddenStyleNames)
        {
            var origObjs = origGeom.ToList();
            var cleanObjs = cleanGeom.ToList();

            int count = Math.Min(origObjs.Count, cleanObjs.Count);

            for (int i = 0; i < count; i++)
            {
                if (origObjs[i] is Solid origSolid && cleanObjs[i] is Solid cleanSolid
                    && origSolid.Faces.Size > 0 && cleanSolid.Faces.Size > 0)
                {
                    CompareEdges(doc, origSolid.Edges, cleanSolid.Edges,
                        ref overriddenEdgeCount, overriddenStyleNames);
                }
                else if (origObjs[i] is GeometryInstance origInst
                    && cleanObjs[i] is GeometryInstance cleanInst)
                {
                    GeometryElement origInstGeom = origInst.GetInstanceGeometry();
                    GeometryElement cleanInstGeom = cleanInst.GetInstanceGeometry();

                    if (origInstGeom != null && cleanInstGeom != null)
                    {
                        CompareGeometry(doc, origInstGeom, cleanInstGeom,
                            ref overriddenEdgeCount, overriddenStyleNames);
                    }
                }
            }
        }

        /// <summary>
        /// 2つの EdgeArray を1対1で比較する
        /// </summary>
        private void CompareEdges(
            Document doc,
            EdgeArray origEdges, EdgeArray cleanEdges,
            ref int overriddenEdgeCount,
            HashSet<string> overriddenStyleNames)
        {
            int edgeCount = Math.Min(origEdges.Size, cleanEdges.Size);

            for (int j = 0; j < edgeCount; j++)
            {
                ElementId origStyleId = origEdges.get_Item(j).GraphicsStyleId;
                ElementId cleanStyleId = cleanEdges.get_Item(j).GraphicsStyleId;

                if (origStyleId == null || cleanStyleId == null) continue;
                if (origStyleId == cleanStyleId) continue;

                // スタイルが異なる → ラインワークで変更されている
                overriddenEdgeCount++;

                GraphicsStyle origStyle = doc.GetElement(origStyleId) as GraphicsStyle;
                if (origStyle != null)
                {
                    overriddenStyleNames.Add(
                        origStyle.GraphicsStyleCategory?.Name ?? origStyle.Name);
                }
            }
        }

        /// <summary>
        /// デバッグログを出力する
        /// </summary>
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

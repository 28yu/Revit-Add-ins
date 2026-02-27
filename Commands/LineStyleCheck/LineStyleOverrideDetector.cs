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
    /// 検出方式: ビュー複製比較
    ///   1. ViewDuplicateOption.Duplicate で一時ビューを作成
    ///      （ラインワーク変更・注釈・詳細項目を含まないクリーンなビュー）
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

                // ラインワーク変更を含まないクリーンなビューを作成
                View cleanView = null;
                using (Transaction trans = new Transaction(doc, "一時ビュー作成"))
                {
                    trans.Start();
                    try
                    {
                        ElementId dupId = view.Duplicate(ViewDuplicateOption.Duplicate);
                        cleanView = doc.GetElement(dupId) as View;
                        _log.AppendLine($"Clean view created: {cleanView?.Name} [{dupId}]");
                    }
                    catch (Exception ex)
                    {
                        _log.AppendLine($"View.Duplicate failed: {ex.Message}");
                    }
                    trans.Commit();
                }

                if (cleanView != null)
                {
                    DetectOverrides(doc, view, cleanView, result);
                }
                else
                {
                    _log.AppendLine("FALLBACK: Clean view creation failed. No detection possible.");
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

                // モデル要素のみ対象（注釈要素は Lines のスタイルを自然に持つため除外）
                if (elem.Category.CategoryType != CategoryType.Model) continue;

                // Lines カテゴリ自体の要素はスキップ
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
        ///
        /// 同一ビュー設定（crop/scale/detail level 等）から取得したジオメトリは
        /// 同一の構造（Solid数、Edge数、Edge順序）を持つため、
        /// インデックスベースの1対1比較が可能。
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

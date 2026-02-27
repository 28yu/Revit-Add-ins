using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Tools28.Commands.LineStyleCheck
{
    /// <summary>
    /// 操作種別
    /// </summary>
    public enum LineStyleRequestType
    {
        None,
        ChangeAll,
        PickAndChange,
        RemoveHighlight,
        Rescan
    }

    /// <summary>
    /// 線種選択肢の情報
    /// </summary>
    public class LineStyleItem
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }

        /// <summary>
        /// 「カテゴリ既定に戻す」を表す特別な値
        /// </summary>
        public bool IsResetToDefault { get; set; }

        public override string ToString() => Name;
    }

    /// <summary>
    /// モードレスダイアログからのリクエストを処理する外部イベントハンドラー
    /// </summary>
    public class LineStyleCheckEventHandler : IExternalEventHandler
    {
        public LineStyleRequestType RequestType { get; set; }
        public ElementId SelectedGraphicsStyleId { get; set; }
        public bool IsResetToDefault { get; set; }
        public List<ElementId> TargetElementIds { get; set; }
        public List<OverriddenElementInfo> OverriddenElements { get; set; }

        /// <summary>
        /// ダイアログへの参照（選択操作後にダイアログを再表示するため）
        /// </summary>
        public LineStyleCheckDialog Dialog { get; set; }

        /// <summary>
        /// 処理完了時の通知イベント
        /// </summary>
        public event Action<string> OperationCompleted;

        /// <summary>
        /// 再スキャン完了時の通知イベント
        /// </summary>
        public event Action<List<OverriddenElementInfo>> RescanCompleted;

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null) return;

            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            try
            {
                switch (RequestType)
                {
                    case LineStyleRequestType.ChangeAll:
                        ExecuteChangeAll(doc, view);
                        break;

                    case LineStyleRequestType.PickAndChange:
                        ExecutePickAndChange(uidoc, doc, view);
                        break;

                    case LineStyleRequestType.RemoveHighlight:
                        ExecuteRemoveHighlight(doc, view);
                        break;

                    case LineStyleRequestType.Rescan:
                        ExecuteRescan(doc, view);
                        break;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // ユーザーがキャンセル
                Dialog?.SafeShow();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("エラー", $"処理中にエラーが発生しました。\n{ex.Message}");
                Dialog?.SafeShow();
            }
            finally
            {
                RequestType = LineStyleRequestType.None;
            }
        }

        public string GetName() => "LineStyleCheckEventHandler";

        /// <summary>
        /// 全ハイライト要素の線種を一括変更する
        /// </summary>
        private void ExecuteChangeAll(Document doc, View view)
        {
            if (OverriddenElements == null || OverriddenElements.Count == 0) return;

            var elementIds = OverriddenElements.Select(e => e.ElementId).ToList();
            int count = ApplyLineStyleChange(doc, view, elementIds);

            OperationCompleted?.Invoke($"{count}個の要素の線種を変更しました。");
        }

        /// <summary>
        /// ユーザーが個別選択した要素の線種を変更する
        /// </summary>
        private void ExecutePickAndChange(UIDocument uidoc, Document doc, View view)
        {
            try
            {
                // ダイアログを一時非表示
                Dialog?.SafeHide();

                // ユーザーに要素を選択させる
                IList<Reference> refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    "線種を変更する要素を選択してください (Escで終了)");

                if (refs != null && refs.Count > 0)
                {
                    var selectedIds = refs.Select(r => r.ElementId).ToList();
                    int count = ApplyLineStyleChange(doc, view, selectedIds);
                    OperationCompleted?.Invoke($"{count}個の要素の線種を変更しました。");
                }
            }
            finally
            {
                Dialog?.SafeShow();
            }
        }

        /// <summary>
        /// ハイライトを解除する
        /// </summary>
        private void ExecuteRemoveHighlight(Document doc, View view)
        {
            if (OverriddenElements == null) return;

            using (Transaction trans = new Transaction(doc, "ハイライト解除"))
            {
                trans.Start();

                foreach (var info in OverriddenElements)
                {
                    // 元のオーバーライドに戻す（もしくはクリア）
                    if (info.OriginalOverrides != null)
                    {
                        view.SetElementOverrides(info.ElementId, info.OriginalOverrides);
                    }
                    else
                    {
                        view.SetElementOverrides(info.ElementId,
                            new OverrideGraphicSettings());
                    }
                }

                trans.Commit();
            }
        }

        /// <summary>
        /// 再スキャンを実行する
        /// </summary>
        private void ExecuteRescan(Document doc, View view)
        {
            var detector = new LineStyleOverrideDetector();
            var newResults = detector.FindOverriddenElements(doc, view);

            // ハイライト表示
            if (newResults.Count > 0)
            {
                using (Transaction trans = new Transaction(doc, "線種変更箇所のハイライト"))
                {
                    trans.Start();

                    OverrideGraphicSettings highlightOgs = new OverrideGraphicSettings();
                    Color red = new Color(255, 0, 0);
                    highlightOgs.SetProjectionLineColor(red);
                    highlightOgs.SetCutLineColor(red);
                    highlightOgs.SetProjectionLineWeight(5);
                    highlightOgs.SetCutLineWeight(5);

                    foreach (var info in newResults)
                    {
                        view.SetElementOverrides(info.ElementId, highlightOgs);
                    }

                    trans.Commit();
                }
            }

            OverriddenElements = newResults;
            RescanCompleted?.Invoke(newResults);
        }

        /// <summary>
        /// 線種変更を適用する
        /// </summary>
        private int ApplyLineStyleChange(Document doc, View view,
            List<ElementId> elementIds)
        {
            int count = 0;

            using (Transaction trans = new Transaction(doc, "線種変更"))
            {
                trans.Start();

                foreach (ElementId elemId in elementIds)
                {
                    if (IsResetToDefault)
                    {
                        // カテゴリ既定に戻す: VGオーバーライドをクリア
                        view.SetElementOverrides(elemId, new OverrideGraphicSettings());
                    }
                    else if (SelectedGraphicsStyleId != null
                        && SelectedGraphicsStyleId != ElementId.InvalidElementId)
                    {
                        // 選択された線種のプロパティを適用
                        GraphicsStyle selectedStyle =
                            doc.GetElement(SelectedGraphicsStyleId) as GraphicsStyle;

                        if (selectedStyle?.GraphicsStyleCategory != null)
                        {
                            Category styleCat = selectedStyle.GraphicsStyleCategory;

                            OverrideGraphicSettings ogs = new OverrideGraphicSettings();

                            // 線パターンを設定
                            ElementId patternId =
                                styleCat.GetLinePatternId(GraphicsStyleType.Projection);
                            if (patternId != null && patternId != ElementId.InvalidElementId)
                            {
                                ogs.SetProjectionLinePatternId(patternId);
                                ogs.SetCutLinePatternId(patternId);
                            }

                            // 線の太さを設定
                            int? weight =
                                styleCat.GetLineWeight(GraphicsStyleType.Projection);
                            if (weight.HasValue && weight.Value > 0)
                            {
                                ogs.SetProjectionLineWeight(weight.Value);
                                ogs.SetCutLineWeight(weight.Value);
                            }

                            // 線の色を設定
                            Color lineColor = styleCat.LineColor;
                            if (lineColor != null && lineColor.IsValid)
                            {
                                ogs.SetProjectionLineColor(lineColor);
                                ogs.SetCutLineColor(lineColor);
                            }

                            view.SetElementOverrides(elemId, ogs);
                        }
                    }

                    count++;
                }

                trans.Commit();
            }

            return count;
        }

        /// <summary>
        /// プロジェクト内の全線種を取得する
        /// </summary>
        public static List<LineStyleItem> GetAvailableLineStyles(Document doc)
        {
            var items = new List<LineStyleItem>();

            // 先頭に「カテゴリ既定に戻す」を追加
            items.Add(new LineStyleItem
            {
                Id = ElementId.InvalidElementId,
                Name = "＜カテゴリ既定に戻す＞",
                IsResetToDefault = true
            });

            // Lines カテゴリのサブカテゴリ（＝プロジェクトの線種）を取得
            Category linesCat = doc.Settings.Categories
                .get_Item(BuiltInCategory.OST_Lines);

            if (linesCat?.SubCategories != null)
            {
                var lineStyles = new List<LineStyleItem>();

                foreach (Category subCat in linesCat.SubCategories)
                {
                    GraphicsStyle projStyle =
                        subCat.GetGraphicsStyle(GraphicsStyleType.Projection);
                    if (projStyle != null)
                    {
                        lineStyles.Add(new LineStyleItem
                        {
                            Id = projStyle.Id,
                            Name = subCat.Name
                        });
                    }
                }

                // 名前順でソート
                items.AddRange(lineStyles.OrderBy(s => s.Name));
            }

            return items;
        }
    }
}

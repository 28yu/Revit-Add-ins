using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Tools28.Localization;

namespace Tools28.Commands.ViewportPosition
{
    /// <summary>
    /// ビューポート位置合わせ。Revit の「タイプ プロパティを一致」(MA) と同じ操作で、
    /// ビューポートの図の位置をそろえる（位置コピー／位置ペーストを1つにまとめたもの）。
    ///
    ///   1. コピー元のビューポートをクリック（実行前に1つ選んでいれば、それがコピー元）
    ///   2. 合わせたいビューポートを続けてクリック。クリックするたびにすぐ動く
    ///   3. 図枠（シートの何もない場所）をクリックするとコピー元を選び直す
    ///   4. Esc で終了
    ///
    /// コピー元は終了後も覚えておくので、別のシートを開いてボタンを押すと 2 から始まる。
    /// プロジェクトブラウザでシート・ビューを選んで押した場合は、従来の位置ペーストと同じく一括で合わせる。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ViewportMatchCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;

                var selected = uidoc.Selection.GetElementIds()
                    .Select(id => doc.GetElement(id))
                    .Where(e => e != null)
                    .ToList();

                // プロジェクトブラウザでシート・ビューを選んでいる場合は、従来の位置ペーストで一括処理。
                // 開いているシート自身はダブルクリックで開いた直後に選択に残ることがあるので数えない
                ElementId activeId = doc.ActiveView?.Id;
                if (selected.Any(e => e is View && e.Id != activeId))
                    return new ExecuteViewportPositionPasteCommand().Execute(commandData, ref message, elements);

                if (!(doc.ActiveView is ViewSheet))
                {
                    message = Loc.S("ViewportPos.NeedSheet");
                    return Result.Failed;
                }

                // 実行前にビューポートを1つだけ選んでいれば、それをコピー元にする
                var preselected = selected.OfType<Viewport>().ToList();
                if (preselected.Count == 1)
                {
                    if (!TrySetSource(doc, preselected[0])) return Result.Cancelled;
                }
                else if (preselected.Count > 1)
                {
                    message = Loc.S("ViewportPos.MultipleViewports");
                    return Result.Failed;
                }

                RunPickLoop(uiapp, doc);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = string.Format(Loc.S("ViewportPos.PasteError"), ex.Message);
                return Result.Failed;
            }
        }

        /// <summary>
        /// コピー元 → 合わせたいビューポート、のクリックを Esc まで繰り返す。
        /// </summary>
        private static void RunPickLoop(UIApplication uiapp, Document doc)
        {
            ElementId sourceId = SourceViewportIdOn(doc);

            while (true)
            {
                UIDocument uidoc = uiapp.ActiveUIDocument;
                View viewBeforePick = uidoc.ActiveView;
                Reference picked;

                try
                {
                    if (!ViewportPositionClipboard.HasCopiedPosition)
                    {
                        // 1. コピー元を選ぶ（ビューポートだけ選べる）
                        picked = uidoc.Selection.PickObject(ObjectType.Element,
                            new MatchSelectionFilter(allowTitleBlock: false),
                            Loc.S("ViewportPos.Match.PickSource"));
                        var src = doc.GetElement(picked) as Viewport;
                        if (src != null && TrySetSource(doc, src)) sourceId = src.Id;
                        continue;
                    }

                    // 2. 合わせたいビューポートを選ぶ（図枠のクリックはコピー元の選び直し）
                    picked = uidoc.Selection.PickObject(ObjectType.Element,
                        new MatchSelectionFilter(allowTitleBlock: true),
                        string.Format(Loc.S("ViewportPos.Match.PickTarget"),
                            ViewportPositionClipboard.SourceViewportInfo));
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    // Esc なら終了。クリック待ちの間に別のシートへ切り替えた場合は、そのシートで続ける
                    if (SwitchedToAnotherSheet(uiapp, doc, viewBeforePick)) continue;
                    return;
                }

                Element target = doc.GetElement(picked);
                if (!(target is Viewport vp))
                {
                    // 図枠（シートの何もない場所）: コピー元を空にして選び直す（MA と同じ）
                    ViewportPositionClipboard.Clear();
                    sourceId = null;
                    continue;
                }

                if (sourceId != null && vp.Id == sourceId) continue;   // コピー元自身は動かさない

                ExecuteViewportPositionPasteCommand.MoveViewports(doc,
                    new List<(ViewportAnchorInfo, Viewport)> { (CurrentSource(), vp) },
                    Loc.S("ViewportPos.Match.Txn"));
                try { uidoc.RefreshActiveView(); } catch { }
            }
        }

        /// <summary>コピー元として覚える。パースビューは図の位置を計算できないので断る。</summary>
        private static bool TrySetSource(Document doc, Viewport vp)
        {
            var anchor = ViewportAnchor.Capture(doc, vp);
            if (anchor.IsPerspective)
            {
                TaskDialog.Show(Loc.S("ViewportPos.Report.Title"), Loc.S("ViewportPos.CopyPerspective"));
                return false;
            }
            ViewportPositionClipboard.SetSource(doc, vp, anchor);
            return true;
        }

        private static ViewportAnchorInfo CurrentSource()
        {
            return ViewportPositionClipboard.SourceAnchor
                ?? new ViewportAnchorInfo { BoxCenter = ViewportPositionClipboard.CopiedPosition };
        }

        /// <summary>
        /// 覚えているコピー元が、このモデルのどのビューポートかを探す（コピー元自身をクリックしても動かさないため）。
        /// 見つからなければ null。
        /// </summary>
        private static ElementId SourceViewportIdOn(Document doc)
        {
            var src = ViewportPositionClipboard.SourceAnchor;
            if (src?.ViewName == null || ViewportPositionClipboard.SourceSheetName == null) return null;
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .FirstOrDefault(v =>
                    (doc.GetElement(v.SheetId) as ViewSheet)?.Name == ViewportPositionClipboard.SourceSheetName &&
                    (doc.GetElement(v.ViewId) as View)?.Name == src.ViewName)
                ?.Id;
        }

        /// <summary>クリック待ちが、同じモデルの別のシートを開いたことで取り消されたか。</summary>
        private static bool SwitchedToAnotherSheet(UIApplication uiapp, Document doc, View before)
        {
            try
            {
                UIDocument now = uiapp.ActiveUIDocument;
                if (now == null || !now.Document.Equals(doc)) return false;
                View active = now.ActiveView;
                return active is ViewSheet && before != null && active.Id != before.Id;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>ビューポート（と、選び直し用に図枠）だけをクリックできるようにする。</summary>
        private class MatchSelectionFilter : ISelectionFilter
        {
            private readonly bool _allowTitleBlock;

            public MatchSelectionFilter(bool allowTitleBlock)
            {
                _allowTitleBlock = allowTitleBlock;
            }

            public bool AllowElement(Element elem)
            {
                if (elem is Viewport) return true;
                return _allowTitleBlock && elem?.Category != null &&
                       elem.Category.Id.IntValue() == (int)BuiltInCategory.OST_TitleBlocks;
            }

            public bool AllowReference(Reference reference, XYZ position) => false;
        }
    }
}

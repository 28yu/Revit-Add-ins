using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Tools28.Localization;

namespace Tools28.Commands.ViewCopy
{
    /// <summary>
    /// 視点合わせを Ctrl+Z で元に戻せるようにする。
    ///
    /// 3Dビューの視点（カメラ）は Revit の「元に戻す」で戻らない（ビューキューブで回したときと同じ扱い）。
    /// 切断ボックスは要素なので戻る。そこで、貼る直前の視点を覚えておき、Ctrl+Z で視点合わせの操作が
    /// 取り消されたこと（DocumentChanged の TransactionUndone）を検知したら、覚えた視点を当て直す。
    /// DocumentChanged の中ではモデルを変更できないので、ExternalEvent に任せる。
    ///
    /// ⚠ 当て直しは新しい操作として記録されるため、そのあと Ctrl+Y（やり直し）はできない。
    /// </summary>
    internal static class ViewOrientationUndo
    {
        private class Record
        {
            public Document Doc;
            public string TransactionName;
            public List<KeyValuePair<ElementId, ViewOrientation3D>> Before;
        }

        /// <summary>覚えておく回数の上限（古いものから捨てる）。</summary>
        private const int MaxRecords = 50;

        private static readonly List<Record> _history = new List<Record>();
        private static readonly List<Record> _pending = new List<Record>();
        private static ExternalEvent _event;
        private static bool _subscribed;

        /// <summary>コマンドの中（API を使える状態）で呼ぶ。最初の1回だけ準備する。</summary>
        public static void EnsureReady(UIApplication uiapp)
        {
            if (_event == null) _event = ExternalEvent.Create(new RestoreHandler());
            if (!_subscribed)
            {
                uiapp.Application.DocumentChanged += OnDocumentChanged;
                _subscribed = true;
            }
        }

        /// <summary>貼る直前の視点を覚える（トランザクションの名前と一緒に）。</summary>
        public static void Remember(Document doc, string transactionName, IEnumerable<View3D> views)
        {
            var before = new List<KeyValuePair<ElementId, ViewOrientation3D>>();
            foreach (var v in views)
            {
                try { before.Add(new KeyValuePair<ElementId, ViewOrientation3D>(v.Id, v.GetOrientation())); }
                catch { }
            }
            if (before.Count == 0) return;

            _history.Add(new Record { Doc = doc, TransactionName = transactionName, Before = before });
            if (_history.Count > MaxRecords) _history.RemoveAt(0);
        }

        private static void OnDocumentChanged(object sender, DocumentChangedEventArgs e)
        {
            try
            {
                if (e.Operation != UndoOperation.TransactionUndone) return;
                Document doc = e.GetDocument();

                // まとめて複数を取り消した場合は名前が複数来る。新しい記録から順に取り出す
                foreach (string name in e.GetTransactionNames())
                {
                    int i = _history.FindLastIndex(r => r.TransactionName == name && r.Doc.Equals(doc));
                    if (i < 0) continue;
                    _pending.Add(_history[i]);
                    _history.RemoveAt(i);
                }

                if (_pending.Count > 0) _event?.Raise();
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[ViewMatch] 元に戻すの検知に失敗: {ex.Message}");
            }
        }

        private class RestoreHandler : IExternalEventHandler
        {
            public string GetName() => "Tools28 ViewMatch Undo";

            public void Execute(UIApplication app)
            {
                var records = _pending.ToList();
                _pending.Clear();
                if (records.Count == 0) return;

                try
                {
                    Document doc = records[0].Doc;
                    var restored = new List<View3D>();
                    using (var t = new Transaction(doc, Loc.S("ViewMatch.TxnUndo")))
                    {
                        t.Start();
                        // 新しい記録から順に当てるので、最後に残るのはいちばん古い（元の）視点
                        foreach (var record in records)
                        {
                            foreach (var kv in record.Before)
                            {
                                if (!(doc.GetElement(kv.Key) is View3D view)) continue;
                                try
                                {
                                    view.SetOrientation(kv.Value);
                                    restored.Add(view);
                                }
                                catch (Exception ex)
                                {
                                    DiagLog.Write($"[ViewMatch] 視点を戻せない view={view.Name}: {ex.Message}");
                                }
                            }
                        }
                        t.Commit();
                    }

                    UIDocument uidoc = app.ActiveUIDocument;
                    if (uidoc != null && uidoc.Document.Equals(doc))
                        View3DMatchCommandBase.RedrawOpenViews(uidoc, restored);
                    DiagLog.Write($"[ViewMatch] 元に戻す: {restored.Count} ビューの視点を戻した");
                }
                catch (Exception ex)
                {
                    DiagLog.Write($"[ViewMatch] 視点を戻すのに失敗: {ex.Message}");
                }
            }
        }
    }
}

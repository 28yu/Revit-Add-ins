using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace Tools28.Commands.RealtimeRender
{
    /// <summary>
    /// ライブ同期: 描画ウィンドウが開いている間、Revit でのモデルの変更を描画ウィンドウに伝える。
    /// 視点は伝えない（描画ウィンドウと Revit のビューは、それぞれ自由に動かせる）。
    ///
    /// - モデルが変わったら（DocumentChanged）、少し待ってから（続けて変更されることが多いので）Revit が手の空いたとき（Idling）に書き出す
    /// - 要素（壁・ドア・家具など）だけが変わったときは、変わった要素だけを書き出して「差分」（シーンファイル名＋.d＋通し番号）を書く。
    ///   描画ウィンドウは前の分を消して差分を足す（大きなモデルでも速い）
    /// - マテリアル・タイプ・3D ビュー自身・フィルタ・太陽などが変わったとき（多くの要素の見た目が変わる）は、全部を書き出して
    ///   シーンファイルを置き換える。描画ウィンドウは更新時刻を見て読み直す
    /// - 描画ウィンドウの「同期」ボタンで止めると「.pause」ができるので、その間は書き出さない（変更は覚えておき、再開したときに書き出す）
    /// - 描画ウィンドウが閉じられたら、その分の見張りをやめる。全部閉じたらイベントの登録も外す
    /// </summary>
    internal static class LiveSync
    {
        /// <summary>モデルが変わってから書き出すまで待つ時間（続けて変更されたときに何度も書き出さないため）。</summary>
        private static readonly TimeSpan Debounce = TimeSpan.FromSeconds(0.3);
        /// <summary>変わった要素がこれより多ければ、差分ではなく全部を書き出す（差分を作るより速いため）。</summary>
        private const int MaxDeltaElements = 3000;

        private class Session
        {
            public Process Process;
            public Document Doc;
            public ElementId ViewId;
            public string ScenePath;
            public bool Dirty;
            public DateTime LastChange;
            public HashSet<ElementId> ExportedIds = new HashSet<ElementId>();  // 描いている要素（削除されたかを調べる）
            public bool FullNeeded;                                             // 全部を書き出し直す
            public HashSet<ElementId> Changed = new HashSet<ElementId>();       // 書き出し直す要素（変わった・増えた）
            public HashSet<ElementId> Removed = new HashSet<ElementId>();       // 描画ウィンドウから消す要素（変わった・消えた）
            public uint Seq;                                                    // 書いた通し番号（最初の書き出しが 0）
        }

        private static readonly List<Session> Sessions = new List<Session>();
        private static UIApplication _uiApp;
        private static int _counter;

        /// <summary>描画ウィンドウごとのシーンファイルの場所。古いもの（2 日以上前）はついでに消す。</summary>
        public static string NewScenePath()
        {
            string dir = Path.Combine(Path.GetTempPath(), "Tools28", "Render");
            Directory.CreateDirectory(dir);
            try
            {
                foreach (string f in Directory.GetFiles(dir, "scene_*"))
                    if (File.GetLastWriteTimeUtc(f) < DateTime.UtcNow.AddDays(-2)) File.Delete(f);
            }
            catch
            {
                // 使用中のファイルなどは残す
            }
            _counter++;
            return Path.Combine(dir, $"scene_{Process.GetCurrentProcess().Id}_{_counter}.t28s");
        }

        public static void Start(UIApplication uiApp, Document doc, View3D view, string scenePath, Process process,
                                 HashSet<ElementId> exportedIds)
        {
            var session = new Session
            {
                Process = process,
                Doc = doc,
                ViewId = view.Id,
                ScenePath = scenePath,
                ExportedIds = exportedIds ?? new HashSet<ElementId>(),
            };
            Sessions.Add(session);
            if (_uiApp == null)
            {
                _uiApp = uiApp;
                uiApp.Idling += OnIdling;
                uiApp.Application.DocumentChanged += OnDocumentChanged;
            }
            DiagLog.Cmd("RealtimeRender", $"ライブ同期 開始 ({Sessions.Count} 個のウィンドウ)");
        }

        private static void Stop()
        {
            if (_uiApp == null) return;
            _uiApp.Idling -= OnIdling;
            _uiApp.Application.DocumentChanged -= OnDocumentChanged;
            _uiApp = null;
            DiagLog.Cmd("RealtimeRender", "ライブ同期 終了");
        }

        private static void OnDocumentChanged(object sender, DocumentChangedEventArgs e)
        {
            Document doc = e.GetDocument();
            foreach (Session s in Sessions)
            {
                if (!s.Doc.Equals(doc)) continue;
                bool changed;
                try
                {
                    // 凡例・シート・注記など、描いている 3D ビューの見た目に関係ない変更では書き出さない（Revit が重くならないように）
                    changed = Collect(doc, s, e);
                }
                catch (Exception ex)
                {
                    // 判断できないときは、念のため全部を書き出す
                    DiagLog.Cmd("RealtimeRender", $"ライブ同期: 変更の判定に失敗 {ex.Message}");
                    s.FullNeeded = true;
                    changed = true;
                }
                if (!changed) continue;
                if (s.Changed.Count + s.Removed.Count > MaxDeltaElements) s.FullNeeded = true;
                s.Dirty = true;
                s.LastChange = DateTime.UtcNow;
            }
        }

        /// <summary>
        /// 変更のうち、描いている 3D ビューの見た目に関係するものを覚える。関係するものがあれば true。
        /// 要素だけの変更は Changed / Removed に、多くの要素に関わる変更は FullNeeded にする。
        /// </summary>
        private static bool Collect(Document doc, Session s, DocumentChangedEventArgs e)
        {
            bool any = false;
            // 削除: 描いていた要素が消えたときだけ
            foreach (ElementId id in e.GetDeletedElementIds())
            {
                if (id == s.ViewId) { s.FullNeeded = true; any = true; }
                else if (s.ExportedIds.Contains(id)) { s.Removed.Add(id); any = true; }
            }

            var view = doc.GetElement(s.ViewId) as View3D;
            if (view == null) return any;
            foreach (ElementId id in e.GetModifiedElementIds())
            {
                switch (Classify(doc.GetElement(id), view))
                {
                    case Relevance.Full: s.FullNeeded = true; any = true; break;
                    case Relevance.Element:
                        s.Removed.Add(id);  // 前の形を消して、書き出し直す
                        s.Changed.Add(id);
                        any = true;
                        break;
                }
            }
            foreach (ElementId id in e.GetAddedElementIds())
            {
                switch (Classify(doc.GetElement(id), view))
                {
                    case Relevance.Full: s.FullNeeded = true; any = true; break;
                    case Relevance.Element: s.Changed.Add(id); any = true; break;
                }
            }
            return any;
        }

        private enum Relevance { None, Element, Full }

        /// <summary>
        /// 変わった要素が 3D ビューの見た目にどう関係するか。
        /// Element: その要素の形だけが変わる（壁・床・ファミリの置き場所・リンクの置き場所など）。
        /// Full: 多くの要素の見た目が変わりうる（マテリアル・外観・タイプ・3D ビュー自身・そのテンプレート・フィルタ・リンクのタイプ・太陽の設定）。
        /// None: 関係しない（特定のビューだけの要素（凡例・注記・寸法・文字など）、ほかのビュー・シート、部屋、その他の設定）。
        /// ほかのビュー（平面・断面・集計表など）で壁を動かしたり値を変えたりしたときは、その要素自体が「変わった要素」として届く。
        /// </summary>
        private static Relevance Classify(Element e, View3D view)
        {
            if (e == null) return Relevance.None;
            if (e.Id == view.Id) return Relevance.Full;  // 表示・非表示、切断ボックス、表示スタイルなど
            if (e is View v) return v.Id == view.ViewTemplateId ? Relevance.Full : Relevance.None;
            if (e.OwnerViewId != ElementId.InvalidElementId) return Relevance.None;  // 特定のビューだけの要素（凡例・注記・詳細）
            if (e is Material || e is AppearanceAssetElement) return Relevance.Full;
            if (e is FilterElement) return view.GetFilters().Contains(e.Id) ? Relevance.Full : Relevance.None;
            if (e is RevitLinkType || e is SunAndShadowSettings) return Relevance.Full;
            if (e is RevitLinkInstance) return Relevance.Element;
            if (e is SpatialElement) return Relevance.None;  // 部屋・エリア・スペースは描かない
            Category c = e.Category;
            if (c == null || c.CategoryType != CategoryType.Model) return Relevance.None;  // 注釈（文字・寸法・タグなど）、フェーズなど
            // タイプ（壁のタイプ・ファミリのタイプなど）は、そのタイプの要素すべての形が変わる
            return e is ElementType ? Relevance.Full : Relevance.Element;
        }

        private static void OnIdling(object sender, IdlingEventArgs e)
        {
            // 閉じられたウィンドウ・閉じられたモデルの分はやめる
            Sessions.RemoveAll(s => HasExited(s.Process) || !s.Doc.IsValidObject);
            if (Sessions.Count == 0)
            {
                Stop();
                return;
            }

            bool waiting = false;
            foreach (Session s in Sessions)
            {
                if (File.Exists(s.ScenePath + ".pause")) continue;
                var view = s.Doc.GetElement(s.ViewId) as View3D;
                if (view == null) continue;

                if (s.Dirty)
                {
                    if (DateTime.UtcNow - s.LastChange < Debounce)
                    {
                        waiting = true;
                        continue;
                    }
                    s.Dirty = false;
                    UpdateScene(s, view);
                }
            }
            // 待っている変更があるときは、待ち時間が過ぎたらすぐ呼ばれるようにする（普段は Revit の操作があったときだけ呼ばれる）
            if (waiting) e.SetRaiseWithoutDelay();
        }

        private static void UpdateScene(Session s, View3D view)
        {
            if (!s.FullNeeded)
            {
                if (UpdateDelta(s, view)) return;
                s.FullNeeded = true;  // 差分を書けなかったときは、全部を書き出す
            }
            try
            {
                var watch = Stopwatch.StartNew();
                var exportedIds = new HashSet<ElementId>();
                RenderScene scene = RealtimeRenderCommand.ExportScene(s.Doc, view, exportedIds);
                if (scene.TriangleCount == 0) return;
                scene.SyncSeq = s.Seq + 1;
                // 書き終わってから置き換える（描画ウィンドウが書きかけのファイルを読まないように）
                string temp = s.ScenePath + ".tmp";
                SceneWriter.Write(temp, scene);
                try
                {
                    // 置き換えは一度に行われるので、描画ウィンドウが読みかけの中身を読むことはない
                    if (File.Exists(s.ScenePath)) File.Replace(temp, s.ScenePath, null);
                    else File.Move(temp, s.ScenePath);
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    // 描画ウィンドウが読んでいる最中だった。少し後にもう一度書く
                    s.Dirty = true;
                    s.LastChange = DateTime.UtcNow;
                    return;
                }
                s.ExportedIds = exportedIds;
                s.Seq++;
                s.FullNeeded = false;
                s.Changed.Clear();
                s.Removed.Clear();
                DiagLog.Cmd("RealtimeRender", $"ライブ同期: モデル全体を更新 #{s.Seq} ({scene.TriangleCount} 三角形・描く数 {scene.DrawnTriangleCount}・形 {scene.Meshes.Count}・置き場所 {scene.Instances.Count}・{watch.ElapsedMilliseconds} ms)");
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"ライブ同期: 書き出しに失敗 {ex.Message}");
            }
        }

        /// <summary>
        /// 変わった要素だけを書き出して差分を書く。書けたら true。
        /// 差分のファイル名は「シーンファイル＋.d＋通し番号（8 桁）」。描画ウィンドウは通し番号の順に足していく
        /// </summary>
        private static bool UpdateDelta(Session s, View3D view)
        {
            try
            {
                var watch = Stopwatch.StartNew();
                var exportedIds = new HashSet<ElementId>();
                // 消えた要素は書き出せないので除く（描画ウィンドウからは消す）
                var only = new HashSet<ElementId>();
                foreach (ElementId id in s.Changed)
                    if (s.Doc.GetElement(id) != null) only.Add(id);
                RenderScene scene = only.Count > 0
                    ? RealtimeRenderCommand.ExportScene(s.Doc, view, exportedIds, only)
                    : new RenderScene();
                uint seq = s.Seq + 1;
                var removed = new List<long>();
                foreach (ElementId id in s.Removed) removed.Add(id.LongValue());
                string path = s.ScenePath + ".d" + seq.ToString("D8");
                string temp = path + ".tmp";
                SceneWriter.WriteDelta(temp, scene, seq, removed);
                File.Move(temp, path);  // 書き終わってから見えるようにする（描画ウィンドウは .tmp を読まない）
                s.Seq = seq;
                foreach (ElementId id in s.Removed) s.ExportedIds.Remove(id);
                s.ExportedIds.UnionWith(exportedIds);
                DiagLog.Cmd("RealtimeRender", $"ライブ同期: 差分 #{seq} (変わった要素 {only.Count}・消す要素 {removed.Count}・{scene.TriangleCount} 三角形・{watch.ElapsedMilliseconds} ms)");
                s.Changed.Clear();
                s.Removed.Clear();
                return true;
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"ライブ同期: 差分を書けませんでした（全部を書き出します） {ex.Message}");
                return false;
            }
        }

        private static bool HasExited(Process p)
        {
            try
            {
                return p.HasExited;
            }
            catch
            {
                return true;
            }
        }
    }
}

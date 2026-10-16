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
    /// - モデルが変わったら（DocumentChanged）、少し待ってから（続けて変更されることが多いので）Revit が手の空いたとき（Idling）に
    ///   3D ビューを書き出し直し、シーンファイルを置き換える。描画ウィンドウは更新時刻を見て読み直す
    /// - 描画ウィンドウの「同期」ボタンで止めると「.pause」ができるので、その間は書き出さない（変更は覚えておき、再開したときに書き出す）
    /// - 描画ウィンドウが閉じられたら、その分の見張りをやめる。全部閉じたらイベントの登録も外す
    /// </summary>
    internal static class LiveSync
    {
        /// <summary>モデルが変わってから書き出すまで待つ時間（続けて変更されたときに何度も書き出さないため）。</summary>
        private static readonly TimeSpan Debounce = TimeSpan.FromSeconds(0.3);

        private class Session
        {
            public Process Process;
            public Document Doc;
            public ElementId ViewId;
            public string ScenePath;
            public bool Dirty;
            public DateTime LastChange;
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

        public static void Start(UIApplication uiApp, Document doc, View3D view, string scenePath, Process process)
        {
            var session = new Session
            {
                Process = process,
                Doc = doc,
                ViewId = view.Id,
                ScenePath = scenePath,
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
                s.Dirty = true;
                s.LastChange = DateTime.UtcNow;
            }
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
            try
            {
                RenderScene scene = RealtimeRenderCommand.ExportScene(s.Doc, view);
                if (scene.TriangleCount == 0) return;
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
                DiagLog.Cmd("RealtimeRender", $"ライブ同期: モデルを更新 ({scene.TriangleCount} 三角形)");
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"ライブ同期: 書き出しに失敗 {ex.Message}");
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

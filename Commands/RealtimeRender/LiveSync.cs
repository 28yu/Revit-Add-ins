using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace Tools28.Commands.RealtimeRender
{
    /// <summary>
    /// ライブ同期: 描画ウィンドウが開いている間、Revit での変更・視点の移動を描画ウィンドウに伝える。
    ///
    /// - モデルが変わったら（DocumentChanged）、少し待ってから（続けて変更されることが多いので）Revit が手の空いたとき（Idling）に
    ///   3D ビューを書き出し直し、シーンファイルを置き換える。描画ウィンドウは更新時刻を見て読み直す
    /// - Revit のビューの視点が変わったら、シーンファイルの横に「.cam」（カメラの位置と向き）を書く
    /// - 描画ウィンドウの「同期」ボタンで止めると「.pause」ができるので、その間は書き出さない（変更は覚えておき、再開したときに書き出す）
    /// - 描画ウィンドウが閉じられたら、その分の見張りをやめる。全部閉じたらイベントの登録も外す
    /// </summary>
    internal static class LiveSync
    {
        /// <summary>モデルが変わってから書き出すまで待つ時間（続けて変更されたときに何度も書き出さないため）。</summary>
        private static readonly TimeSpan Debounce = TimeSpan.FromSeconds(0.8);

        private class Session
        {
            public Process Process;
            public Document Doc;
            public ElementId ViewId;
            public string ScenePath;
            public bool Dirty;
            public DateTime LastChange;
            public string LastCameraKey;
            public RenderScene Bounds;  // アイソメのビューでカメラを置き直すのに使う（モデル全体を囲む箱だけ）
            public int CameraSeq;
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

        public static void Start(UIApplication uiApp, Document doc, View3D view, string scenePath, Process process, RenderScene scene)
        {
            var session = new Session
            {
                Process = process,
                Doc = doc,
                ViewId = view.Id,
                ScenePath = scenePath,
                Bounds = BoundsOnly(scene),
            };
            session.LastCameraKey = CameraKey(view);
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
                UpdateCamera(s, view);
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
                s.Bounds = BoundsOnly(scene);
                DiagLog.Cmd("RealtimeRender", $"ライブ同期: モデルを更新 ({scene.TriangleCount} 三角形)");
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"ライブ同期: 書き出しに失敗 {ex.Message}");
            }
        }

        private static void UpdateCamera(Session s, View3D view)
        {
            string key;
            try
            {
                key = CameraKey(view);
            }
            catch
            {
                return;
            }
            if (key == s.LastCameraKey) return;
            s.LastCameraKey = key;
            try
            {
                RenderCamera c = RealtimeRenderCommand.BuildCamera(view, s.Bounds, quiet: true);
                s.CameraSeq++;
                var ci = CultureInfo.InvariantCulture;
                string text = string.Join(" ", new[]
                {
                    s.CameraSeq.ToString(ci),
                    F(c.Eye[0]), F(c.Eye[1]), F(c.Eye[2]),
                    F(c.Forward[0]), F(c.Forward[1]), F(c.Forward[2]),
                    F(c.Up[0]), F(c.Up[1]), F(c.Up[2]),
                    F(c.FovY),
                });
                File.WriteAllText(s.ScenePath + ".cam", text + "\n");
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"ライブ同期: 視点の書き出しに失敗 {ex.Message}");
            }
        }

        private static string F(float v) => v.ToString("R", CultureInfo.InvariantCulture);

        /// <summary>ビューの視点を表す文字列（変わったかどうかを比べるため）。</summary>
        private static string CameraKey(View3D view)
        {
            ViewOrientation3D o = view.GetOrientation();
            string crop = "";
            if (view.IsPerspective)
            {
                BoundingBoxXYZ box = view.CropBox;
                crop = $"{box.Min.Y:F4},{box.Max.Y:F4},{box.Max.Z:F4}";
            }
            return $"{o.EyePosition.X:F4},{o.EyePosition.Y:F4},{o.EyePosition.Z:F4}|" +
                   $"{o.ForwardDirection.X:F5},{o.ForwardDirection.Y:F5},{o.ForwardDirection.Z:F5}|" +
                   $"{o.UpDirection.X:F5},{o.UpDirection.Y:F5},{o.UpDirection.Z:F5}|{view.IsPerspective}|{crop}";
        }

        /// <summary>モデル全体を囲む箱の 2 隅だけを持つシーン（カメラの計算用。メモリを使わないように）。</summary>
        private static RenderScene BoundsOnly(RenderScene scene)
        {
            var result = new RenderScene();
            var b = scene.Bounds();
            if (b.HasValue)
            {
                result.Vertices.Add(new RenderVertex { PX = b.Value.min[0], PY = b.Value.min[1], PZ = b.Value.min[2] });
                result.Vertices.Add(new RenderVertex { PX = b.Value.max[0], PY = b.Value.max[1], PZ = b.Value.max[2] });
            }
            return result;
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

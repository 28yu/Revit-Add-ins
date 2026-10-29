using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Localization;

namespace Tools28.Commands.RealtimeRender
{
    /// <summary>
    /// アクティブな 3D ビューの形状・マテリアル・カメラ・太陽の位置を書き出し、
    /// レイトレーシングの描画プログラム（Tools28Render.exe）を別のプログラムとして起動する。
    ///
    /// 描画は Revit の外で行うので、描画側で問題が起きても Revit は止まらない。
    /// モデルは読むだけで変更しない。
    /// </summary>
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class RealtimeRenderCommand : IExternalCommand
    {
        private const string RendererExeName = "Tools28Render.exe";
        private const double FeetToMeters = 0.3048;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application?.ActiveUIDocument;
            Document doc = uidoc?.Document;
            var view = doc?.ActiveView as View3D;
            if (view == null || view.IsTemplate)
            {
                TaskDialog.Show(Loc.S("RtRender.Title"), Loc.S("RtRender.Err.Not3DView"));
                return Result.Cancelled;
            }

            string exePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "", RendererExeName);
            if (!File.Exists(exePath))
            {
                TaskDialog.Show(Loc.S("RtRender.Title"), string.Format(Loc.S("RtRender.Err.NoRenderer"), exePath));
                return Result.Failed;
            }

            try
            {
                DiagLog.Cmd("RealtimeRender", $"開始 view={view.Name}");
                var watch = Stopwatch.StartNew();

                // フェーズの「上書き」表示のマテリアル。まだ選んでいないモデルでは選ぶ画面を出す（Shift を押しながらなら選び直す）
                bool shift = (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) != 0;
                PhaseMaterials.Resolve(doc, view, commandData, true, shift, out bool phaseCancelled);
                if (phaseCancelled) return Result.Cancelled;

                var exportedIds = new HashSet<ElementId>();
                RenderScene scene = ExportScene(doc, view, exportedIds);
                if (scene.TriangleCount == 0)
                {
                    TaskDialog.Show(Loc.S("RtRender.Title"), Loc.S("RtRender.Err.NoGeometry"));
                    return Result.Cancelled;
                }

                // 描画ウィンドウごとに別のファイルにする（ライブ同期で、ウィンドウごとに書き直すため）
                string scenePath = LiveSync.NewScenePath();
                SceneWriter.Write(scenePath, scene);

                var psi = new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = $"--scene \"{scenePath}\" --lang {Loc.CurrentLang} --title \"{MakeTitle(doc, view)}\" --view \"{CleanArg(view.Name)}\" --live",
                    UseShellExecute = false,
                    WorkingDirectory = Path.GetDirectoryName(exePath) ?? "",
                };
                Process process = Process.Start(psi);
                DiagLog.Cmd("RealtimeRender", $"描画プログラム起動 ({watch.ElapsedMilliseconds} ms)");
                if (process != null)
                    LiveSync.Start(commandData.Application, doc, view, scenePath, process, exportedIds);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"例外: {ex}");
                TaskDialog.Show(Loc.S("RtRender.Title"), string.Format(Loc.S("RtRender.Err.Failed"), ex.Message));
                return Result.Failed;
            }
        }

        /// <summary>
        /// 3D ビューに表示されている形状・マテリアルを集め、カメラと太陽の向きを決める（ライブ同期でも使う）。
        /// </summary>
        /// <param name="elementIds">指定すると、書き出した要素の ID を入れる（ライブ同期用）</param>
        internal static RenderScene ExportScene(Document doc, View3D view, HashSet<ElementId> elementIds = null)
        {
            var watch = Stopwatch.StartNew();
            // フェーズの「上書き」表示のマテリアル（選んだ内容を使う。ライブ同期では画面を出さない）
            PhaseOverrides phase = PhaseMaterials.Resolve(doc, view, null, false, false, out _);
            var context = new SceneExportContext(doc, phase, view);
            var exporter = new CustomExporter(doc, context) { IncludeGeometricObjects = false };
            exporter.Export(view);

            RenderScene scene = context.Scene;
            if (elementIds != null)
            {
                elementIds.Clear();
                elementIds.UnionWith(context.ElementIds);
            }
            DiagLog.Cmd("RealtimeRender",
                $"書き出し: 要素 {context.ElementCount} / 三角形 {scene.TriangleCount}（描く数 {scene.DrawnTriangleCount}）/ 形 {scene.Meshes.Count} / 置き場所 {scene.Instances.Count}（使い回し {context.ReusedShapes}・読み飛ばし {context.SkippedInstances}・確認で不一致 {context.VerifyMismatches}）/ マテリアル {scene.Materials.Count} / テクスチャ {scene.Textures.Count} / {watch.ElapsedMilliseconds} ms");
            foreach (string missing in context.MissingTextures)
                DiagLog.Cmd("RealtimeRender", $"テクスチャ画像が見つかりません: {missing}");
            if (scene.TriangleCount > 0)
            {
                scene.Camera = BuildCamera(view, scene);
                ApplySun(doc, view, scene);
                ApplyRenderingSettings(view, scene);
            }
            return scene;
        }

        /// <summary>
        /// Revit のビューと同じ位置・向きのカメラを作る。
        /// アイソメ（平行投影）のビューは視点が遠く離れているので、モデル全体が見える位置に置き直す。
        /// </summary>
        internal static RenderCamera BuildCamera(View3D view, RenderScene scene)
        {
            ViewOrientation3D o = view.GetOrientation();
            var cam = new RenderCamera
            {
                Forward = ToFloats(o.ForwardDirection.Normalize()),
                Up = ToFloats(o.UpDirection.Normalize()),
            };

            if (view.IsPerspective)
            {
                cam.Eye = ToFloats(o.EyePosition, FeetToMeters);
                cam.FovY = PerspectiveFovY(view);
            }
            else
            {
                var bounds = scene.Bounds();
                if (bounds.HasValue)
                {
                    float[] min = bounds.Value.min, max = bounds.Value.max;
                    float cx = (min[0] + max[0]) / 2, cy = (min[1] + max[1]) / 2, cz = (min[2] + max[2]) / 2;
                    float dx = max[0] - min[0], dy = max[1] - min[1], dz = max[2] - min[2];
                    // Enscape と同じく広い画角（横 約 90°）で、モデルが画面いっぱいに入る距離に置く
                    float distance = (float)Math.Sqrt(dx * dx + dy * dy + dz * dz) * 0.55f;
                    cam.Eye = new[]
                    {
                        cx - cam.Forward[0] * distance,
                        cy - cam.Forward[1] * distance,
                        cz - cam.Forward[2] * distance,
                    };
                }
                cam.FovY = 65f;
            }
            DiagLog.Cmd("RealtimeRender",
                $"カメラ: 透視={view.IsPerspective} 位置=({cam.Eye[0]:F2},{cam.Eye[1]:F2},{cam.Eye[2]:F2}) 画角={cam.FovY:F1}");
            return cam;
        }

        /// <summary>透視ビューのトリミング領域から縦の画角を求める（求められなければ 50°）。</summary>
        private static float PerspectiveFovY(View3D view)
        {
            try
            {
                BoundingBoxXYZ crop = view.CropBox;
                double halfHeight = (crop.Max.Y - crop.Min.Y) / 2;
                double distance = Math.Abs(crop.Max.Z);
                if (halfHeight > 0 && distance > 1e-6)
                {
                    double fov = 2 * Math.Atan(halfHeight / distance) * 180 / Math.PI;
                    if (fov >= 10 && fov <= 120) return (float)fov;
                }
            }
            catch
            {
                // 求められないときは既定値を使う
            }
            return 50f;
        }

        /// <summary>
        /// 3D ビューのプロパティの「レンダリング設定」（照明のスキーム・背景）を、描画ウィンドウの最初の見た目にする。
        /// 露出（露出値・彩度など）は描画プログラムの自動露出と考え方が違うので使わない。
        /// </summary>
        private static void ApplyRenderingSettings(View3D view, RenderScene scene)
        {
            try
            {
                RenderingSettings rs = view.GetRenderingSettings();
                switch (rs.LightingSource)
                {
                    case LightingSource.ExteriorSun:
                    case LightingSource.InteriorSun:
                        scene.SunOn = true; scene.ArtificialOn = false; break;
                    case LightingSource.ExteriorArtificial:
                    case LightingSource.InteriorArtificial:
                        scene.SunOn = false; scene.ArtificialOn = true; break;
                    default:
                        scene.SunOn = true; scene.ArtificialOn = true; break;
                }

                BackgroundStyle style = rs.BackgroundStyle;
                switch (style)
                {
                    case BackgroundStyle.SkyNoClouds: scene.CloudAmount = 0f; break;
                    case BackgroundStyle.SkyVeryFewClouds: scene.CloudAmount = 0.25f; break;
                    case BackgroundStyle.SkyFewClouds: scene.CloudAmount = 0.5f; break;
                    case BackgroundStyle.SkyCloudy: scene.CloudAmount = 0.75f; break;
                    case BackgroundStyle.SkyVeryCloudy: scene.CloudAmount = 1f; break;
                    case BackgroundStyle.Transparent: scene.BackgroundMode = 2; break;
                    case BackgroundStyle.Color:
                        BackgroundSettings bg = rs.GetBackgroundSettings();
                        scene.BackgroundMode = 1;
                        if (bg is ColorBackgroundSettings color && color.Color.IsValid)
                            scene.BackgroundColor = ToBgr(color.Color);
                        else if (bg is GradientBackgroundSettings gradient)
                        {
                            // グラデーション: 空の色を背景色に、地面の色を地面の色にする（描画プログラムの背景は 1 色）
                            if (gradient.SkyColor.IsValid) scene.BackgroundColor = ToBgr(gradient.SkyColor);
                            if (gradient.GroundColor.IsValid) scene.GroundColor = 0xFF000000u | ToBgr(gradient.GroundColor);
                        }
                        break;
                    default:
                        // 画像の背景は描画プログラムでは使えないので、空にする
                        break;
                }
                DiagLog.Cmd("RealtimeRender",
                    $"レンダリング設定: 照明={rs.LightingSource} 背景={style} → 太陽={scene.SunOn} 照明器具={scene.ArtificialOn} " +
                    $"背景={scene.BackgroundMode} 色={scene.BackgroundColor:X6} 地面={scene.GroundColor:X8} 雲={scene.CloudAmount}");
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"レンダリング設定を読めませんでした: {ex.Message}");
            }
        }

        private static uint ToBgr(Color c) => (uint)(c.Red | (c.Green << 8) | (c.Blue << 16));

        /// <summary>
        /// ビューの「太陽の設定」から太陽の方向を求める。
        /// Revit の方位は真北が基準なので、プロジェクトの北（モデルの +Y）とのずれを補正する。
        /// </summary>
        private static void ApplySun(Document doc, View3D view, RenderScene scene)
        {
            try
            {
                SunAndShadowSettings sun = view.SunAndShadowSettings;
                if (sun == null) return;
                double altitude = sun.GetFrameAltitude(sun.ActiveFrame);
                double azimuth = sun.GetFrameAzimuth(sun.ActiveFrame);
                double trueNorth = doc.ActiveProjectLocation?.GetProjectPosition(XYZ.Zero)?.Angle ?? 0;

                double a = azimuth + trueNorth;
                scene.SunX = (float)(Math.Sin(a) * Math.Cos(altitude));
                scene.SunY = (float)(Math.Cos(a) * Math.Cos(altitude));
                scene.SunZ = (float)Math.Sin(altitude);
                // 描画プログラムで時間帯を変えるとき（太陽の通り道）に使う
                scene.TrueNorth = (float)trueNorth;
                try { scene.Latitude = (float)(doc.SiteLocation.Latitude * 180 / Math.PI); } catch (Exception) { }
                try { scene.DayOfYear = sun.GetFrameTime(sun.ActiveFrame).DayOfYear; } catch (Exception) { }
                DiagLog.Cmd("RealtimeRender",
                    $"太陽: 方位={azimuth * 180 / Math.PI:F1}° 高度={altitude * 180 / Math.PI:F1}° 真北補正={trueNorth * 180 / Math.PI:F1}°");
            }
            catch (Exception ex)
            {
                // 太陽の設定が読めないビューもある。そのときは既定の向き（南東・高さ約40°）を使う
                DiagLog.Cmd("RealtimeRender", $"太陽の設定を読めませんでした: {ex.Message}");
            }
        }

        /// <summary>
        /// 描画ウィンドウの題名・書き出すファイルの名前に使う「モデル名 - ビュー名」。
        /// コマンドラインの引数に入れるので、引用符は使わない文字に置き換える。
        /// </summary>
        private static string MakeTitle(Document doc, View view)
        {
            return CleanArg($"{doc.Title} - {view.Name}");
        }

        /// <summary>コマンドラインの引数に入れる文字列（引用符・\ は使わない文字に置き換える）。</summary>
        private static string CleanArg(string s) => (s ?? "").Replace('"', '\'').Replace("\\", "_").Trim();

                private static float[] ToFloats(XYZ v, double scale = 1.0) =>
            new[] { (float)(v.X * scale), (float)(v.Y * scale), (float)(v.Z * scale) };
    }
}

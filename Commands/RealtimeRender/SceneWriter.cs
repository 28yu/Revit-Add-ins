using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Tools28.Commands.RealtimeRender
{
    /// <summary>
    /// 描画プログラム（Tools28Render.exe）用のシーンファイル（.t28s）を書く。
    /// 形式は Renderer/src/scene.cpp の説明と同じ（変えるときは両方を直すこと）。
    /// </summary>
    internal static class SceneWriter
    {
        // 4: 太陽の後ろに 緯度・真北・日付を追加（描画プログラムで時間帯を変えるため）
        // 5: その後ろに 3D ビューのレンダリング設定（照明・背景・地面の色・雲の量）を追加
        // 6: 最後に 形と置き場所（同じファミリの同じ形は 1 回だけ持ち、位置・向きだけを何度も書く）
        // 7: 最後に 置き場所・光源ごとの要素 ID と、ライブ同期の通し番号（変わった要素の分だけを送るため）
        private const uint Version = 7;

        public static void Write(string path, RenderScene scene)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (var w = new BinaryWriter(stream))  // BinaryWriter はリトルエンディアンで書く
                WriteScene(w, scene);
        }

        /// <summary>
        /// ライブ同期の差分（.d＋通し番号 8 桁）を書く。removed の要素の分を消してから、scene（変わった要素だけを書き出したもの）を足す。
        /// 形式: "T28D" | uint32 版（1）| uint32 通し番号 | uint32 消す要素の数 | int64 要素 ID… | シーン（.t28s と同じ形式）
        /// </summary>
        public static void WriteDelta(string path, RenderScene scene, uint seq, ICollection<long> removed)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (var w = new BinaryWriter(stream))
            {
                w.Write(Encoding.ASCII.GetBytes("T28D"));
                w.Write(1u);
                w.Write(seq);
                w.Write((uint)removed.Count);
                foreach (long id in removed) w.Write(id);
                WriteScene(w, scene);
            }
        }

        private static void WriteScene(BinaryWriter w, RenderScene scene)
        {
            {
                w.Write(Encoding.ASCII.GetBytes("T28S"));
                w.Write(Version);

                RenderCamera c = scene.Camera;
                foreach (float v in c.Eye) w.Write(v);
                foreach (float v in c.Forward) w.Write(v);
                foreach (float v in c.Up) w.Write(v);
                w.Write(c.FovY);

                w.Write(scene.SunX);
                w.Write(scene.SunY);
                w.Write(scene.SunZ);
                w.Write(scene.SunIntensity);
                w.Write(scene.SkyIntensity);
                w.Write(scene.Latitude);
                w.Write(scene.TrueNorth);
                w.Write(scene.DayOfYear);
                w.Write((uint)((scene.SunOn ? 1 : 0) | (scene.ArtificialOn ? 2 : 0)));
                w.Write(scene.BackgroundMode);
                w.Write(scene.BackgroundColor);
                w.Write(scene.GroundColor);
                w.Write(scene.CloudAmount);

                w.Write((uint)scene.Materials.Count);
                foreach (RenderMaterial m in scene.Materials)
                {
                    w.Write(m.R);
                    w.Write(m.G);
                    w.Write(m.B);
                    w.Write(m.Opacity);
                    w.Write(m.Roughness);
                    w.Write(m.Metallic);
                    w.Write(m.Emissive);
                    w.Write(m.F0);
                    w.Write(m.ColorTexture);
                    foreach (float v in m.ColorXform) w.Write(v);
                    w.Write(m.BumpTexture);
                    foreach (float v in m.BumpXform) w.Write(v);
                    w.Write(m.BumpType);
                    w.Write(m.BumpDepth);
                    byte[] name = Encoding.UTF8.GetBytes(m.Name ?? "");
                    w.Write((uint)name.Length);
                    w.Write(name);
                }

                w.Write((uint)scene.Vertices.Count);
                foreach (RenderVertex v in scene.Vertices)
                {
                    w.Write(v.PX); w.Write(v.PY); w.Write(v.PZ);
                    w.Write(v.NX); w.Write(v.NY); w.Write(v.NZ);
                    w.Write(v.U); w.Write(v.V);
                }

                w.Write((uint)scene.TriangleCount);
                foreach (int i in scene.Triangles) w.Write((uint)i);

                w.Write((uint)scene.Textures.Count);
                foreach (RenderTexture t in scene.Textures)
                {
                    w.Write((uint)t.Width);
                    w.Write((uint)t.Height);
                    w.Write(t.Linear ? 1u : 0u);
                    w.Write(t.Rgba);
                }

                w.Write((uint)scene.Lights.Count);
                foreach (RenderLight l in scene.Lights)
                {
                    foreach (float v in l.Position) w.Write(v);
                    w.Write(l.Radius);
                    foreach (float v in l.Intensity) w.Write(v);
                    w.Write((float)l.Type);
                    foreach (float v in l.Direction) w.Write(v);
                    w.Write(l.CosInner);
                    w.Write(l.CosOuter);
                }

                // 形と置き場所（形が 0 なら、全部の三角形がそのままの位置）
                w.Write((uint)scene.Meshes.Count);
                foreach (RenderMesh m in scene.Meshes)
                {
                    w.Write((uint)m.FirstTri);
                    w.Write((uint)m.TriCount);
                }
                w.Write((uint)(scene.Meshes.Count > 0 ? scene.Instances.Count : 0));
                if (scene.Meshes.Count > 0)
                    foreach (RenderInstance inst in scene.Instances)
                    {
                        w.Write((uint)inst.Mesh);
                        foreach (float v in inst.M) w.Write(v);
                    }

                // 版 7: 置き場所ごと・光源ごとの要素 ID と、ライブ同期の通し番号
                if (scene.Meshes.Count > 0)
                    foreach (RenderInstance inst in scene.Instances) w.Write(inst.Element);
                foreach (RenderLight l in scene.Lights) w.Write(l.Element);
                w.Write(scene.SyncSeq);
            }
        }
    }
}

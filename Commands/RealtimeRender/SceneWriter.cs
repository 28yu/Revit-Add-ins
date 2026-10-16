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
        private const uint Version = 1;

        public static void Write(string path, RenderScene scene)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (var w = new BinaryWriter(stream))  // BinaryWriter はリトルエンディアンで書く
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
                    byte[] name = Encoding.UTF8.GetBytes(m.Name ?? "");
                    w.Write((uint)name.Length);
                    w.Write(name);
                }

                w.Write((uint)scene.Vertices.Count);
                foreach (RenderVertex v in scene.Vertices)
                {
                    w.Write(v.PX); w.Write(v.PY); w.Write(v.PZ);
                    w.Write(v.NX); w.Write(v.NY); w.Write(v.NZ);
                }

                w.Write((uint)scene.TriangleCount);
                foreach (int i in scene.Triangles) w.Write((uint)i);
            }
        }
    }
}

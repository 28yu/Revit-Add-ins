using System;
using System.Collections.Generic;

namespace Tools28.Commands.RealtimeRender
{
    /// <summary>描画プログラムに渡すモデルのデータ。座標はメートル、Z が上。</summary>
    internal class RenderScene
    {
        public List<RenderVertex> Vertices { get; } = new List<RenderVertex>();
        /// <summary>三角形（頂点番号×3 とマテリアル番号）を並べたもの。</summary>
        public List<int> Triangles { get; } = new List<int>();
        public List<RenderMaterial> Materials { get; } = new List<RenderMaterial>();
        /// <summary>テクスチャ画像（マテリアルの外観の「イメージ」）。</summary>
        public List<RenderTexture> Textures { get; } = new List<RenderTexture>();
        /// <summary>照明器具の光源。</summary>
        public List<RenderLight> Lights { get; } = new List<RenderLight>();

        public RenderCamera Camera { get; set; } = new RenderCamera();
        /// <summary>太陽のある方向（単位ベクトル）。</summary>
        public float SunX { get; set; } = 0.5f;
        public float SunY { get; set; } = -0.6f;
        public float SunZ { get; set; } = 0.62f;
        public float SunIntensity { get; set; } = 5f;
        public float SkyIntensity { get; set; } = 1f;

        public int TriangleCount => Triangles.Count / 4;

        public void AddTriangle(int a, int b, int c, int material)
        {
            Triangles.Add(a);
            Triangles.Add(b);
            Triangles.Add(c);
            Triangles.Add(material);
        }

        /// <summary>全頂点を囲む箱（メートル）。頂点がなければ null。</summary>
        public (float[] min, float[] max)? Bounds()
        {
            if (Vertices.Count == 0) return null;
            var min = new[] { float.MaxValue, float.MaxValue, float.MaxValue };
            var max = new[] { float.MinValue, float.MinValue, float.MinValue };
            foreach (var v in Vertices)
            {
                min[0] = Math.Min(min[0], v.PX); max[0] = Math.Max(max[0], v.PX);
                min[1] = Math.Min(min[1], v.PY); max[1] = Math.Max(max[1], v.PY);
                min[2] = Math.Min(min[2], v.PZ); max[2] = Math.Max(max[2], v.PZ);
            }
            return (min, max);
        }
    }

    internal struct RenderVertex
    {
        public float PX, PY, PZ;
        public float NX, NY, NZ;
        /// <summary>面の上の位置（メートル）。画像上の位置へは、マテリアルの貼り方（Xform）で直す。</summary>
        public float U, V;
    }

    internal class RenderMaterial
    {
        public string Name = "";
        public float R = 0.8f, G = 0.8f, B = 0.8f;  // リニアの色
        public float Opacity = 1f;
        public float Roughness = 0.5f;
        public float Metallic;
        public float Emissive = 0f;  // 自己発光（照明器具の対応は今後）
        public float F0 = 0.04f;     // 正面から見たときの反射率（外観の「反射率」）
        public int ColorTexture = -1;  // 色の画像の番号（-1 = 無し）
        public float[] ColorXform = BitmapInfo.Identity;
        public int BumpTexture = -1;   // 凹凸の画像の番号（-1 = 無し）
        public float[] BumpXform = BitmapInfo.Identity;
        public int BumpType;           // 0 = 高さの画像, 1 = 法線の画像
        public float BumpDepth;        // 高さの画像: 白と黒の差（メートル）、法線の画像: 強さ
        /// <summary>画像を使うか（使うときだけ面の上の位置 U, V を求める）。</summary>
        public bool UsesUv => ColorTexture >= 0 || BumpTexture >= 0;
    }

    /// <summary>テクスチャ画像。1 画素 4 バイト（R, G, B, A）。色の画像は sRGB、凹凸の画像はそのままの値（Linear）。</summary>
    internal class RenderTexture
    {
        public int Width, Height;
        public bool Linear;
        public byte[] Rgba;
    }

    /// <summary>照明器具の光源。明るさは「光度（カンデラ）÷ 20000」（描画プログラムの単位。太陽 5 = 10 万ルクス）。</summary>
    internal class RenderLight
    {
        public float[] Position = new float[3];
        public float Radius = 0.05f;
        public float[] Intensity = new float[3];
        public int Type;  // 0 = 全方向, 1 = 半球, 2 = スポット
        public float[] Direction = { 0, 0, -1 };
        public float CosInner = 1f, CosOuter;
    }

    internal class RenderCamera
    {
        public float[] Eye = { 0, -5, 1.6f };
        public float[] Forward = { 0, 1, 0 };
        public float[] Up = { 0, 0, 1 };
        public float FovY = 60f;
    }

    /// <summary>マテリアル名から、ガラス・金属らしさを推測する（Revit の色・透明度だけでは見分けにくいため）。</summary>
    internal static class MaterialHints
    {
        private static readonly string[] GlassWords = { "glass", "glazing", "ガラス", "硝子", "玻璃" };
        private static readonly string[] MetalWords =
        {
            "metal", "steel", "aluminum", "aluminium", "stainless", "chrome", "copper", "brass",
            "金属", "スチール", "アルミ", "ステンレス", "鋼", "鉄", "銅", "钢", "铝", "铁",
        };

        public static bool LooksLikeGlass(string name) => ContainsAny(name, GlassWords);
        public static bool LooksMetallic(string name) => ContainsAny(name, MetalWords);

        private static bool ContainsAny(string name, string[] words)
        {
            if (string.IsNullOrEmpty(name)) return false;
            string lower = name.ToLowerInvariant();
            foreach (string w in words)
                if (lower.Contains(w)) return true;
            return false;
        }
    }
}

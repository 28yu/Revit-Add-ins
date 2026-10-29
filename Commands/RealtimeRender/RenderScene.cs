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
        /// <summary>
        /// 形（三角形の並びのうち連続した一続き）。同じファミリの同じ形は 1 回だけ持ち、Instances で何度も置く
        /// （GPU のメモリ・ファイルの大きさを抑えるため）。頂点は形ごとの座標（ファミリから見た座標。メートル）。
        /// </summary>
        public List<RenderMesh> Meshes { get; } = new List<RenderMesh>();
        /// <summary>置き場所: どの形を、どの位置・向きに置くか。</summary>
        public List<RenderInstance> Instances { get; } = new List<RenderInstance>();

        public RenderCamera Camera { get; set; } = new RenderCamera();
        /// <summary>太陽のある方向（単位ベクトル）。</summary>
        public float SunX { get; set; } = 0.5f;
        public float SunY { get; set; } = -0.6f;
        public float SunZ { get; set; } = 0.62f;
        public float SunIntensity { get; set; } = 5f;
        public float SkyIntensity { get; set; } = 1f;
        // 描画プログラムで時間帯を変えるときに使う敷地の情報（緯度=度、真北=ラジアン、日付=1月1日が1）
        public float Latitude { get; set; } = 35.7f;
        public float TrueNorth { get; set; } = 0f;
        public float DayOfYear { get; set; } = 80f;
        // 3D ビューの「レンダリング設定」（描画ウィンドウの最初の見た目。ウィンドウで変えればそちらが優先）
        public bool SunOn { get; set; } = true;          // 照明のスキームで太陽を使う
        public bool ArtificialOn { get; set; } = true;   // 照明のスキームで人工照明を使う
        public uint BackgroundMode { get; set; }         // 0 = 空, 1 = 色（グラデーションも）, 2 = 透明
        public uint BackgroundColor { get; set; } = 0xFFFFFF;  // 0xBBGGRR
        public uint GroundColor { get; set; }            // 0xFFBBGGRR（グラデーションの地面の色。0 なら指定なし）
        public float CloudAmount { get; set; } = 0.5f;   // 0 = 雲なし 〜 1 = 曇り

        /// <summary>持っている三角形の数（同じ形は 1 回だけ数える）。</summary>
        public int TriangleCount => Triangles.Count / 4;

        /// <summary>置き場所の分も数えた三角形の数（画面に描かれる数）。</summary>
        public long DrawnTriangleCount
        {
            get
            {
                if (Meshes.Count == 0) return TriangleCount;
                long n = 0;
                foreach (RenderInstance inst in Instances) n += Meshes[inst.Mesh].TriCount;
                return n;
            }
        }

        public void AddTriangle(int a, int b, int c, int material)
        {
            Triangles.Add(a);
            Triangles.Add(b);
            Triangles.Add(c);
            Triangles.Add(material);
        }

        /// <summary>モデル全体を囲む箱（メートル。置き場所の位置まで含める）。頂点がなければ null。</summary>
        public (float[] min, float[] max)? Bounds()
        {
            if (Vertices.Count == 0) return null;
            var min = new[] { float.MaxValue, float.MaxValue, float.MaxValue };
            var max = new[] { float.MinValue, float.MinValue, float.MinValue };
            void Grow(float x, float y, float z)
            {
                min[0] = Math.Min(min[0], x); max[0] = Math.Max(max[0], x);
                min[1] = Math.Min(min[1], y); max[1] = Math.Max(max[1], y);
                min[2] = Math.Min(min[2], z); max[2] = Math.Max(max[2], z);
            }
            if (Meshes.Count == 0)
            {
                foreach (var v in Vertices) Grow(v.PX, v.PY, v.PZ);
                return (min, max);
            }
            // 形ごとの箱（形の座標）を、置き場所ごとに 8 つの角を動かして囲む
            var meshMin = new float[Meshes.Count][];
            var meshMax = new float[Meshes.Count][];
            for (int mi = 0; mi < Meshes.Count; mi++)
            {
                var lo = new[] { float.MaxValue, float.MaxValue, float.MaxValue };
                var hi = new[] { float.MinValue, float.MinValue, float.MinValue };
                RenderMesh m = Meshes[mi];
                for (int t = m.FirstTri; t < m.FirstTri + m.TriCount; t++)
                    for (int k = 0; k < 3; k++)
                    {
                        RenderVertex v = Vertices[Triangles[t * 4 + k]];
                        lo[0] = Math.Min(lo[0], v.PX); hi[0] = Math.Max(hi[0], v.PX);
                        lo[1] = Math.Min(lo[1], v.PY); hi[1] = Math.Max(hi[1], v.PY);
                        lo[2] = Math.Min(lo[2], v.PZ); hi[2] = Math.Max(hi[2], v.PZ);
                    }
                meshMin[mi] = lo;
                meshMax[mi] = hi;
            }
            foreach (RenderInstance inst in Instances)
            {
                float[] lo = meshMin[inst.Mesh], hi = meshMax[inst.Mesh], m = inst.M;
                if (lo[0] > hi[0]) continue;
                for (int c = 0; c < 8; c++)
                {
                    float x = (c & 1) != 0 ? hi[0] : lo[0], y = (c & 2) != 0 ? hi[1] : lo[1], z = (c & 4) != 0 ? hi[2] : lo[2];
                    Grow(m[0] * x + m[1] * y + m[2] * z + m[3], m[4] * x + m[5] * y + m[6] * z + m[7],
                         m[8] * x + m[9] * y + m[10] * z + m[11]);
                }
            }
            return min[0] <= max[0] ? (min, max) : ((float[], float[])?)null;
        }
    }

    /// <summary>形: 三角形の並びのうち、FirstTri から TriCount 個。</summary>
    internal struct RenderMesh
    {
        public int FirstTri, TriCount;
    }

    /// <summary>置き場所: 形の番号と、3 行 4 列の行列（行ごと。世界の座標 = 左 3 列 × 形の座標 + 右 1 列。メートル）。</summary>
    internal struct RenderInstance
    {
        public int Mesh;
        public float[] M;
    }

    internal struct RenderVertex : IEquatable<RenderVertex>
    {
        public float PX, PY, PZ;
        public float NX, NY, NZ;
        /// <summary>面の上の位置（メートル）。画像上の位置へは、マテリアルの貼り方（Xform）で直す。</summary>
        public float U, V;

        // 同じ形かを比べるのに使う（既定の比べ方は遅いので、値を直接比べる）
        public bool Equals(RenderVertex o) =>
            PX == o.PX && PY == o.PY && PZ == o.PZ && NX == o.NX && NY == o.NY && NZ == o.NZ && U == o.U && V == o.V;
        public override bool Equals(object obj) => obj is RenderVertex o && Equals(o);
        public override int GetHashCode() => PX.GetHashCode() ^ (PY.GetHashCode() * 31) ^ (PZ.GetHashCode() * 17);
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

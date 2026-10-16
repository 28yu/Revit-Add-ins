using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Visual;

namespace Tools28.Commands.RealtimeRender
{
    /// <summary>
    /// 3D ビューに表示されている形状を三角形の集まりとして集める（Revit の CustomExporter から呼ばれる）。
    /// ファミリ・リンクの位置（Transform）を積み重ねて、すべてプロジェクト座標（メートル）にそろえる。
    /// </summary>
    internal class SceneExportContext : IExportContext
    {
        private const double FeetToMeters = 0.3048;

        private readonly Stack<Transform> _transforms = new Stack<Transform>();
        private readonly Stack<Document> _documents = new Stack<Document>();
        private readonly Dictionary<string, int> _materialIndex = new Dictionary<string, int>();
        private readonly Dictionary<string, int> _textureIndex = new Dictionary<string, int>();
        private readonly TextureLoader _textures;

        private int _currentMaterial;

        public RenderScene Scene { get; } = new RenderScene();
        public int ElementCount { get; private set; }
        /// <summary>見つからなかったテクスチャ画像（ログ用）。</summary>
        public IReadOnlyList<string> MissingTextures => _textures.Missing;

        public SceneExportContext(Document doc)
        {
            _documents.Push(doc);
            _transforms.Push(Transform.Identity);
            _textures = new TextureLoader(doc.PathName, doc.Application.VersionNumber);
        }

        public bool Start()
        {
            // マテリアルが一度も指定されない形状のための既定の色
            Scene.Materials.Add(new RenderMaterial { Name = "Default", R = 0.7f, G = 0.7f, B = 0.7f });
            return true;
        }

        public void Finish() { }

        public bool IsCanceled() => false;

        public RenderNodeAction OnViewBegin(ViewNode node)
        {
            // 曲面をどれだけ細かい三角形にするか（0〜15）。大きいほど滑らかだが重くなる
            node.LevelOfDetail = 8;
            return RenderNodeAction.Proceed;
        }

        public void OnViewEnd(ElementId elementId) { }

        public RenderNodeAction OnElementBegin(ElementId elementId)
        {
            ElementCount++;
            return RenderNodeAction.Proceed;
        }

        public void OnElementEnd(ElementId elementId) { }

        public RenderNodeAction OnInstanceBegin(InstanceNode node)
        {
            _transforms.Push(_transforms.Peek().Multiply(node.GetTransform()));
            return RenderNodeAction.Proceed;
        }

        public void OnInstanceEnd(InstanceNode node) => _transforms.Pop();

        public RenderNodeAction OnLinkBegin(LinkNode node)
        {
            _transforms.Push(_transforms.Peek().Multiply(node.GetTransform()));
            _documents.Push(node.GetDocument());
            return RenderNodeAction.Proceed;
        }

        public void OnLinkEnd(LinkNode node)
        {
            _transforms.Pop();
            _documents.Pop();
        }

        public RenderNodeAction OnFaceBegin(FaceNode node) => RenderNodeAction.Proceed;

        public void OnFaceEnd(FaceNode node) { }

        public void OnRPC(RPCNode node) { }

        public void OnLight(LightNode node) { }

        public void OnMaterial(MaterialNode node)
        {
            Document doc = _documents.Peek();
            Color color = node.Color;
            bool hasColor = color != null && color.IsValid;
            byte r = hasColor ? color.Red : (byte)180, g = hasColor ? color.Green : (byte)180, b = hasColor ? color.Blue : (byte)180;

            // 同じマテリアルでも、ビューの上書きで色・透明度が変わることがあるので、それも含めて区別する
            string key = $"{doc?.PathName}|{node.MaterialId?.ToString()}|{r},{g},{b}|{node.Transparency:F3}|{node.Smoothness}|{node.HasOverriddenAppearance}";
            if (_materialIndex.TryGetValue(key, out int index))
            {
                _currentMaterial = index;
                return;
            }

            string name = "";
            if (doc != null && node.MaterialId != null && node.MaterialId != ElementId.InvalidElementId)
                name = (doc.GetElement(node.MaterialId) as Material)?.Name ?? "";

            // マテリアルの「外観」タブの内容（色・イメージ・反射率・粗さ など）
            AppearanceInfo appearance = null;
            try
            {
                Asset asset = node.HasOverriddenAppearance ? node.GetAppearanceOverride() : node.GetAppearance();
                appearance = AppearanceReader.Read(asset);
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"外観を読めませんでした: {name} ({ex.Message})");
            }

            // 色は外観の色を優先する（Revit の画面表示用の色は外観と違うことがある）
            double[] srgb = appearance?.Color ?? new[] { r / 255.0, g / 255.0, b / 255.0 };
            var material = new RenderMaterial
            {
                Name = name,
                // 色は画面表示用（sRGB）なので、光の計算用（リニア）に直す
                R = (float)SrgbToLinear(srgb[0]),
                G = (float)SrgbToLinear(srgb[1]),
                B = (float)SrgbToLinear(srgb[2]),
                Opacity = (float)Clamp(1.0 - node.Transparency, 0, 1),
                // Revit の「滑らかさ」0〜100 → 粗さ 1〜0（外観に粗さがあればそちらを使う）
                Roughness = (float)Clamp(1.0 - node.Smoothness / 100.0, 0.05, 1),
                Metallic = MaterialHints.LooksMetallic(name) ? 1f : 0f,
            };

            if (appearance != null)
            {
                if (appearance.ColorImage != null)
                {
                    material.TextureIndex = GetTexture(appearance.ColorImage, srgb, appearance.ImageFade);
                    if (material.TextureIndex >= 0)
                    {
                        // 色は画像に焼き込んだので、掛け合わせる色は白にする
                        material.R = material.G = material.B = 1f;
                        material.Mapping = appearance.ColorImage;
                    }
                }

                double? roughness = appearance.Roughness;
                if (appearance.RoughnessImage != null)
                {
                    string path = _textures.Resolve(appearance.RoughnessImage.Paths);
                    double? mean = path != null ? TextureLoader.AverageLuminance(path) : null;
                    if (mean.HasValue) roughness = appearance.RoughnessImage.Invert ? 1.0 - mean.Value : mean.Value;
                }
                if (roughness.HasValue) material.Roughness = (float)Clamp(roughness.Value, 0.05, 1);
                if (appearance.F0.HasValue) material.F0 = (float)Clamp(appearance.F0.Value, 0, 1);
                if (appearance.Transparency.HasValue)
                    material.Opacity = Math.Min(material.Opacity, (float)Clamp(1.0 - appearance.Transparency.Value, 0, 1));
                if (appearance.Metal) material.Metallic = 1f;
                if (appearance.Mirror)
                {
                    material.Metallic = 1f;
                    material.Roughness = 0.05f;
                }
                if (appearance.Glass && material.Opacity > 0.5f) material.Opacity = 0.3f;
            }
            if (MaterialHints.LooksLikeGlass(name) && material.Opacity > 0.5f)
                material.Opacity = 0.3f;

            index = Scene.Materials.Count;
            Scene.Materials.Add(material);
            _materialIndex[key] = index;
            _currentMaterial = index;
            DiagLog.Cmd("RealtimeRender",
                $"マテリアル: {name} 色={BitmapInfo.FormatColor(srgb)} 画像={(material.TextureIndex >= 0 ? appearance?.ColorImage?.Paths : "-")} " +
                $"粗さ={material.Roughness:F2} 反射率={material.F0:F2} 不透明度={material.Opacity:F2} 金属={material.Metallic}");
        }

        /// <summary>
        /// テクスチャ画像を読み込み、明るさ・色付け・「イメージ フェード」（色との混ぜ具合）を焼き込んで登録する。
        /// 戻り値はテクスチャ番号（読めなければ -1）。
        /// </summary>
        private int GetTexture(BitmapInfo bitmap, double[] baseSrgb, double fade)
        {
            string key = $"{bitmap.BakeKey}|{fade:F3}|{BitmapInfo.FormatColor(fade < 1 ? baseSrgb : null)}";
            if (_textureIndex.TryGetValue(key, out int index)) return index;
            index = -1;
            string path = _textures.Resolve(bitmap.Paths);
            // テクスチャが多いときは小さめにして、GPU のメモリを使いすぎないようにする
            RenderTexture texture = path != null ? TextureLoader.Decode(path, Scene.Textures.Count < 32 ? 1024 : 512) : null;
            if (texture != null)
            {
                Bake(texture, bitmap, baseSrgb, fade);
                index = Scene.Textures.Count;
                Scene.Textures.Add(texture);
            }
            _textureIndex[key] = index;
            return index;
        }

        private static void Bake(RenderTexture texture, BitmapInfo bitmap, double[] baseSrgb, double fade)
        {
            bool tint = bitmap.Tint != null;
            if (Math.Abs(bitmap.Brightness - 1.0) < 1e-3 && !bitmap.Invert && !tint && fade >= 0.999) return;
            var lut = new double[256];
            for (int i = 0; i < 256; i++) lut[i] = SrgbToLinear(i / 255.0);
            byte[] p = texture.Rgba;
            for (int i = 0; i < p.Length; i += 4)
            {
                for (int c = 0; c < 3; c++)
                {
                    byte s = bitmap.Invert ? (byte)(255 - p[i + c]) : p[i + c];
                    double lin = lut[s] * bitmap.Brightness;
                    if (tint) lin *= SrgbToLinear(bitmap.Tint[c]);
                    if (fade < 0.999) lin = SrgbToLinear(baseSrgb[c]) * (1 - fade) + lin * fade;
                    p[i + c] = LinearToSrgbByte(lin);
                }
            }
        }

        public void OnPolymesh(PolymeshTopology mesh)
        {
            Transform t = _transforms.Peek();
            IList<XYZ> points = mesh.GetPoints();
            IList<XYZ> normals = mesh.GetNormals();
            IList<PolymeshFacet> facets = mesh.GetFacets();
            if (points.Count == 0 || facets.Count == 0) return;
            // テクスチャのあるマテリアルだけ、面上の位置（UV）からテクスチャ画像上の位置を求める
            BitmapInfo mapping = Scene.Materials[_currentMaterial].Mapping;
            IList<UV> uvs = mapping != null ? mesh.GetUVs() : null;
            RenderVertex Vertex(int i, XYZ n) =>
                MakeVertex(t, points[i], n, mapping, uvs != null && i < uvs.Count ? uvs[i] : null);

            switch (mesh.DistributionOfNormals)
            {
                case DistributionOfNormals.AtEachPoint:
                {
                    int baseIndex = Scene.Vertices.Count;
                    for (int i = 0; i < points.Count; i++)
                        Scene.Vertices.Add(Vertex(i, i < normals.Count ? normals[i] : null));
                    foreach (PolymeshFacet f in facets)
                        Scene.AddTriangle(baseIndex + f.V1, baseIndex + f.V2, baseIndex + f.V3, _currentMaterial);
                    break;
                }
                case DistributionOfNormals.OnePerFace:
                {
                    XYZ n = normals.Count > 0 ? normals[0] : null;
                    int baseIndex = Scene.Vertices.Count;
                    for (int i = 0; i < points.Count; i++)
                        Scene.Vertices.Add(Vertex(i, n));
                    foreach (PolymeshFacet f in facets)
                        Scene.AddTriangle(baseIndex + f.V1, baseIndex + f.V2, baseIndex + f.V3, _currentMaterial);
                    break;
                }
                default: // OnEachFacet: 三角形ごとに向きが違うので、頂点を三角形ごとに分ける
                {
                    for (int i = 0; i < facets.Count; i++)
                    {
                        PolymeshFacet f = facets[i];
                        XYZ n = i < normals.Count ? normals[i] : null;
                        int baseIndex = Scene.Vertices.Count;
                        Scene.Vertices.Add(Vertex(f.V1, n));
                        Scene.Vertices.Add(Vertex(f.V2, n));
                        Scene.Vertices.Add(Vertex(f.V3, n));
                        Scene.AddTriangle(baseIndex, baseIndex + 1, baseIndex + 2, _currentMaterial);
                    }
                    break;
                }
            }
        }

        private static RenderVertex MakeVertex(Transform t, XYZ point, XYZ normal, BitmapInfo mapping, UV uv)
        {
            XYZ p = t.OfPoint(point);
            var v = new RenderVertex
            {
                PX = (float)(p.X * FeetToMeters),
                PY = (float)(p.Y * FeetToMeters),
                PZ = (float)(p.Z * FeetToMeters),
            };
            if (normal != null)
            {
                XYZ n = t.OfVector(normal);
                double len = n.GetLength();
                if (len > 1e-9)
                {
                    v.NX = (float)(n.X / len);
                    v.NY = (float)(n.Y / len);
                    v.NZ = (float)(n.Z / len);
                }
            }
            if (mapping != null && uv != null) mapping.Map(uv.U, uv.V, out v.U, out v.V);
            return v;
        }

        private static double SrgbToLinear(double s) =>
            s <= 0.04045 ? s / 12.92 : Math.Pow((s + 0.055) / 1.055, 2.4);

        private static byte LinearToSrgbByte(double l)
        {
            l = Clamp(l, 0, 1);
            double s = l <= 0.0031308 ? l * 12.92 : 1.055 * Math.Pow(l, 1 / 2.4) - 0.055;
            return (byte)Math.Round(s * 255);
        }

        private static double Clamp(double v, double min, double max) => v < min ? min : v > max ? max : v;
    }
}

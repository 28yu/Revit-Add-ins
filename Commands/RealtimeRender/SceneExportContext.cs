using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Lighting;
using Autodesk.Revit.DB.Visual;

namespace Tools28.Commands.RealtimeRender
{
    /// <summary>
    /// 3D ビューに表示されている形状を三角形の集まりとして集める（Revit の CustomExporter から呼ばれる）。
    /// ファミリ・リンクの位置（Transform）を積み重ねて、プロジェクト座標（メートル）にそろえる。
    /// ファミリの置き場所（InstanceNode）の中の形は、ファミリから見た座標のまま集めて中身を比べ、
    /// 同じ形がすでにあれば形は捨てて「置き場所（位置・向き）」だけを足す（同じ柱・車・照明を何度も持たない）。
    /// それ以外（壁・床など）は 1 つの形にまとめ、そのままの位置に 1 回置く。
    /// Revit 2023 以降は、ファミリの形の番号（SymbolGeometryId）が同じ置き場所の 2 回目からは中身を読まない
    /// （1 回目に覚えた「形と位置」を並べるだけ。書き出しが速くなる）。見た目が変わりうる条件（フェーズの上書き・
    /// ビューのフィルタ・要素ごとの上書き）も見分けの目印に含め、光源を含むファミリは毎回読む。
    /// 念のため 16 回に 1 回は実際に読んで比べ、違っていたらそのファミリは以後も毎回読む。
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
        private ElementId _currentElement = ElementId.InvalidElementId;
        // フェーズの「グラフィックの上書き」: 今の要素を描くマテリアル（上書きしないなら InvalidElementId）
        private readonly PhaseOverrides _phase;
        private ElementId _phaseMaterial = ElementId.InvalidElementId;

        public RenderScene Scene { get; } = new RenderScene();

        /// <summary>集めている途中の形（置き場所 1 つ分、または置き場所の外の全部）。頂点番号はこの中での番号。</summary>
        private class ShapeBuilder
        {
            public Transform World;         // 形の座標 → プロジェクト座標（フィート）。置き場所の外なら単位行列
            public Transform WorldInverse;
            public readonly List<RenderVertex> Vertices = new List<RenderVertex>();
            public readonly List<int> Triangles = new List<int>();  // 頂点番号×3 とマテリアル番号
            // 2 回目から読まないための記録（Revit 2023 以降）
            public string TemplateKey;              // 見分けの目印（null = 記録しない）
            public List<TemplateEntry> Recording;   // この置き場所の中で置いた形と、この置き場所から見た位置
            public ShapeTemplate Skipped;           // 読まずに済ませた（覚えた形を並べる）
            public ShapeTemplate Verify;            // 抜き取り確認（読んだ結果を覚えた形と比べる）
            public bool HasLight;                   // 中に光源がある（覚えない。光源は置き場所ごとに読む必要がある）
        }
        private struct TemplateEntry
        {
            public int Mesh;
            public Transform Relative;  // 置き場所の座標 → 形の座標の位置（フィート）
        }
        private class ShapeTemplate
        {
            public List<TemplateEntry> Entries;
            public int Uses;
            public bool Unsafe;  // 抜き取り確認で違っていた（以後は毎回読む）
        }
        private const int VerifyEvery = 16;
        private readonly Dictionary<string, ShapeTemplate> _templates = new Dictionary<string, ShapeTemplate>();
        /// <summary>中身を読まずに済ませた置き場所の数（ログ用）。</summary>
        public int SkippedInstances { get; private set; }
        /// <summary>抜き取り確認で違っていたファミリの数（ログ用）。</summary>
        public int VerifyMismatches { get; private set; }

        // 見た目が変わりうる条件（ビューのフィルタ・要素ごとの上書き）
        private readonly View _view;
        private readonly List<(ElementId id, ICollection<ElementId> categories, ElementFilter filter, ICollection<ElementId> selection)> _viewFilters =
            new List<(ElementId, ICollection<ElementId>, ElementFilter, ICollection<ElementId>)>();
        private string _elementSignature = "";
        private readonly Stack<ShapeBuilder> _shapes = new Stack<ShapeBuilder>();
        private readonly ShapeBuilder _loose = new ShapeBuilder { World = Transform.Identity, WorldInverse = Transform.Identity };
        // 中身のハッシュ → 同じハッシュの形の番号（中身を比べて確かめる）
        private readonly Dictionary<ulong, List<int>> _shapeByHash = new Dictionary<ulong, List<int>>();
        /// <summary>置き場所のうち、すでにある形を使い回した数（ログ用）。</summary>
        public int ReusedShapes { get; private set; }
        public int ElementCount { get; private set; }
        /// <summary>書き出した要素の ID（このモデルの物）。</summary>
        public HashSet<ElementId> ElementIds { get; } = new HashSet<ElementId>();
        /// <summary>見つからなかったテクスチャ画像（ログ用）。</summary>
        public IReadOnlyList<string> MissingTextures => _textures.Missing;

        public SceneExportContext(Document doc, PhaseOverrides phase = null, View view = null)
        {
            _phase = phase;
            _view = view;
            _documents.Push(doc);
            _transforms.Push(Transform.Identity);
            _textures = new TextureLoader(doc.PathName, doc.Application.VersionNumber);
            if (view != null)
            {
                try
                {
                    foreach (ElementId fid in view.GetFilters())
                    {
                        if (doc.GetElement(fid) is ParameterFilterElement pf)
                            _viewFilters.Add((fid, pf.GetCategories(), pf.GetElementFilter(), null));
                        else if (doc.GetElement(fid) is SelectionFilterElement sf)
                            _viewFilters.Add((fid, null, null, sf.GetElementIds()));
                    }
                }
                catch (Exception ex)
                {
                    DiagLog.Cmd("RealtimeRender", $"ビューのフィルタを読めませんでした ({ex.Message})");
                }
            }
        }

        public bool Start()
        {
            // マテリアルが一度も指定されない形状のための既定の色
            Scene.Materials.Add(new RenderMaterial { Name = "Default", R = 0.7f, G = 0.7f, B = 0.7f });
            return true;
        }

        public void Finish()
        {
            // 置き場所の外の形（壁・床など）は 1 つにまとめて、そのままの位置に置く
            CommitShape(_loose);
        }

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
            _currentElement = elementId;
            _phaseMaterial = PhaseMaterialOf(elementId);
            _elementSignature = OverrideSignatureOf(elementId);
            // 書き出した要素（このモデルの物だけ。リンクの中の要素はリンク自体で見る）。ライブ同期で、削除された要素が
            // 描いていた物かどうかを調べるのに使う
            if (_documents.Count == 1) ElementIds.Add(elementId);
            ElementCount++;
            return RenderNodeAction.Proceed;
        }

        public void OnElementEnd(ElementId elementId) { }

        public RenderNodeAction OnInstanceBegin(InstanceNode node)
        {
            Transform world = _transforms.Peek().Multiply(node.GetTransform());
            _transforms.Push(world);
            var b = new ShapeBuilder { World = world, WorldInverse = world.Inverse };
            string key = TemplateKeyOf(node);
            if (key != null)
            {
                if (_templates.TryGetValue(key, out ShapeTemplate tpl) && !tpl.Unsafe)
                {
                    tpl.Uses++;
                    if (tpl.Uses % VerifyEvery != 0)
                    {
                        // 同じ形を前に読んでいる: 中身は読まない（OnInstanceEnd で覚えた形を並べる）
                        b.Skipped = tpl;
                        _shapes.Push(b);
                        SkippedInstances++;
                        return RenderNodeAction.Skip;
                    }
                    b.Verify = tpl;  // 抜き取り確認: 読んで、覚えた形と比べる
                }
                b.TemplateKey = key;
                b.Recording = new List<TemplateEntry>();
            }
            _shapes.Push(b);
            return RenderNodeAction.Proceed;
        }

        // Skip を返したときも呼ばれる（RenderNodeAction.Skip の説明より）
        public void OnInstanceEnd(InstanceNode node)
        {
            ShapeBuilder b = _shapes.Peek();
            if (b.Skipped != null)
            {
                foreach (TemplateEntry e in b.Skipped.Entries) AddPlacement(e.Mesh, b.World.Multiply(e.Relative));
            }
            else
            {
                CommitShape(b);
            }
            _shapes.Pop();
            _transforms.Pop();
            if (b.Skipped != null || b.Recording == null || b.HasLight) return;
            if (b.Verify != null)
            {
                if (!SameEntries(b.Verify.Entries, b.Recording))
                {
                    b.Verify.Unsafe = true;
                    VerifyMismatches++;
                    DiagLog.Cmd("RealtimeRender", $"同じ形のはずのファミリで中身が違いました（以後は毎回読みます）: {b.TemplateKey}");
                }
            }
            else if (!_templates.ContainsKey(b.TemplateKey))
            {
                _templates[b.TemplateKey] = new ShapeTemplate { Entries = b.Recording };
            }
        }

        /// <summary>置き場所を足し、記録中の外側の置き場所（覚えようとしているファミリ）にも、そこから見た位置で覚えさせる。</summary>
        private void AddPlacement(int mesh, Transform world)
        {
            Scene.Instances.Add(new RenderInstance { Mesh = mesh, M = ToMatrix(world) });
            foreach (ShapeBuilder f in _shapes)
                if (f.Recording != null)
                    f.Recording.Add(new TemplateEntry { Mesh = mesh, Relative = f.WorldInverse.Multiply(world) });
        }

        private static bool SameEntries(List<TemplateEntry> a, List<TemplateEntry> b)
        {
            if (a.Count != b.Count) return false;
            for (int i = 0; i < a.Count; i++)
            {
                if (a[i].Mesh != b[i].Mesh) return false;
                Transform x = a[i].Relative, y = b[i].Relative;
                if (!x.Origin.IsAlmostEqualTo(y.Origin) || !x.BasisX.IsAlmostEqualTo(y.BasisX) ||
                    !x.BasisY.IsAlmostEqualTo(y.BasisY) || !x.BasisZ.IsAlmostEqualTo(y.BasisZ))
                    return false;
            }
            return true;
        }

        /// <summary>
        /// 2 回目から読まないための見分けの目印。ファミリの形の番号（Revit 2023 以降）に、文書・フェーズの上書き・
        /// ビューのフィルタと要素ごとの上書きを足したもの。使えなければ null（毎回読む）。
        /// </summary>
        private string TemplateKeyOf(InstanceNode node)
        {
#if REVIT2021 || REVIT2022
            return null;
#else
            try
            {
                SymbolGeometryId sid = node.GetSymbolGeometryId();
                string id = sid?.AsUniqueIdentifier();
                if (string.IsNullOrEmpty(id)) return null;
                return $"{_documents.Peek().PathName}|{id}|{_phaseMaterial}|{_elementSignature}";
            }
            catch (Exception)
            {
                return null;
            }
#endif
        }

        /// <summary>要素の見た目を変えうるビューの設定（当てはまるフィルタ・要素ごとの上書き）を文字にしたもの。</summary>
        private string OverrideSignatureOf(ElementId elementId)
        {
            if (_view == null || _documents.Count != 1) return "";
            try
            {
                Element e = _documents.Peek().GetElement(elementId);
                if (e == null) return "";
                var sb = new System.Text.StringBuilder();
                OverrideGraphicSettings o = _view.GetElementOverrides(elementId);
                Color c = o.SurfaceForegroundPatternColor;
                if (c != null && c.IsValid) sb.Append($"c{c.Red},{c.Green},{c.Blue}");
                sb.Append('t').Append(o.Transparency).Append(o.Halftone ? "h" : "");
                ElementId cat = e.Category?.Id;
                foreach (var f in _viewFilters)
                {
                    bool passes = f.selection != null
                        ? f.selection.Contains(elementId)
                        : cat != null && f.categories.Contains(cat) && (f.filter == null || f.filter.PassesFilter(e));
                    if (passes) sb.Append('|').Append(f.id);
                }
                return sb.ToString();
            }
            catch (Exception)
            {
                return "?" + elementId;  // 調べられなければ、その要素だけの目印にする（使い回さない）
            }
        }

        /// <summary>
        /// 集めた形をシーンに入れる。同じ中身の形がすでにあれば、その形を使い回して置き場所だけを足す。
        /// </summary>
        private void CommitShape(ShapeBuilder b)
        {
            if (b.Triangles.Count == 0) return;
            ulong hash = HashShape(b);
            if (!_shapeByHash.TryGetValue(hash, out List<int> candidates))
            {
                candidates = new List<int>();
                _shapeByHash[hash] = candidates;
            }
            int mesh = -1;
            foreach (int c in candidates)
                if (SameShape(c, b)) { mesh = c; break; }
            if (mesh >= 0)
            {
                ReusedShapes++;
            }
            else
            {
                int firstTri = Scene.TriangleCount, baseIndex = Scene.Vertices.Count;
                Scene.Vertices.AddRange(b.Vertices);
                for (int i = 0; i < b.Triangles.Count; i += 4)
                    Scene.AddTriangle(baseIndex + b.Triangles[i], baseIndex + b.Triangles[i + 1], baseIndex + b.Triangles[i + 2], b.Triangles[i + 3]);
                mesh = Scene.Meshes.Count;
                Scene.Meshes.Add(new RenderMesh { FirstTri = firstTri, TriCount = b.Triangles.Count / 4 });
                candidates.Add(mesh);
            }
            AddPlacement(mesh, b.World);
        }

        private static ulong HashShape(ShapeBuilder b)
        {
            // FNV-1a（64 ビット）。頂点の値とマテリアル番号をそのまま混ぜる
            ulong h = 14695981039346656037UL;
            void Mix(uint x)
            {
                for (int k = 0; k < 4; k++)
                {
                    h ^= (x >> (k * 8)) & 0xFF;
                    h *= 1099511628211UL;
                }
            }
            uint Bits(float f) => (uint)BitConverter.ToInt32(BitConverter.GetBytes(f), 0);
            Mix((uint)b.Vertices.Count);
            foreach (RenderVertex v in b.Vertices)
            {
                Mix(Bits(v.PX)); Mix(Bits(v.PY)); Mix(Bits(v.PZ));
                Mix(Bits(v.NX)); Mix(Bits(v.NY)); Mix(Bits(v.NZ));
                Mix(Bits(v.U)); Mix(Bits(v.V));
            }
            foreach (int i in b.Triangles) Mix((uint)i);
            return h;
        }

        /// <summary>シーンの形 mesh と、集めた形 b の中身が同じか（頂点の値・並び・マテリアル）。</summary>
        private bool SameShape(int mesh, ShapeBuilder b)
        {
            RenderMesh m = Scene.Meshes[mesh];
            if (m.TriCount * 4 != b.Triangles.Count) return false;
            int baseIndex = Scene.Triangles[m.FirstTri * 4];  // 形の頂点は続けて入れてあるので、最初の三角形の最初の頂点が先頭
            for (int i = 0; i < b.Triangles.Count; i++)
            {
                int stored = Scene.Triangles[m.FirstTri * 4 + i];
                if ((i % 4) == 3 ? stored != b.Triangles[i] : stored - baseIndex + b.Triangles[0] != b.Triangles[i]) return false;
            }
            int first = b.Triangles[0];
            for (int i = 0; i < b.Vertices.Count; i++)
            {
                int si = baseIndex - first + i;
                if (si < 0 || si >= Scene.Vertices.Count || !Scene.Vertices[si].Equals(b.Vertices[i])) return false;
            }
            return true;
        }

        /// <summary>Revit の Transform（フィート）を、メートルの 3 行 4 列（行ごと）にする。</summary>
        private static float[] ToMatrix(Transform t)
        {
            XYZ x = t.BasisX, y = t.BasisY, z = t.BasisZ, o = t.Origin;
            return new[]
            {
                (float)x.X, (float)y.X, (float)z.X, (float)(o.X * FeetToMeters),
                (float)x.Y, (float)y.Y, (float)z.Y, (float)(o.Y * FeetToMeters),
                (float)x.Z, (float)y.Z, (float)z.Z, (float)(o.Z * FeetToMeters),
            };
        }

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

        /// <summary>
        /// 照明器具の光源。ファミリの「測光」の設定（明るさ・色温度・配光・光源の形）から、描画プログラムの光源を作る。
        /// </summary>
        public void OnLight(LightNode node)
        {
            foreach (ShapeBuilder f in _shapes) f.HasLight = true;
            try
            {
                Document doc = _documents.Peek();
                RenderLight light = LightReader.Read(doc, _currentElement, _transforms.Peek().Multiply(node.GetTransform()));
                if (light != null) Scene.Lights.Add(light);
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"光源を読めませんでした: {_currentElement} ({ex.Message})");
            }
        }

        /// <summary>
        /// このビューのフェーズで「上書き」表示になる要素なら、上書きのマテリアル（このモデルの要素だけ。リンクはフェーズの対応が別）。
        /// </summary>
        private ElementId PhaseMaterialOf(ElementId elementId)
        {
            if (_phase == null || _documents.Count != 1) return ElementId.InvalidElementId;
            try
            {
                Element e = _documents.Peek().GetElement(elementId);
                if (e == null || !e.HasPhases()) return ElementId.InvalidElementId;
                return _phase.Materials.TryGetValue(e.GetPhaseStatus(_phase.ViewPhaseId), out ElementId m) ? m : ElementId.InvalidElementId;
            }
            catch (Exception)
            {
                return ElementId.InvalidElementId;
            }
        }

        public void OnMaterial(MaterialNode node)
        {
            Document doc = _documents.Peek();
            if (_phaseMaterial != ElementId.InvalidElementId && doc.GetElement(_phaseMaterial) is Material phaseMaterial)
            {
                UsePhaseMaterial(doc, phaseMaterial);
                return;
            }
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

            _currentMaterial = AddMaterial(key, name, r, g, b, node.Transparency, node.Smoothness,
                () => node.HasOverriddenAppearance ? node.GetAppearanceOverride() : node.GetAppearance());
        }

        /// <summary>フェーズの上書きのマテリアルで描く（マテリアルの要素の色・外観から作る）。</summary>
        private void UsePhaseMaterial(Document doc, Material m)
        {
            string key = $"phase|{doc.PathName}|{m.Id}";
            if (_materialIndex.TryGetValue(key, out int index))
            {
                _currentMaterial = index;
                return;
            }
            Color c = m.Color;
            bool valid = c != null && c.IsValid;
            _currentMaterial = AddMaterial(key, m.Name, valid ? c.Red : (byte)180, valid ? c.Green : (byte)180, valid ? c.Blue : (byte)180,
                m.Transparency / 100.0, m.Smoothness,
                () => (doc.GetElement(m.AppearanceAssetId) as AppearanceAssetElement)?.GetRenderingAsset());
        }

        /// <summary>
        /// マテリアルを作って登録し、その番号を返す。getAsset はマテリアルの「外観」（無ければ null）。
        /// </summary>
        private int AddMaterial(string key, string name, byte r, byte g, byte b, double transparency, int smoothness, Func<Asset> getAsset)
        {
            // マテリアルの「外観」タブの内容（色・イメージ・反射率・粗さ など）
            AppearanceInfo appearance = null;
            try
            {
                Asset asset = getAsset();
                if (asset != null) appearance = AppearanceReader.Read(asset);
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
                Opacity = (float)Clamp(1.0 - transparency, 0, 1),
                // Revit の「滑らかさ」0〜100 → 粗さ 1〜0（外観に粗さがあればそちらを使う）
                Roughness = (float)Clamp(1.0 - smoothness / 100.0, 0.05, 1),
                Metallic = MaterialHints.LooksMetallic(name) ? 1f : 0f,
            };

            if (appearance != null)
            {
                if (appearance.ColorImage != null)
                {
                    material.ColorTexture = GetTexture(appearance.ColorImage, srgb, appearance.ImageFade, false);
                    if (material.ColorTexture >= 0)
                    {
                        // 色は画像に焼き込んだので、掛け合わせる色は白にする
                        material.R = material.G = material.B = 1f;
                        material.ColorXform = appearance.ColorImage.Xform();
                    }
                }
                if (appearance.BumpImage != null && Math.Abs(appearance.BumpDepth) > 1e-6)
                {
                    material.BumpTexture = GetTexture(appearance.BumpImage, null, 1.0, true);
                    if (material.BumpTexture >= 0)
                    {
                        material.BumpXform = appearance.BumpImage.Xform();
                        material.BumpType = appearance.BumpType;
                        material.BumpDepth = (float)appearance.BumpDepth;
                    }
                }
                // 自己発光（輝度 cd/m² → 描画プログラムの単位。太陽 5 = 10 万ルクス）
                if (appearance.Luminance.HasValue) material.Emissive = (float)(appearance.Luminance.Value / 20000.0);

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

            int index = Scene.Materials.Count;
            Scene.Materials.Add(material);
            _materialIndex[key] = index;
            DiagLog.Cmd("RealtimeRender",
                $"マテリアル: {name} 色={BitmapInfo.FormatColor(srgb)} 画像={(material.ColorTexture >= 0 ? appearance?.ColorImage?.Paths : "-")} " +
                $"凹凸={(material.BumpTexture >= 0 ? $"{appearance?.BumpImage?.Paths} 種類={material.BumpType} 深さ={material.BumpDepth:G3}" : "-")} " +
                $"粗さ={material.Roughness:F2} 反射率={material.F0:F2} 不透明度={material.Opacity:F2} 金属={material.Metallic}");
            return index;
        }

        /// <summary>
        /// テクスチャ画像を読み込み、明るさ・色付け・「イメージ フェード」（色との混ぜ具合）を焼き込んで登録する。
        /// 戻り値はテクスチャ番号（読めなければ -1）。
        /// </summary>
        private int GetTexture(BitmapInfo bitmap, double[] baseSrgb, double fade, bool linear)
        {
            string key = $"{bitmap.BakeKey}|{fade:F3}|{BitmapInfo.FormatColor(fade < 1 ? baseSrgb : null)}|{linear}";
            if (_textureIndex.TryGetValue(key, out int index)) return index;
            index = -1;
            string path = _textures.Resolve(bitmap.Paths);
            // テクスチャが多いときは小さめにして、GPU のメモリを使いすぎないようにする
            RenderTexture texture = path != null ? TextureLoader.Decode(path, Scene.Textures.Count < 32 ? 1024 : 512) : null;
            if (texture != null)
            {
                texture.Linear = linear;
                if (linear)
                {
                    // 凹凸の画像は色ではないので、反転だけ行う
                    if (bitmap.Invert)
                        for (int i = 0; i < texture.Rgba.Length; i++)
                            if (i % 4 != 3) texture.Rgba[i] = (byte)(255 - texture.Rgba[i]);
                }
                else
                {
                    Bake(texture, bitmap, baseSrgb, fade);
                }
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
            // 置き場所の中なら、その置き場所から見た座標で集める（ふつうは置き場所の位置そのものなので、ほぼ単位行列）。
            // 置き場所の外（壁・床・リンクの中の要素など）は、プロジェクト座標で集める
            ShapeBuilder target = _shapes.Count > 0 ? _shapes.Peek() : _loose;
            Transform t = _shapes.Count > 0 ? target.WorldInverse.Multiply(_transforms.Peek()) : _transforms.Peek();
            IList<XYZ> points = mesh.GetPoints();
            IList<XYZ> normals = mesh.GetNormals();
            IList<PolymeshFacet> facets = mesh.GetFacets();
            if (points.Count == 0 || facets.Count == 0) return;
            // テクスチャのあるマテリアルだけ、面の上の位置（UV）を読む
            IList<UV> uvs = Scene.Materials[_currentMaterial].UsesUv ? mesh.GetUVs() : null;
            RenderVertex Vertex(int i, XYZ n) =>
                MakeVertex(t, points[i], n, uvs != null && i < uvs.Count ? uvs[i] : null);
            List<RenderVertex> verts = target.Vertices;
            List<int> tris = target.Triangles;
            void AddTriangle(int a, int b, int c)
            {
                tris.Add(a);
                tris.Add(b);
                tris.Add(c);
                tris.Add(_currentMaterial);
            }

            switch (mesh.DistributionOfNormals)
            {
                case DistributionOfNormals.AtEachPoint:
                {
                    int baseIndex = verts.Count;
                    for (int i = 0; i < points.Count; i++)
                        verts.Add(Vertex(i, i < normals.Count ? normals[i] : null));
                    foreach (PolymeshFacet f in facets)
                        AddTriangle(baseIndex + f.V1, baseIndex + f.V2, baseIndex + f.V3);
                    break;
                }
                case DistributionOfNormals.OnePerFace:
                {
                    XYZ n = normals.Count > 0 ? normals[0] : null;
                    int baseIndex = verts.Count;
                    for (int i = 0; i < points.Count; i++)
                        verts.Add(Vertex(i, n));
                    foreach (PolymeshFacet f in facets)
                        AddTriangle(baseIndex + f.V1, baseIndex + f.V2, baseIndex + f.V3);
                    break;
                }
                default: // OnEachFacet: 三角形ごとに向きが違うので、頂点を三角形ごとに分ける
                {
                    for (int i = 0; i < facets.Count; i++)
                    {
                        PolymeshFacet f = facets[i];
                        XYZ n = i < normals.Count ? normals[i] : null;
                        int baseIndex = verts.Count;
                        verts.Add(Vertex(f.V1, n));
                        verts.Add(Vertex(f.V2, n));
                        verts.Add(Vertex(f.V3, n));
                        AddTriangle(baseIndex, baseIndex + 1, baseIndex + 2);
                    }
                    break;
                }
            }
        }

        private static RenderVertex MakeVertex(Transform t, XYZ point, XYZ normal, UV uv)
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
            if (uv != null)
            {
                // Revit の UV はフィート（面の上の長さ）
                v.U = (float)(uv.U * FeetToMeters);
                v.V = (float)(uv.V * FeetToMeters);
            }
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

using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

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

        private int _currentMaterial;

        public RenderScene Scene { get; } = new RenderScene();
        public int ElementCount { get; private set; }

        public SceneExportContext(Document doc)
        {
            _documents.Push(doc);
            _transforms.Push(Transform.Identity);
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
            string key = $"{doc?.PathName}|{node.MaterialId?.ToString()}|{r},{g},{b}|{node.Transparency:F3}|{node.Smoothness}";
            if (_materialIndex.TryGetValue(key, out int index))
            {
                _currentMaterial = index;
                return;
            }

            string name = "";
            if (doc != null && node.MaterialId != null && node.MaterialId != ElementId.InvalidElementId)
                name = (doc.GetElement(node.MaterialId) as Material)?.Name ?? "";

            var material = new RenderMaterial
            {
                Name = name,
                // Revit の色は画面表示用（sRGB）なので、光の計算用（リニア）に直す
                R = SrgbToLinear(r),
                G = SrgbToLinear(g),
                B = SrgbToLinear(b),
                Opacity = (float)Math.Max(0.0, Math.Min(1.0, 1.0 - node.Transparency)),
                // Revit の「滑らかさ」0〜100 → 粗さ 1〜0
                Roughness = (float)Math.Max(0.05, Math.Min(1.0, 1.0 - node.Smoothness / 100.0)),
                Metallic = MaterialHints.LooksMetallic(name) ? 1f : 0f,
            };
            if (MaterialHints.LooksLikeGlass(name) && material.Opacity > 0.5f)
                material.Opacity = 0.3f;

            index = Scene.Materials.Count;
            Scene.Materials.Add(material);
            _materialIndex[key] = index;
            _currentMaterial = index;
        }

        public void OnPolymesh(PolymeshTopology mesh)
        {
            Transform t = _transforms.Peek();
            IList<XYZ> points = mesh.GetPoints();
            IList<XYZ> normals = mesh.GetNormals();
            IList<PolymeshFacet> facets = mesh.GetFacets();
            if (points.Count == 0 || facets.Count == 0) return;

            switch (mesh.DistributionOfNormals)
            {
                case DistributionOfNormals.AtEachPoint:
                {
                    int baseIndex = Scene.Vertices.Count;
                    for (int i = 0; i < points.Count; i++)
                        Scene.Vertices.Add(MakeVertex(t, points[i], i < normals.Count ? normals[i] : null));
                    foreach (PolymeshFacet f in facets)
                        Scene.AddTriangle(baseIndex + f.V1, baseIndex + f.V2, baseIndex + f.V3, _currentMaterial);
                    break;
                }
                case DistributionOfNormals.OnePerFace:
                {
                    XYZ n = normals.Count > 0 ? normals[0] : null;
                    int baseIndex = Scene.Vertices.Count;
                    for (int i = 0; i < points.Count; i++)
                        Scene.Vertices.Add(MakeVertex(t, points[i], n));
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
                        Scene.Vertices.Add(MakeVertex(t, points[f.V1], n));
                        Scene.Vertices.Add(MakeVertex(t, points[f.V2], n));
                        Scene.Vertices.Add(MakeVertex(t, points[f.V3], n));
                        Scene.AddTriangle(baseIndex, baseIndex + 1, baseIndex + 2, _currentMaterial);
                    }
                    break;
                }
            }
        }

        private static RenderVertex MakeVertex(Transform t, XYZ point, XYZ normal)
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
            return v;
        }

        private static float SrgbToLinear(byte c)
        {
            double s = c / 255.0;
            return (float)(s <= 0.04045 ? s / 12.92 : Math.Pow((s + 0.055) / 1.055, 2.4));
        }
    }
}

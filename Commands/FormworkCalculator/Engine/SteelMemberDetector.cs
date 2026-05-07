using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 構造柱・構造フレームの中から、型枠不要な鉄骨部材
    /// （H形鋼、角形/円形鋼管、溝形鋼、山形鋼、CFT等）を識別する。
    ///
    /// 4 層フォールバック:
    ///   L1: FamilyInstance.StructuralMaterialType == Steel
    ///   L2: 断面形状分析（ExtrusionAnalyzer による中空検出 + 充実率判定）
    ///   L3: 構造材マテリアルの MaterialClass / 名前
    ///   L4: ファミリ名 / タイプ名のキーワード判定
    ///
    /// SRC柱は L1=Concrete, L2=中実凸, L3=Concrete → 保持される（型枠必要）。
    /// CFT柱は L2 が中空を検出する（鋼管モデリング時）か、L4 で "CFT" を検出する。
    /// </summary>
    internal static class SteelMemberDetector
    {
        public enum DetectionLayer
        {
            None,
            StructuralMaterialType,
            Shape,
            MaterialClass,
            NamePattern,
        }

        public class DetectionResult
        {
            public bool IsSteel;
            public DetectionLayer Layer;
            public string Reason = string.Empty;
        }

        /// <summary>形状判定の充実率しきい値。これ未満なら鉄骨と判定。</summary>
        private const double AreaRatioThreshold = 0.5;

        /// <summary>
        /// 「明確に鉄骨」と判断できるキーワード。Family/Type/Element 名のいずれかに含まれていれば鉄骨。
        /// </summary>
        private static readonly string[] _strongKeywords = new[]
        {
            // 日本語: JIS 形状名
            "H形鋼", "I形鋼", "C形鋼", "L形鋼", "T形鋼",
            "角形鋼管", "円形鋼管", "溝形鋼", "山形鋼", "平鋼",
            "ハット形鋼", "リップ溝形鋼",
            // 日本語: 工法・略称
            "鉄骨", "ＳＲＣ", "CFT", "ＣＦＴ",
            // JIS 鋼材記号
            "STKR", "STKN", "STKM", "BCR", "BCP", "STK-", "FB-",
            // 英語
            "WIDE FLANGE", "I-BEAM", "I BEAM", "H-SECTION",
            "C-CHANNEL", "CHANNEL", "ANGLE",
            "HSS", "HOLLOW STRUCTURAL", "HOLLOW SECTION",
            "STEEL TUBE", "PIPE COLUMN", "PIPE BEAM",
        };

        /// <summary>
        /// ファミリ名 / タイプ名 が以下のいずれかで始まれば鉄骨と判定する弱パターン。
        /// （単独の "C-" 等は概念単独だと曖昧なので、文字列の先頭でのみ採用）
        /// </summary>
        private static readonly string[] _prefixPatterns = new[]
        {
            "H-", "BH-", "WF-", "I-",
            "□-", "■-", "○-", "●-",
            "C-", "L-", "T-", "FB-", "PL-", "BOX-", "PIPE-",
            "Ｈ－", "Ｃ－", "Ｌ－", "Ｔ－", "□－", "○－",
        };

        /// <summary>
        /// マテリアル名 / MaterialClass がこれらを含めば鉄骨と判定。
        /// </summary>
        private static readonly string[] _steelMaterialKeywords = new[]
        {
            "STEEL", "METAL", "鋼", "鉄",
        };

        /// <summary>
        /// 要素を 4 層のロジックで判定する。鉄骨と判断されなければ IsSteel=false を返す。
        /// </summary>
        internal static DetectionResult Detect(Element elem, Document doc)
        {
            if (elem == null) return new DetectionResult();

            // L1: StructuralMaterialType
            try
            {
                if (elem is FamilyInstance fi)
                {
                    var smt = fi.StructuralMaterialType;
                    if (smt == StructuralMaterialType.Steel)
                    {
                        return new DetectionResult
                        {
                            IsSteel = true,
                            Layer = DetectionLayer.StructuralMaterialType,
                            Reason = "StructuralMaterialType=Steel",
                        };
                    }
                }
            }
            catch { }

            // L2: 断面形状分析
            try
            {
                var shapeRes = AnalyzeShape(elem);
                if (shapeRes != null && shapeRes.IsSteel) return shapeRes;
            }
            catch (Exception ex)
            {
                FormworkDebugLog.Log($"  [SteelDetect] L2 shape analysis exception: {ex.Message}");
            }

            // L3: マテリアルクラス / 名前
            try
            {
                var matRes = AnalyzeMaterial(elem, doc);
                if (matRes != null && matRes.IsSteel) return matRes;
            }
            catch { }

            // L4: 名前パターン
            try
            {
                var nameRes = AnalyzeNamePattern(elem);
                if (nameRes != null && nameRes.IsSteel) return nameRes;
            }
            catch { }

            return new DetectionResult { IsSteel = false, Layer = DetectionLayer.None, Reason = "no match" };
        }

        // ---------------- Layer 2: 形状分析 ----------------

        private static DetectionResult AnalyzeShape(Element elem)
        {
            var solids = SolidUnionProcessor.GetSolids(elem);
            if (solids.Count == 0) return null;

            // 最大体積の Solid を主部材とみなす
            Solid mainSolid = null;
            double maxVol = 0;
            foreach (var s in solids)
            {
                if (s.Volume > maxVol)
                {
                    maxVol = s.Volume;
                    mainSolid = s;
                }
            }
            if (mainSolid == null) return null;

            XYZ axis = GetLongitudinalAxis(elem);
            if (axis == null) return null;

            BoundingBoxXYZ bb;
            try { bb = elem.get_BoundingBox(null); } catch { bb = null; }
            XYZ origin;
            if (bb != null)
                origin = (bb.Min + bb.Max) * 0.5;
            else
                origin = XYZ.Zero;

            Plane plane;
            try
            {
                plane = Plane.CreateByNormalAndOrigin(axis, origin);
            }
            catch { return null; }

            Face profile = null;
            try
            {
                var ea = ExtrusionAnalyzer.Create(mainSolid, plane, axis);
                profile = ea?.GetExtrusionBase();
            }
            catch
            {
                // 単純押し出しでない要素 (テーパ、ハンチ等) は判定不可 → null で次の層へ
                return null;
            }
            if (profile == null) return null;

            // 中空判定: エッジループが 2 つ以上あれば穴あり
            IList<CurveLoop> loops = null;
            try { loops = profile.GetEdgesAsCurveLoops(); } catch { }
            if (loops != null && loops.Count >= 2)
            {
                return new DetectionResult
                {
                    IsSteel = true,
                    Layer = DetectionLayer.Shape,
                    Reason = $"hollow profile (loops={loops.Count})",
                };
            }

            // 充実率: プロファイル面積 / UV バウンディングボックス面積
            double profileArea = 0;
            try { profileArea = profile.Area; } catch { }
            if (profileArea <= 0) return null;

            double bboxArea = 0;
            try
            {
                var uvBb = profile.GetBoundingBox();
                if (uvBb != null)
                    bboxArea = (uvBb.Max.U - uvBb.Min.U) * (uvBb.Max.V - uvBb.Min.V);
            }
            catch { }
            if (bboxArea <= 0) return null;

            double ratio = profileArea / bboxArea;
            if (ratio < AreaRatioThreshold)
            {
                return new DetectionResult
                {
                    IsSteel = true,
                    Layer = DetectionLayer.Shape,
                    Reason = $"non-convex profile (areaRatio={ratio:F3})",
                };
            }

            return new DetectionResult
            {
                IsSteel = false,
                Layer = DetectionLayer.Shape,
                Reason = $"solid convex profile (areaRatio={ratio:F3})",
            };
        }

        private static XYZ GetLongitudinalAxis(Element elem)
        {
            // 構造柱: 鉛直 Z 方向
            if (elem.Category != null &&
                elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns)
            {
                return XYZ.BasisZ;
            }

            // 構造フレーム (梁): LocationCurve から軸方向を取得
            if (elem is FamilyInstance fi)
            {
                var loc = fi.Location as LocationCurve;
                if (loc?.Curve != null)
                {
                    try
                    {
                        XYZ p1 = loc.Curve.GetEndPoint(0);
                        XYZ p2 = loc.Curve.GetEndPoint(1);
                        var dir = p2 - p1;
                        if (dir.GetLength() > 1e-6) return dir.Normalize();
                    }
                    catch { }
                }
            }
            return null;
        }

        // ---------------- Layer 3: マテリアル分析 ----------------

        private static DetectionResult AnalyzeMaterial(Element elem, Document doc)
        {
            ElementId matId = ElementId.InvalidElementId;

            try
            {
                if (elem is FamilyInstance fi)
                    matId = fi.StructuralMaterialId;
            }
            catch { }

            if (matId == null || matId == ElementId.InvalidElementId)
            {
                try
                {
                    var p = elem.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM);
                    if (p != null && p.HasValue && p.StorageType == StorageType.ElementId)
                        matId = p.AsElementId();
                }
                catch { }
            }

            if (matId == null || matId == ElementId.InvalidElementId) return null;

            var mat = doc.GetElement(matId) as Material;
            if (mat == null) return null;

            string matName = mat.Name ?? string.Empty;
            string matClass = string.Empty;
            try { matClass = mat.MaterialClass ?? string.Empty; } catch { }

            string nameUpper = matName.ToUpperInvariant();
            string classUpper = matClass.ToUpperInvariant();

            foreach (var kw in _steelMaterialKeywords)
            {
                if (classUpper.Contains(kw) || nameUpper.Contains(kw))
                {
                    return new DetectionResult
                    {
                        IsSteel = true,
                        Layer = DetectionLayer.MaterialClass,
                        Reason = $"material '{matName}' class='{matClass}' matched '{kw}'",
                    };
                }
            }
            return null;
        }

        // ---------------- Layer 4: 名前パターン ----------------

        private static DetectionResult AnalyzeNamePattern(Element elem)
        {
            string famName = string.Empty;
            string typeName = string.Empty;
            string elemName = elem.Name ?? string.Empty;

            try
            {
                if (elem is FamilyInstance fi)
                {
                    typeName = fi.Symbol?.Name ?? string.Empty;
                    famName = fi.Symbol?.Family?.Name ?? string.Empty;
                }
            }
            catch { }

            string combined = (famName + " " + typeName + " " + elemName).ToUpperInvariant();

            foreach (var kw in _strongKeywords)
            {
                if (combined.Contains(kw.ToUpperInvariant()))
                {
                    return new DetectionResult
                    {
                        IsSteel = true,
                        Layer = DetectionLayer.NamePattern,
                        Reason = $"strong keyword '{kw}' in name (fam='{famName}' type='{typeName}')",
                    };
                }
            }

            // 弱パターンはファミリ名 / タイプ名の先頭 (前置きスペース許容) でのみ採用
            string famUpper = famName.ToUpperInvariant().TrimStart();
            string typeUpper = typeName.ToUpperInvariant().TrimStart();
            foreach (var prefix in _prefixPatterns)
            {
                string p = prefix.ToUpperInvariant();
                if (famUpper.StartsWith(p) || typeUpper.StartsWith(p))
                {
                    return new DetectionResult
                    {
                        IsSteel = true,
                        Layer = DetectionLayer.NamePattern,
                        Reason = $"prefix '{prefix}' (fam='{famName}' type='{typeName}')",
                    };
                }
            }

            return null;
        }
    }
}

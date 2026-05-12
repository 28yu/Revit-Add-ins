using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 階段カテゴリの中から、型枠不要な鉄骨階段を識別する。
    ///
    /// 2 層フォールバック:
    ///   L1: タイプ名 / ファミリ名 / 要素名のキーワード判定
    ///   L2: 構造材マテリアル (StairsType の MaterialClass / 名前)
    ///
    /// RC階段・SRC階段は L1/L2 とも鉄骨キーワードに引っかからず保持される（型枠必要）。
    /// </summary>
    internal static class SteelStairDetector
    {
        /// <summary>
        /// タイプ名・要素名に含まれていれば鉄骨階段と判定するキーワード。
        /// </summary>
        private static readonly string[] _nameKeywords = new[]
        {
            // 日本語
            "鉄骨", "鉄階段", "鋼製", "Ｓ造", "S造", "S階段", "Ｓ階段",
            // 英語
            "STEEL", "METAL", "STK",
        };

        /// <summary>
        /// マテリアル名 / MaterialClass に含まれていれば鉄骨と判定するキーワード。
        /// </summary>
        private static readonly string[] _materialKeywords = new[]
        {
            "STEEL", "METAL", "鋼", "鉄",
        };

        public class DetectionResult
        {
            public bool IsSteel;
            public string Layer = string.Empty;
            public string Reason = string.Empty;
        }

        /// <summary>
        /// 階段要素を鉄骨判定する。鉄骨と判断されなければ IsSteel=false を返す。
        /// </summary>
        internal static DetectionResult Detect(Element elem, Document doc)
        {
            if (elem == null) return new DetectionResult();

            // L1: 名前パターン (タイプ名 / ファミリ名 / 要素名)
            var nameRes = AnalyzeNamePattern(elem, doc);
            if (nameRes != null && nameRes.IsSteel) return nameRes;

            // L2: マテリアル分析 (StairsType の構造材料)
            var matRes = AnalyzeMaterial(elem, doc);
            if (matRes != null && matRes.IsSteel) return matRes;

            return new DetectionResult { IsSteel = false, Layer = "None", Reason = "no match" };
        }

        // ---------------- Layer 1: 名前パターン ----------------

        private static DetectionResult AnalyzeNamePattern(Element elem, Document doc)
        {
            string typeName = string.Empty;
            string famName = string.Empty;
            string elemName = elem.Name ?? string.Empty;

            try
            {
                var typeId = elem.GetTypeId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    var et = doc.GetElement(typeId) as ElementType;
                    typeName = et?.Name ?? string.Empty;
                    famName = et?.FamilyName ?? string.Empty;
                }
            }
            catch { }

            string combined = (typeName + " " + famName + " " + elemName).ToUpperInvariant();
            foreach (var kw in _nameKeywords)
            {
                if (combined.Contains(kw.ToUpperInvariant()))
                {
                    return new DetectionResult
                    {
                        IsSteel = true,
                        Layer = "NamePattern",
                        Reason = $"keyword '{kw}' (type='{typeName}' fam='{famName}')",
                    };
                }
            }
            return null;
        }

        // ---------------- Layer 2: マテリアル分析 ----------------

        private static DetectionResult AnalyzeMaterial(Element elem, Document doc)
        {
            // 階段は MATERIAL_ID_PARAM (材料 ID パラメータ) を持たないが、
            // タイプから階段の各部 (踏板・蹴上げ・側桁) の構造材を取得できる。
            // ここでは MaterialIds を全て収集し、ひとつでも鉄骨マテリアルがあれば鉄骨と判定する。

            try
            {
                var typeId = elem.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId) return null;
                var et = doc.GetElement(typeId) as ElementType;
                if (et == null) return null;

                // タイプの全パラメータをスキャンし、StorageType=ElementId かつ Material を参照するものを収集
                foreach (Parameter p in et.Parameters)
                {
                    if (p == null || !p.HasValue) continue;
                    if (p.StorageType != StorageType.ElementId) continue;
                    var id = p.AsElementId();
                    if (id == null || id == ElementId.InvalidElementId) continue;
                    var mat = doc.GetElement(id) as Material;
                    if (mat == null) continue;

                    string matName = (mat.Name ?? string.Empty).ToUpperInvariant();
                    string matClass = string.Empty;
                    try { matClass = (mat.MaterialClass ?? string.Empty).ToUpperInvariant(); } catch { }

                    foreach (var kw in _materialKeywords)
                    {
                        string ku = kw.ToUpperInvariant();
                        if (matClass.Contains(ku) || matName.Contains(ku))
                        {
                            return new DetectionResult
                            {
                                IsSteel = true,
                                Layer = "MaterialClass",
                                Reason = $"material '{mat.Name}' class='{matClass}' matched '{kw}'",
                            };
                        }
                    }
                }
            }
            catch { }
            return null;
        }
    }
}

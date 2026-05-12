using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 壁カテゴリの中から、LGS壁 (軽量鉄骨下地の乾式壁) を識別する。
    ///
    /// LGS壁は石膏ボード (PB) を主構成とする乾式壁で、コンクリート打設を伴わず型枠不要。
    /// RC壁に石膏ボード仕上げを貼った構成と区別するため、以下の条件で判定:
    ///   - L1: WallType 名に "LGS" / "軽鉄" / "軽量鉄骨" / "軽量間仕切" を含む
    ///   - L2: CompoundStructure に石膏ボード層が存在し、かつコンクリート層が存在しない
    ///        (RC壁 + 石膏ボード仕上げ等を誤検出しないため)
    ///
    /// 石膏ボードの代表的な厚さ: 12.5mm, 9.5mm, 15mm, 21mm (日本)
    /// ただし厚さ単独では判定せず、必ず材料名のキーワードと組み合わせて判定する。
    /// </summary>
    internal static class LgsWallDetector
    {
        /// <summary>
        /// 石膏ボード系材料のキーワード (大文字比較)。
        /// </summary>
        private static readonly string[] _gypsumKeywords = new[]
        {
            "GYPSUM", "PLASTERBOARD", "PLASTER BOARD", "PLASTER-BOARD",
            "GWB", "DRYWALL", "SHEETROCK",
            "石膏ボード", "プラスターボード", "石膏",
        };

        /// <summary>
        /// コンクリート系材料のキーワード (大文字比較)。
        /// これがあれば RC壁 + 石膏ボード仕上げと判断し、LGS壁とはしない。
        /// </summary>
        private static readonly string[] _concreteKeywords = new[]
        {
            "CONCRETE", "コンクリート",
        };

        /// <summary>
        /// タイプ名から直接 LGS壁と判別するキーワード。
        /// </summary>
        private static readonly string[] _lgsTypeNameKeywords = new[]
        {
            "LGS", "軽鉄", "軽量鉄骨", "軽量間仕切",
        };

        /// <summary>
        /// 壁要素を LGS壁と判定する。判定できなければ false を返す。
        /// </summary>
        internal static bool IsLgsWall(Element elem, Document doc, out string reason)
        {
            reason = string.Empty;
            if (!(elem is Wall wall)) return false;
            var wt = doc.GetElement(wall.GetTypeId()) as WallType;
            if (wt == null) return false;

            // L1: タイプ名キーワード
            string typeName = wt.Name ?? string.Empty;
            string typeUpper = typeName.ToUpperInvariant();
            foreach (var kw in _lgsTypeNameKeywords)
            {
                if (typeUpper.Contains(kw.ToUpperInvariant()))
                {
                    reason = $"type name '{typeName}' contains '{kw}'";
                    return true;
                }
            }

            // L2: CompoundStructure の層を走査
            CompoundStructure cs = null;
            try { cs = wt.GetCompoundStructure(); } catch { }
            if (cs == null) return false;

            bool hasGypsum = false;
            bool hasConcrete = false;
            string gypsumDesc = string.Empty;

            foreach (var layer in cs.GetLayers())
            {
                ElementId matId = layer.MaterialId;
                if (matId == null || matId == ElementId.InvalidElementId) continue;
                var mat = doc.GetElement(matId) as Material;
                if (mat == null) continue;

                string matName = (mat.Name ?? string.Empty).ToUpperInvariant();
                string matClass = string.Empty;
                try { matClass = (mat.MaterialClass ?? string.Empty).ToUpperInvariant(); } catch { }
                string combined = matName + " " + matClass;

                bool layerGypsum = false;
                foreach (var kw in _gypsumKeywords)
                {
                    if (combined.Contains(kw.ToUpperInvariant()))
                    {
                        layerGypsum = true;
                        break;
                    }
                }
                if (layerGypsum)
                {
                    hasGypsum = true;
                    if (gypsumDesc.Length == 0)
                    {
                        double thicknessMm = 0;
                        try
                        {
                            thicknessMm = UnitUtils.ConvertFromInternalUnits(
                                layer.Width, UnitTypeId.Millimeters);
                        }
                        catch { }
                        gypsumDesc = $"layer '{mat.Name}' ({thicknessMm:F1}mm)";
                    }
                    continue;
                }

                foreach (var kw in _concreteKeywords)
                {
                    if (combined.Contains(kw.ToUpperInvariant()))
                    {
                        hasConcrete = true;
                        break;
                    }
                }
            }

            if (hasGypsum && !hasConcrete)
            {
                reason = $"non-concrete wall with gypsum board layer: {gypsumDesc}";
                return true;
            }
            return false;
        }
    }
}

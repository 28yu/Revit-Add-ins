using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 壁カテゴリの中から、型枠不要な ALC パネル (軽量気泡コンクリートパネル) や
    /// ECP パネル (押出成形セメント板) を識別する。
    ///
    /// 判定方法: 壁タイプ名・ファミリ名・要素名のいずれかに "ALC" または "ECP" を含めば
    /// 工場製品の取付パネルとみなし型枠不要として除外する。
    /// </summary>
    internal static class AlcEcpPanelDetector
    {
        /// <summary>
        /// 大文字小文字を区別せず "ALC" または "ECP" の文字列が含まれているか判定。
        /// </summary>
        internal static bool IsAlcEcpPanel(Element elem, Document doc, out string reason)
        {
            reason = string.Empty;
            if (elem == null) return false;

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
            if (combined.Contains("ALC"))
            {
                reason = $"keyword 'ALC' (type='{typeName}' fam='{famName}')";
                return true;
            }
            if (combined.Contains("ECP"))
            {
                reason = $"keyword 'ECP' (type='{typeName}' fam='{famName}')";
                return true;
            }
            return false;
        }
    }
}

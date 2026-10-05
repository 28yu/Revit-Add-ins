using Autodesk.Revit.DB;

namespace Tools28
{
    /// <summary>
    /// パラメータグループ（プロパティパレットの見出し。例: 「寸法」「識別情報」）の表示名を取得する共通ヘルパー。
    /// Revit 2022 で API が変わったため、2021 のみ旧 API（BuiltInParameterGroup）を使う。
    /// 表示名は Revit 本体の UI 言語で返る。
    /// </summary>
    internal static class ParameterGroupHelper
    {
        /// <summary>
        /// パラメータグループの表示名を返す。
        /// グループ未設定（「その他」）や取得失敗時は otherLabel を返す。
        /// </summary>
        public static string GetGroupLabel(Definition definition, string otherLabel)
        {
            if (definition == null) return otherLabel;
            try
            {
#if REVIT2021
                var group = definition.ParameterGroup;
                if (group == BuiltInParameterGroup.INVALID)
                    return otherLabel;
                string label = LabelUtils.GetLabelFor(group);
#else
                // 「その他」グループは空の ForgeTypeId で返る
                var groupId = definition.GetGroupTypeId();
                if (groupId == null || string.IsNullOrEmpty(groupId.TypeId))
                    return otherLabel;
                string label = LabelUtils.GetLabelForGroup(groupId);
#endif
                return string.IsNullOrEmpty(label) ? otherLabel : label;
            }
            catch
            {
                return otherLabel;
            }
        }
    }
}

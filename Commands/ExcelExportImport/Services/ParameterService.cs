using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.ExcelExportImport.Models;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// パラメータの取得/設定サービス
    /// </summary>
    public static class ParameterService
    {
        /// <summary>
        /// 指定カテゴリの要素から全パラメータ情報を取得
        /// </summary>
        public static List<ParameterInfo> GetParametersForCategory(
            Document doc,
            BuiltInCategory category,
            string categoryName,
            Models.ExportScope scope = Models.ExportScope.EntireProject,
            View activeView = null,
            ICollection<ElementId> selectionIds = null)
        {
            var parameters = new HashSet<ParameterInfo>();
            var seenTypeIds = new HashSet<ElementId>();

            var elements = RevitCategoryHelper.GetElementsByCategory(
                doc, category, scope, activeView, selectionIds);

            // 各タイプにつき先頭インスタンス1件だけ調べる
            // （同カテゴリ・同タイプのインスタンスパラメータは同一なので十分）
            bool instanceParamsCollected = false;
            foreach (var elem in elements)
            {
                var typeId = elem.GetTypeId();
                bool isNewType = typeId != null && typeId != ElementId.InvalidElementId
                                 && seenTypeIds.Add(typeId);

                // 初回または新タイプのインスタンスからパラメータを収集
                if (!instanceParamsCollected || isNewType)
                {
                    CollectParameters(elem, false, categoryName, parameters);
                    instanceParamsCollected = true;
                }

                if (isNewType)
                {
                    var elemType = doc.GetElement(typeId);
                    if (elemType != null)
                        CollectParameters(elemType, true, categoryName, parameters);
                }
            }

            // カテゴリに属する全タイプからもパラメータを収集
            // （スコープ外・未使用タイプのパラメータも拾う）
            var allTypes = new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsElementType()
                .ToList();
            foreach (var et in allTypes)
            {
                if (!seenTypeIds.Add(et.Id))
                    continue;
                CollectParameters(et, true, categoryName, parameters);
            }

            // インスタンスパラメータが取れていなければ、
            // プロジェクト全体から1件だけサンプリング（スコープ外でも）
            if (!instanceParamsCollected)
            {
                var anyInstance = new FilteredElementCollector(doc)
                    .OfCategory(category)
                    .WhereElementIsNotElementType()
                    .FirstElement();
                if (anyInstance != null)
                    CollectParameters(anyInstance, false, categoryName, parameters);
            }

            return parameters.OrderBy(p => p.DisplayName).ToList();
        }

        private static void CollectParameters(
            Element element,
            bool isTypeParameter,
            string categoryName,
            HashSet<ParameterInfo> bucket)
        {
            if (element == null) return;

            foreach (Parameter param in element.Parameters)
            {
                if (param?.Definition == null || string.IsNullOrEmpty(param.Definition.Name))
                    continue;

                bucket.Add(new ParameterInfo(
                    param.Definition.Name,
                    isTypeParameter,
                    param.IsReadOnly,
                    categoryName));
            }
        }

        /// <summary>
        /// パラメータ値を文字列として取得
        /// </summary>
        public static string GetParameterValueAsString(Parameter param)
        {
            if (param == null || !param.HasValue)
                return "";

            switch (param.StorageType)
            {
                case StorageType.String:
                    return param.AsString() ?? "";
                case StorageType.Integer:
                    return param.AsInteger().ToString();
                case StorageType.Double:
                    return param.AsValueString() ?? param.AsDouble().ToString();
                case StorageType.ElementId:
#if REVIT2026
                    return param.AsValueString() ?? param.AsElementId().Value.ToString();
#else
                    return param.AsValueString() ?? param.AsElementId().IntegerValue.ToString();
#endif
                default:
                    return "";
            }
        }

        /// <summary>
        /// パラメータ値を文字列から設定
        /// </summary>
        public static bool SetParameterValue(Parameter param, string value)
        {
            if (param == null || param.IsReadOnly)
                return false;

            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        param.Set(value);
                        return true;

                    case StorageType.Integer:
                        if (int.TryParse(value, out int intVal))
                        {
                            param.Set(intVal);
                            return true;
                        }
                        return false;

                    case StorageType.Double:
                        // AsValueString形式での設定を試行
                        if (param.SetValueString(value))
                            return true;
                        // 直接数値設定
                        if (double.TryParse(value, out double dblVal))
                        {
                            param.Set(dblVal);
                            return true;
                        }
                        return false;

                    case StorageType.ElementId:
                        // SetValueStringを試行
                        try
                        {
                            if (param.SetValueString(value))
                                return true;
                        }
                        catch { }
                        // 直接ElementId設定
                        if (int.TryParse(value, out int idVal))
                        {
                            param.Set(new ElementId(idVal));
                            return true;
                        }
                        return false;

                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 要素からパラメータを名前で検索（インスタンス優先）
        /// </summary>
        public static Parameter FindParameter(Element elem, string paramName, bool isTypeParameter, Document doc)
        {
            if (isTypeParameter)
            {
                var typeId = elem.GetTypeId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    var elemType = doc.GetElement(typeId);
                    if (elemType != null)
                    {
                        return FindParameterByName(elemType, paramName);
                    }
                }
                return null;
            }
            else
            {
                return FindParameterByName(elem, paramName);
            }
        }

        /// <summary>
        /// 要素のタイプを名前で変更する（「タイプ」パラメータ用）
        /// </summary>
        public static bool ChangeElementType(Element elem, string typeName, Document doc)
        {
            if (elem == null || doc == null || string.IsNullOrEmpty(typeName))
                return false;

            try
            {
                // 現在のタイプのファミリIDを取得して、同一ファミリ内のタイプを検索
                var currentTypeId = elem.GetTypeId();
                if (currentTypeId == null || currentTypeId == ElementId.InvalidElementId)
                    return false;

                var currentType = doc.GetElement(currentTypeId) as ElementType;
                if (currentType == null)
                    return false;

                // 同カテゴリの全ファミリタイプから名前が一致するものを検索
                var collector = new FilteredElementCollector(doc)
                    .OfClass(currentType.GetType())
                    .WhereElementIsElementType();

                // FamilySymbolの場合、同じファミリ内のタイプを優先検索
                var familySymbol = currentType as FamilySymbol;
                ElementId targetTypeId = null;

                // 入力値を正規化（コロン前後のスペース差異を吸収）
                string normalizedInput = NormalizeColonSpacing(typeName.Trim());
                // 入力値からタイプ名部分を抽出（"ファミリ名: タイプ名" → "タイプ名"）
                string extractedTypeName = ExtractTypeNamePart(typeName.Trim());

                foreach (var typeElem in collector)
                {
                    var et = typeElem as ElementType;
                    if (et == null) continue;

                    string candidateName = et.Name;

                    // FamilySymbolの場合は "ファミリ名: タイプ名" 形式でチェック
                    if (familySymbol != null && typeElem is FamilySymbol fs)
                    {
                        // AsValueString() は "ファミリ名 : タイプ名" 形式を返すことがあるため
                        // コロン前後のスペースを正規化して比較
                        string fullName = fs.FamilyName + ": " + fs.Name;
                        string normalizedFullName = NormalizeColonSpacing(fullName.Trim());

                        if (string.Equals(normalizedFullName, normalizedInput, StringComparison.OrdinalIgnoreCase))
                        {
                            targetTypeId = fs.Id;
                            break;
                        }
                    }

                    // タイプ名のみで比較
                    if (string.Equals(candidateName.Trim(), typeName.Trim(), StringComparison.OrdinalIgnoreCase)
                        || string.Equals(candidateName.Trim(), extractedTypeName, StringComparison.OrdinalIgnoreCase))
                    {
                        // 同じファミリ内のタイプを優先
                        if (familySymbol != null && typeElem is FamilySymbol fs2
                            && fs2.FamilyName == familySymbol.FamilyName)
                        {
                            targetTypeId = fs2.Id;
                            break;
                        }
                        // ファミリが異なる場合も候補として保持
                        if (targetTypeId == null)
                            targetTypeId = et.Id;
                    }
                }

                if (targetTypeId != null && targetTypeId != currentTypeId)
                {
                    elem.ChangeTypeId(targetTypeId);
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// コロン前後のスペースを正規化（" : " や ": " や " :" を全て ":" に統一）
        /// AsValueString() と FamilyName + ": " + Name の形式差異を吸収
        /// </summary>
        private static string NormalizeColonSpacing(string name)
        {
            if (string.IsNullOrEmpty(name))
                return name;
            // " : " → ":", ": " → ":", " :" → ":"
            return name.Replace(" : ", ":").Replace(": ", ":").Replace(" :", ":");
        }

        /// <summary>
        /// "ファミリ名: タイプ名" 形式からタイプ名部分を抽出
        /// コロンが含まれない場合は入力値をそのまま返す
        /// </summary>
        private static string ExtractTypeNamePart(string fullName)
        {
            if (string.IsNullOrEmpty(fullName))
                return fullName;
            int colonIndex = fullName.LastIndexOf(':');
            if (colonIndex > 0 && colonIndex < fullName.Length - 1)
                return fullName.Substring(colonIndex + 1).Trim();
            return fullName;
        }

        /// <summary>
        /// パラメータがタイプ変更パラメータ（ELEM_TYPE_PARAM）かどうか判定
        /// </summary>
        public static bool IsTypeChangeParameter(Parameter param)
        {
            if (param == null || param.Definition == null)
                return false;

            // BuiltInParameterの ELEM_TYPE_PARAM は「タイプ」パラメータ
#if REVIT2026
            if (param.Id.Value == (long)BuiltInParameter.ELEM_TYPE_PARAM)
                return true;
#else
            if (param.Id.IntegerValue == (int)BuiltInParameter.ELEM_TYPE_PARAM)
                return true;
#endif

            // パラメータ名が「タイプ」でStorageTypeがElementIdの場合も対象
            if (param.StorageType == StorageType.ElementId
                && param.Definition.Name == "タイプ")
                return true;

            return false;
        }

        private static Parameter FindParameterByName(Element elem, string paramName)
        {
            foreach (Parameter param in elem.Parameters)
            {
                if (param.Definition != null && param.Definition.Name == paramName)
                    return param;
            }
            return null;
        }
    }
}

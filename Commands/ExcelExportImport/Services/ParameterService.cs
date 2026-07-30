using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28;
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
                    return param.AsValueString() ?? param.AsElementId().IntValue().ToString();
#endif
                default:
                    return "";
            }
        }

        /// <summary>
        /// パラメータ値を文字列から設定
        /// </summary>
        public static bool SetParameterValue(Parameter param, string value, Document doc)
        {
            if (param == null || param.IsReadOnly)
            {
                DiagLog.Write($"[SetParam] skip: param={(param == null ? "null" : "readonly")} value='{value}'");
                return false;
            }

            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        param.Set(value);
                        return true;

                    case StorageType.Integer:
                        if (int.TryParse(value, out int intVal))
                            return param.Set(intVal);
                        DiagLog.Write($"[SetParam] Integer 解析失敗 '{value}' ({param.Definition?.Name})");
                        return false;

                    case StorageType.Double:
                        // Excel の数値は「表示単位」（mm・㎡ 等）。表示単位→内部単位へ変換して設定する。
                        // 直接 param.Set(数値) すると内部単位(ft 等)として扱われ、例えば部屋の
                        // 上限オフセットに -100 を入れても -100ft 相当になり誤った値になる。
                        if (double.TryParse(value, out double dblVal))
                        {
                            // 1) 単位が取得できるなら 表示単位→内部単位 に変換して設定（最も確実）
                            try
                            {
                                var unitTypeId = param.GetUnitTypeId();
                                if (unitTypeId != null && !unitTypeId.Empty())
                                    return param.Set(UnitUtils.ConvertToInternalUnits(dblVal, unitTypeId));
                            }
                            catch
                            {
                                // 単位が取得できないパラメータ → 下のフォールバックへ
                            }

                            // 2) 単位不明時はまず表示文字列として設定（Revit に単位を解釈させる）
                            if (param.SetValueString(value))
                                return true;

                            // 3) それも不可なら単位を持たない数値とみなして生値で設定（最後の手段）
                            return param.Set(dblVal);
                        }
                        // 数値として解釈できない（書式付き文字列等）は表示文字列として設定を試みる
                        if (param.SetValueString(value))
                            return true;
                        return false;

                    case StorageType.ElementId:
                        return SetElementIdParameter(param, value, doc);

                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[SetParam] 例外 param='{param.Definition?.Name}' storage={param.StorageType} value='{value}': {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// ElementId 型パラメータ（部屋の上部/下部レベル等）を設定する。
        /// エクスポートは要素名（例: レベル名 "1FL(2FL)"）で書き出すため、
        /// 数値ID・名前検索・SetValueString の順に解決する。
        /// </summary>
        private static bool SetElementIdParameter(Parameter param, string value, Document doc)
        {
            // 1) 数値ならそのまま ElementId として設定
            if (int.TryParse(value, out int idVal))
                return param.Set(new ElementId(idVal));

            // 2) 名前から要素を検索して設定（レベル名など）
            var targetId = ResolveElementIdByName(param, value, doc);
            if (targetId != null && targetId != ElementId.InvalidElementId)
            {
                bool ok = param.Set(targetId);
                DiagLog.Write($"[SetParam] ElementId 名前解決 '{value}' -> {targetId} set={ok} ({param.Definition?.Name})");
                return ok;
            }

            // 3) 最後に SetValueString を試す
            try
            {
                if (param.SetValueString(value))
                    return true;
            }
            catch { }

            DiagLog.Write($"[SetParam] ElementId 名前解決失敗 '{value}'（該当要素が見つからない） ({param.Definition?.Name})");
            return false;
        }

        /// <summary>
        /// 名前から要素の ElementId を検索する。対象クラスを優先順に試す:
        /// 現在値の要素クラス → パラメータ種別から推定（イメージ=ImageType 等）→ Level/Material。
        /// エクスポートは ElementId 参照を要素名で書き出すため、その名前から実要素を引き当てる。
        /// </summary>
        private static ElementId ResolveElementIdByName(Parameter param, string name, Document doc)
        {
            if (doc == null || string.IsNullOrEmpty(name))
                return ElementId.InvalidElementId;

            var candidateClasses = new List<Type>();

            // 1) 現在値の要素と同じクラス（値の入れ替えで最も確実）
            try
            {
                var currentId = param.AsElementId();
                if (currentId != null && currentId != ElementId.InvalidElementId)
                {
                    var cur = doc.GetElement(currentId);
                    if (cur != null)
                        candidateClasses.Add(cur.GetType());
                }
            }
            catch { }

            // 2) パラメータ種別から対象クラスを推定
            var bip = GetBuiltInParameter(param);
            if (bip == BuiltInParameter.ALL_MODEL_IMAGE || bip == BuiltInParameter.ALL_MODEL_TYPE_IMAGE)
                candidateClasses.Add(typeof(ImageType)); // 「イメージ」= ラスター画像への参照

            // 3) 汎用の頻出クラス（現在値が未設定でも対応）
            candidateClasses.Add(typeof(Level));
            candidateClasses.Add(typeof(Material));

            foreach (var cls in candidateClasses.Distinct())
            {
                try
                {
                    var hit = new FilteredElementCollector(doc)
                        .OfClass(cls)
                        .FirstOrDefault(e => SafeName(e) == name);
                    if (hit != null)
                        return hit.Id;
                }
                catch { }
            }

            return ElementId.InvalidElementId;
        }

        /// <summary>パラメータの BuiltInParameter を安全に取得（共有/カスタムは INVALID）</summary>
        private static BuiltInParameter GetBuiltInParameter(Parameter param)
        {
            try
            {
                if (param?.Definition is InternalDefinition intDef)
                    return intDef.BuiltInParameter;
            }
            catch { }
            return BuiltInParameter.INVALID;
        }

        /// <summary>Element.Name は一部要素で例外を投げるため安全に取得する</summary>
        private static string SafeName(Element e)
        {
            try { return e?.Name; }
            catch { return null; }
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
            if (param.Id.IntValue() == (int)BuiltInParameter.ELEM_TYPE_PARAM)
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
            // LookupParameter はネイティブの名前引き（全パラメータの線形走査を回避）。
            // 同名が複数ある場合は最初の1件を返す（旧実装と同じ挙動）。
            return elem.LookupParameter(paramName);
        }
    }
}

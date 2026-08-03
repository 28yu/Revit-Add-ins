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

            var list = parameters.ToList();

            // 同名パラメータ（例: エリアの「用途」×2）は名前だけでは区別できないため、
            // 同カテゴリ・同種別(I-/T-)・同名のグループに種別ベースの曖昧性解消接尾辞を付ける。
            AssignDisambiguationSuffixes(list);

            return list.OrderBy(p => p.DisplayName).ToList();
        }

        /// <summary>
        /// 同カテゴリ・同種別(インスタンス/タイプ)・同名で複数存在するパラメータに、
        /// 種別（組み込み/共有/プロジェクト）ラベルの接尾辞を付けて区別できるようにする。
        /// 同一種別でさらに複数ある場合は ParamId 昇順で「#連番」を付与する。
        /// 重複が無いパラメータは接尾辞なし（＝従来どおりの表示名/ヘッダー）。
        /// </summary>
        private static void AssignDisambiguationSuffixes(List<ParameterInfo> list)
        {
            foreach (var group in list.GroupBy(p => (p.IsTypeParameter ? "T" : "I") + "|" + p.RawName))
            {
                var members = group.ToList();
                if (members.Count <= 1)
                    continue; // 重複なし → 接尾辞不要

                foreach (var kindGroup in members.GroupBy(m => m.Kind))
                {
                    var ordered = kindGroup.OrderBy(m => m.ParamId).ToList();
                    bool sameKindCollision = ordered.Count > 1;
                    for (int i = 0; i < ordered.Count; i++)
                    {
                        int index = sameKindCollision ? (i + 1) : 0;
                        ordered[i].DisambigSuffix =
                            ParameterKindHelper.BuildSuffix(ordered[i].Kind, index);
                    }
                }
            }
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

                var bip = GetBuiltInParameter(param);
                bucket.Add(new ParameterInfo(
                    param.Definition.Name,
                    isTypeParameter,
                    param.IsReadOnly,
                    categoryName)
                {
                    // 要素参照型（レベル/材料/タイプ/画像等）＝文字値ではなく既存要素名が必要
                    IsElementReference = param.StorageType == StorageType.ElementId,
                    // 画像参照は文字値を一切設定できない（画像ピッカー専用）
                    IsImage = bip == BuiltInParameter.ALL_MODEL_IMAGE
                              || bip == BuiltInParameter.ALL_MODEL_TYPE_IMAGE,
                    // 同名パラメータの区別用（識別子＝安定 ID、種別＝組み込み/共有/プロジェクト）
                    ParamId = ParamIdToLong(param.Id),
                    Kind = ParameterKindHelper.Determine(param)
                });
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

        /// <summary>Parameter.Id を long 化（バージョン差異を吸収）。同名パラメータの識別子に使う。</summary>
        public static long ParamIdToLong(ElementId id)
        {
            if (id == null) return 0;
#if REVIT2026
            return id.Value;
#else
            return id.IntValue();
#endif
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
            return FindParameter(elem, paramName, isTypeParameter, null, 0, doc);
        }

        /// <summary>
        /// 要素からパラメータを検索する（種別対応版）。
        /// kind が null（曖昧性解消接尾辞が無いヘッダー）の場合は従来どおり名前引き。
        /// kind 指定時は同名候補の中から種別で絞り込み、同種別で複数ある場合は
        /// ParamId 昇順の index（1始まり）で特定する。エクスポート時の連番付与と対応する。
        /// </summary>
        public static Parameter FindParameter(
            Element elem, string paramName, bool isTypeParameter,
            Models.ParameterKind? kind, int index, Document doc)
        {
            var container = GetParameterContainer(elem, isTypeParameter, doc);
            if (container == null)
                return null;

            // 種別指定なし → 従来の高速な名前引き（同名が無い通常ケース・旧Excel互換）
            if (kind == null)
                return FindParameterByName(container, paramName);

            // 種別指定あり → 同名候補を種別で絞り込む
            var matches = new List<Parameter>();
            foreach (Parameter p in container.Parameters)
            {
                if (p?.Definition == null || p.Definition.Name != paramName)
                    continue;
                if (ParameterKindHelper.Determine(p) == kind.Value)
                    matches.Add(p);
            }

            if (matches.Count == 0)
                return FindParameterByName(container, paramName); // 種別一致なし → 名前引きにフォールバック
            if (matches.Count == 1)
                return matches[0];

            matches.Sort((a, b) => ParamIdToLong(a.Id).CompareTo(ParamIdToLong(b.Id)));
            int i = index - 1;
            return (i >= 0 && i < matches.Count) ? matches[i] : matches[0];
        }

        /// <summary>
        /// エクスポート用: ParameterInfo が保持する安定 ID で、要素上の該当パラメータを厳密に特定する。
        /// 同名パラメータがあっても正しい方の値を書き出せる。ID 不一致時は同名の先頭にフォールバック。
        /// </summary>
        public static Parameter FindByIdentity(Element container, string rawName, long paramId)
        {
            if (container == null) return null;

            Parameter fallback = null;
            foreach (Parameter p in container.Parameters)
            {
                if (p?.Definition == null || p.Definition.Name != rawName)
                    continue;
                if (ParamIdToLong(p.Id) == paramId)
                    return p;
                if (fallback == null)
                    fallback = p;
            }
            return fallback;
        }

        /// <summary>インスタンス/タイプに応じたパラメータの入れ物（要素本体 or タイプ要素）を返す。</summary>
        private static Element GetParameterContainer(Element elem, bool isTypeParameter, Document doc)
        {
            if (!isTypeParameter)
                return elem;

            var typeId = elem?.GetTypeId();
            if (typeId == null || typeId == ElementId.InvalidElementId)
                return null;
            return doc.GetElement(typeId);
        }

        /// <summary>
        /// Excel ヘッダー由来の表示名（マーカー除去済み）を、生パラメータ名・種別・連番へ分解する。
        /// 例: <c>I-用途【共有】</c> → RawName="用途", IsType=false, Kind=Shared。
        /// </summary>
        public struct ParsedParamName
        {
            public string RawName;
            public bool IsTypeParameter;
            public Models.ParameterKind? Kind;
            public int Index;
        }

        /// <summary>表示名（マーカー除去済み）を分解する。曖昧性解消接尾辞が無ければ Kind=null。</summary>
        public static ParsedParamName ParseDisplayName(string displayNameNoMarker)
        {
            var result = new ParsedParamName { RawName = displayNameNoMarker ?? "", Index = 0 };
            string s = result.RawName;

            result.IsTypeParameter = s.StartsWith("T-");
            if (s.StartsWith("T-") || s.StartsWith("I-"))
                s = s.Substring(2);

            if (ParameterKindHelper.TryExtractSuffix(s, out string baseName, out var kind, out int index))
            {
                s = baseName;
                result.Kind = kind;
                result.Index = index;
            }

            result.RawName = s;
            return result;
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

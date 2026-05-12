using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// プロジェクト内のパラメータ候補を、名前に含まれるキーワードでフィルタして列挙する。
    /// 工区別・型枠種別の集計時にユーザーがパラメータ名を選択するためのコンボボックス候補。
    /// </summary>
    internal static class ParameterCandidateScanner
    {
        /// <summary>工区関連のキーワード (パラメータ名にいずれかを含めば候補化)。</summary>
        internal static readonly string[] ZoneKeywords = new[]
        {
            "工区", "ゾーン", "Zone", "エリア", "Area", "ブロック", "Block",
            "区分", "Section", "範囲", "Phase",
        };

        /// <summary>型枠種別関連のキーワード。</summary>
        internal static readonly string[] FormworkTypeKeywords = new[]
        {
            "型枠", "種別", "Formwork", "Type", "パターン", "Pattern", "仕様", "Spec",
        };

        /// <summary>
        /// 文書内のパラメータ名のうち、いずれかのキーワードを名前に含むものを返す。
        /// プロジェクト/共有パラメータバインドと、各カテゴリの先頭インスタンス・タイプから収集する。
        /// </summary>
        internal static List<string> FindCandidates(Document doc, string[] keywords)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (doc == null || keywords == null || keywords.Length == 0) return new List<string>();

            // 1. ParameterBindings (プロジェクトパラメータ・共有パラメータ)
            try
            {
                var bindings = doc.ParameterBindings;
                var it = bindings.ForwardIterator();
                while (it.MoveNext())
                {
                    var def = it.Key;
                    if (def == null) continue;
                    var name = def.Name;
                    if (NameMatches(name, keywords)) names.Add(name);
                }
            }
            catch { }

            // 2. 主要カテゴリの先頭インスタンス + そのタイプから収集
            //    (組込みパラメータも含むため網羅性が増す)
            var cats = new[]
            {
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_Stairs,
                BuiltInCategory.OST_Roofs,
                BuiltInCategory.OST_Columns,
                BuiltInCategory.OST_GenericModel,
            };
            foreach (var cat in cats)
            {
                try
                {
                    var instances = new FilteredElementCollector(doc)
                        .OfCategory(cat)
                        .WhereElementIsNotElementType()
                        .Take(3)
                        .ToList();
                    foreach (var elem in instances)
                    {
                        CollectFrom(elem, keywords, names);
                        try
                        {
                            var typeId = elem.GetTypeId();
                            if (typeId != null && typeId != ElementId.InvalidElementId)
                            {
                                var typeElem = doc.GetElement(typeId);
                                if (typeElem != null) CollectFrom(typeElem, keywords, names);
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }

            return names.OrderBy(n => n).ToList();
        }

        private static void CollectFrom(Element elem, string[] keywords, HashSet<string> names)
        {
            if (elem == null) return;
            try
            {
                foreach (Parameter p in elem.Parameters)
                {
                    try
                    {
                        var def = p?.Definition;
                        if (def == null) continue;
                        string name = def.Name;
                        if (NameMatches(name, keywords)) names.Add(name);
                    }
                    catch { }
                }
            }
            catch { }
        }

        private static bool NameMatches(string name, string[] keywords)
        {
            if (string.IsNullOrEmpty(name)) return false;
            foreach (var kw in keywords)
            {
                if (string.IsNullOrEmpty(kw)) continue;
                if (name.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
    }
}

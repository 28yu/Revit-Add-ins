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
        public static List<ParameterInfo> GetParametersForCategory(Document doc, BuiltInCategory category, string categoryName)
        {
            var parameters = new HashSet<ParameterInfo>();

            var elements = new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .Take(50) // パラメータ列挙のため少数サンプルで十分
                .ToList();

            foreach (var elem in elements)
            {
                // インスタンスパラメータ
                foreach (Parameter param in elem.Parameters)
                {
                    if (param.Definition == null || string.IsNullOrEmpty(param.Definition.Name))
                        continue;

                    var info = new ParameterInfo(
                        param.Definition.Name,
                        false,
                        param.IsReadOnly,
                        categoryName);
                    parameters.Add(info);
                }

                // タイプパラメータ
                var typeId = elem.GetTypeId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    var elemType = doc.GetElement(typeId);
                    if (elemType != null)
                    {
                        foreach (Parameter param in elemType.Parameters)
                        {
                            if (param.Definition == null || string.IsNullOrEmpty(param.Definition.Name))
                                continue;

                            var info = new ParameterInfo(
                                param.Definition.Name,
                                true,
                                param.IsReadOnly,
                                categoryName);
                            parameters.Add(info);
                        }
                    }
                }
            }

            return parameters.OrderBy(p => p.DisplayName).ToList();
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
                    return param.AsValueString() ?? param.AsElementId().IntegerValue.ToString();
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

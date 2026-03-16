using System;
using System.IO;
using System.Text;
using Autodesk.Revit.DB;
using RevitApp = Autodesk.Revit.ApplicationServices.Application;

namespace Tools28.Commands.BeamTopLevel
{
    /// <summary>
    /// 梁天端色分け用の共有パラメータ管理
    /// </summary>
    public static class ParameterManager
    {
        public const string ParamRefLevel = "梁天端_基準レベル";
        public const string ParamLevelDiff = "梁天端_レベル差";
        public const string ParamDisplay = "梁天端_表示";
        public const string ParamError = "梁天端_エラー";

        private const string SharedParamGroupName = "Tools28_梁天端";
        private const string SharedParamFileName = "Tools28_BeamTopLevel.txt";

        /// <summary>
        /// 共有パラメータが存在することを確認し、なければ作成・バインドする
        /// </summary>
        public static void EnsureSharedParameters(Document doc, RevitApp app)
        {
            if (IsParameterBound(doc, ParamDisplay))
                return;

            string originalFilePath = app.SharedParametersFilename;

            try
            {
                string tempFilePath = GetSharedParamFilePath();
                CreateSharedParameterFile(tempFilePath);
                app.SharedParametersFilename = tempFilePath;

                DefinitionFile defFile = app.OpenSharedParameterFile();
                if (defFile == null)
                    throw new Exception("共有パラメータファイルを開けませんでした。");

                DefinitionGroup defGroup = defFile.Groups.get_Item(SharedParamGroupName);
                if (defGroup == null)
                    defGroup = defFile.Groups.Create(SharedParamGroupName);

                CategorySet catSet = app.Create.NewCategorySet();
                Category structuralFramingCat = doc.Settings.Categories.get_Item(
                    BuiltInCategory.OST_StructuralFraming);
                catSet.Insert(structuralFramingCat);

                InstanceBinding binding = app.Create.NewInstanceBinding(catSet);
                BindingMap bindingMap = doc.ParameterBindings;

                CreateAndBindTextParameter(defGroup, bindingMap, binding, ParamRefLevel);
                CreateAndBindLengthParameter(defGroup, bindingMap, binding, ParamLevelDiff);
                CreateAndBindTextParameter(defGroup, bindingMap, binding, ParamDisplay);
                CreateAndBindTextParameter(defGroup, bindingMap, binding, ParamError);
            }
            finally
            {
                try
                {
                    app.SharedParametersFilename = originalFilePath ?? "";
                }
                catch
                {
                    // 復元失敗は無視
                }
            }
        }

        /// <summary>
        /// 梁にパラメータ値を書き込み
        /// </summary>
        public static void WriteValues(FamilyInstance beam, BeamCalculationResult result)
        {
            if (result.Success)
            {
                SetParameterValue(beam, ParamRefLevel, result.RefLevelName);
                SetParameterLength(beam, ParamLevelDiff, result.LevelDifference);
                SetParameterValue(beam, ParamDisplay, result.DisplayValue);
                SetParameterValue(beam, ParamError, "");
            }
            else
            {
                SetParameterValue(beam, ParamRefLevel, "");
                SetParameterLength(beam, ParamLevelDiff, 0);
                SetParameterValue(beam, ParamDisplay, "");
                SetParameterValue(beam, ParamError, result.Error ?? "不明なエラー");
            }
        }

        private static bool IsParameterBound(Document doc, string paramName)
        {
            BindingMap bindingMap = doc.ParameterBindings;
            DefinitionBindingMapIterator iter = bindingMap.ForwardIterator();
            while (iter.MoveNext())
            {
                if (iter.Key.Name == paramName)
                    return true;
            }
            return false;
        }

        private static string GetSharedParamFilePath()
        {
            string appDataPath = Environment.GetFolderPath(
                Environment.SpecialFolder.ApplicationData);
            string toolsDir = Path.Combine(appDataPath, "Tools28");

            if (!Directory.Exists(toolsDir))
                Directory.CreateDirectory(toolsDir);

            return Path.Combine(toolsDir, SharedParamFileName);
        }

        private static void CreateSharedParameterFile(string filePath)
        {
            if (File.Exists(filePath))
                return;

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("# This is a Revit shared parameter file.");
            sb.AppendLine("# Do not edit manually.");
            sb.AppendLine("*META\tVERSION\tMINVERSION");
            sb.AppendLine("META\t2\t1");
            sb.AppendLine("*GROUP\tID\tNAME");
            sb.AppendLine("*PARAM\tGUID\tNAME\tDATATYPE\tDATACAT\tGROUP\tDESCRIPTION\tUSERMODIFIABLE\tVISIBLE");

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        private static void CreateAndBindTextParameter(
            DefinitionGroup defGroup,
            BindingMap bindingMap,
            InstanceBinding binding,
            string paramName)
        {
            Definition definition = defGroup.Definitions.get_Item(paramName);
            if (definition == null)
            {
                var options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.String.Text);
                options.Visible = true;
                definition = defGroup.Definitions.Create(options);
            }

            BindParameter(bindingMap, definition, binding);
        }

        private static void CreateAndBindLengthParameter(
            DefinitionGroup defGroup,
            BindingMap bindingMap,
            InstanceBinding binding,
            string paramName)
        {
            Definition definition = defGroup.Definitions.get_Item(paramName);
            if (definition == null)
            {
                var options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Length);
                options.Visible = true;
                definition = defGroup.Definitions.Create(options);
            }

            BindParameter(bindingMap, definition, binding);
        }

        private static void BindParameter(
            BindingMap bindingMap,
            Definition definition,
            InstanceBinding binding)
        {
            if (!bindingMap.Contains(definition))
            {
                bindingMap.Insert(definition, binding);
            }
        }

        private static void SetParameterValue(Element elem, string paramName, string value)
        {
            Parameter param = elem.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly)
                param.Set(value ?? "");
        }

        private static void SetParameterLength(Element elem, string paramName, double value)
        {
            Parameter param = elem.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly)
                param.Set(value);
        }
    }
}

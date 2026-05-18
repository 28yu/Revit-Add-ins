using System;
using System.IO;
using System.Text;
using Autodesk.Revit.DB;
using RevitApp = Autodesk.Revit.ApplicationServices.Application;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 型枠数量算出 DirectShape 用の共有パラメータ管理。
    /// OST_GenericModel（DirectShape が属するカテゴリ）にバインドする。
    /// </summary>
    internal static class FormworkParameterManager
    {
        public const string ParamMarker = "28Tools_FormworkMarker";   // 識別用マーカー（"28Tools_Formwork" を書き込む）
        public const string ParamCategory = "28Tools_Formwork_部位";   // 柱 / 梁 / 壁 / ...
        public const string ParamLevel = "28Tools_Formwork_レベル";    // 参照レベル名
        public const string ParamGroupKey = "28Tools_Formwork_区分";   // 部位 or 工区 or 型枠種別の値
        public const string ParamArea = "28Tools_Formwork_面積";       // 面積（㎡, Length^2）
        public const string ParamPartialContact = "28Tools_Formwork_部分接触"; // "Yes"/"No" - T字結合等で主面の一部が他要素に接触
        public const string ParamSource = "28Tools_Formwork_ソース";    // "ホスト" or リンクファイル名
        public const string ParamSourceView = "28Tools_Formwork_ソースビュー"; // 算出元の3Dビュー名

        // ビュー/集計表/シート用の出力タグ。リネームされても追跡可能にする。
        public const string ParamRelatedSourceView = "28Tools_Formwork_関連ソースビュー"; // 紐づくソースビュー名
        public const string ParamOutputKind = "28Tools_Formwork_出力種別";           // OutputKind* 定数の値

        public const string OutputKindAnalysisView = "AnalysisView";
        public const string OutputKindMainSchedule = "MainSchedule";
        public const string OutputKindSummarySchedule = "SummarySchedule";
        public const string OutputKindSheet = "FormworkSheet";

        public const string MarkerValue = "28Tools_Formwork";
        // 型枠不要として除外された要素 DirectShape のマーカー値。集計表は MarkerValue のみを
        // 対象とし、この値を持つ DirectShape は集計から除外される。
        // クリーンアップは "28Tools_Formwork" で始まる全マーカーを対象とする。
        public const string MarkerValueExcluded = "28Tools_Formwork_Excluded";

        // 部位ラベル（種別ごと）。集計表に表示される（ただし現状は MarkerValue でフィルタされ非表示）。
        public const string SteelExcludedLabel = "鉄骨(除外)";
        public const string DeckSlabExcludedLabel = "デッキスラブ(除外)";
        public const string WallSweepExcludedLabel = "壁スイープ(除外)";
        public const string SteelStairExcludedLabel = "鉄骨階段(除外)";
        public const string LgsWallExcludedLabel = "LGS壁(除外)";

        // View Filter のグループキー。除外要素は全て同じキー・色・フィルタを共有し、
        // 解析ビュー上では既定で非表示（ユーザーが手動で ON にすると確認可能）。
        public const string ExcludedGroupKey = "除外";

        private const string SharedParamGroupName = "Tools28_Formwork";
        private const string SharedParamFileName = "Tools28_Formwork.txt";

        internal static void EnsureParameters(Document doc, RevitApp app)
        {
            if (IsAllBound(doc)) return;

            string originalFilePath = app.SharedParametersFilename;
            try
            {
                string tempFilePath = GetSharedParamFilePath();
                EnsureSharedParameterFile(tempFilePath);
                app.SharedParametersFilename = tempFilePath;

                DefinitionFile defFile = app.OpenSharedParameterFile();
                if (defFile == null)
                    throw new Exception("共有パラメータファイルを開けませんでした。");

                DefinitionGroup defGroup = defFile.Groups.get_Item(SharedParamGroupName)
                    ?? defFile.Groups.Create(SharedParamGroupName);

                CategorySet catSet = app.Create.NewCategorySet();
                var cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel);
                if (cat != null) catSet.Insert(cat);

                var binding = app.Create.NewInstanceBinding(catSet);
                var bindingMap = doc.ParameterBindings;

                CreateAndBindText(defGroup, bindingMap, binding, ParamMarker);
                CreateAndBindText(defGroup, bindingMap, binding, ParamCategory);
                CreateAndBindText(defGroup, bindingMap, binding, ParamLevel);
                CreateAndBindText(defGroup, bindingMap, binding, ParamGroupKey);
                CreateAndBindArea(defGroup, bindingMap, binding, ParamArea);
                CreateAndBindText(defGroup, bindingMap, binding, ParamPartialContact);
                CreateAndBindText(defGroup, bindingMap, binding, ParamSource);
                CreateAndBindText(defGroup, bindingMap, binding, ParamSourceView);

                // 出力タグ (ビュー / 集計表 / シート向け)。リネーム後も追跡できるようにする。
                CategorySet viewCatSet = app.Create.NewCategorySet();
                foreach (var bic in new[]
                {
                    BuiltInCategory.OST_Views,
                    BuiltInCategory.OST_Schedules,
                    BuiltInCategory.OST_Sheets,
                })
                {
                    try
                    {
                        var vc = doc.Settings.Categories.get_Item(bic);
                        if (vc != null) viewCatSet.Insert(vc);
                    }
                    catch { }
                }
                var viewBinding = app.Create.NewInstanceBinding(viewCatSet);
                CreateAndBindText(defGroup, bindingMap, viewBinding, ParamRelatedSourceView);
                CreateAndBindText(defGroup, bindingMap, viewBinding, ParamOutputKind);
            }
            finally
            {
                try { app.SharedParametersFilename = originalFilePath ?? ""; } catch { }
            }
        }

        internal static void SetInstanceValues(
            Element ds, string categoryLabel, string levelName, string groupKey, double areaM2,
            bool hasPartialContact = false, string sourceName = null, string sourceViewName = null)
        {
            SetInstanceValues(ds, MarkerValue, categoryLabel, levelName, groupKey, areaM2,
                hasPartialContact, sourceName, sourceViewName);
        }

        internal static void SetInstanceValues(
            Element ds, string markerValue,
            string categoryLabel, string levelName, string groupKey, double areaM2,
            bool hasPartialContact = false, string sourceName = null, string sourceViewName = null)
        {
            SetString(ds, ParamMarker, markerValue ?? MarkerValue);
            SetString(ds, ParamCategory, categoryLabel ?? string.Empty);
            SetString(ds, ParamLevel, levelName ?? string.Empty);
            SetString(ds, ParamGroupKey, groupKey ?? string.Empty);

            double feetSq = UnitUtils.ConvertToInternalUnits(areaM2, UnitTypeId.SquareMeters);
            SetDouble(ds, ParamArea, feetSq);

            SetString(ds, ParamPartialContact, hasPartialContact ? "Yes" : "No");
            SetString(ds, ParamSource, sourceName ?? "ホスト");
            SetString(ds, ParamSourceView, sourceViewName ?? string.Empty);
        }

        private static bool IsAllBound(Document doc)
        {
            var map = doc.ParameterBindings;
            var it = map.ForwardIterator();
            bool m = false, c = false, l = false, g = false, a = false, p = false, s = false, sv = false,
                 rsv = false, ok = false;
            while (it.MoveNext())
            {
                var name = it.Key.Name;
                if (name == ParamMarker) m = true;
                else if (name == ParamCategory) c = true;
                else if (name == ParamLevel) l = true;
                else if (name == ParamGroupKey) g = true;
                else if (name == ParamArea) a = true;
                else if (name == ParamPartialContact) p = true;
                else if (name == ParamSource) s = true;
                else if (name == ParamSourceView) sv = true;
                else if (name == ParamRelatedSourceView) rsv = true;
                else if (name == ParamOutputKind) ok = true;
            }
            return m && c && l && g && a && p && s && sv && rsv && ok;
        }

        /// <summary>
        /// ビュー / 集計表 / シートに出力タグを書き込む。
        /// outputKind: OutputKindAnalysisView / OutputKindMainSchedule / OutputKindSummarySchedule / OutputKindSheet
        /// relatedSourceView: 紐づくソースビュー名 (合計表・シートでは空文字)
        /// </summary>
        internal static void SetOutputTag(Element e, string outputKind, string relatedSourceView = "")
        {
            if (e == null) return;
            SetString(e, ParamOutputKind, outputKind ?? string.Empty);
            SetString(e, ParamRelatedSourceView, relatedSourceView ?? string.Empty);
        }

        /// <summary>要素から出力種別タグを取得する。タグが無ければ空文字。</summary>
        internal static string GetOutputKind(Element e)
        {
            if (e == null) return string.Empty;
            try
            {
                var p = e.LookupParameter(ParamOutputKind);
                return (p != null && p.StorageType == StorageType.String)
                    ? (p.AsString() ?? string.Empty) : string.Empty;
            }
            catch { return string.Empty; }
        }

        /// <summary>要素から紐づくソースビュー名タグを取得する。タグが無ければ空文字。</summary>
        internal static string GetRelatedSourceView(Element e)
        {
            if (e == null) return string.Empty;
            try
            {
                var p = e.LookupParameter(ParamRelatedSourceView);
                return (p != null && p.StorageType == StorageType.String)
                    ? (p.AsString() ?? string.Empty) : string.Empty;
            }
            catch { return string.Empty; }
        }

        private static string GetSharedParamFilePath()
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string toolsDir = Path.Combine(appDataPath, "Tools28");
            if (!Directory.Exists(toolsDir)) Directory.CreateDirectory(toolsDir);
            return Path.Combine(toolsDir, SharedParamFileName);
        }

        private static void EnsureSharedParameterFile(string filePath)
        {
            if (File.Exists(filePath)) return;
            var sb = new StringBuilder();
            sb.AppendLine("# This is a Revit shared parameter file.");
            sb.AppendLine("# Do not edit manually.");
            sb.AppendLine("*META\tVERSION\tMINVERSION");
            sb.AppendLine("META\t2\t1");
            sb.AppendLine("*GROUP\tID\tNAME");
            sb.AppendLine("*PARAM\tGUID\tNAME\tDATATYPE\tDATACAT\tGROUP\tDESCRIPTION\tUSERMODIFIABLE\tVISIBLE");
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        private static void CreateAndBindText(
            DefinitionGroup grp, BindingMap map, InstanceBinding binding, string name)
        {
            Definition def = grp.Definitions.get_Item(name);
            if (def == null)
            {
#if REVIT2021
                var opts = new ExternalDefinitionCreationOptions(name, ParameterType.Text) { Visible = true };
#else
                var opts = new ExternalDefinitionCreationOptions(name, SpecTypeId.String.Text) { Visible = true };
#endif
                def = grp.Definitions.Create(opts);
            }
            if (!map.Contains(def)) map.Insert(def, binding);
        }

        private static void CreateAndBindArea(
            DefinitionGroup grp, BindingMap map, InstanceBinding binding, string name)
        {
            Definition def = grp.Definitions.get_Item(name);
            if (def == null)
            {
#if REVIT2021
                var opts = new ExternalDefinitionCreationOptions(name, ParameterType.Area) { Visible = true };
#else
                var opts = new ExternalDefinitionCreationOptions(name, SpecTypeId.Area) { Visible = true };
#endif
                def = grp.Definitions.Create(opts);
            }
            if (!map.Contains(def)) map.Insert(def, binding);
        }

        private static void SetString(Element e, string paramName, string value)
        {
            var p = e.LookupParameter(paramName);
            if (p != null && !p.IsReadOnly && p.StorageType == StorageType.String)
                p.Set(value ?? "");
        }

        private static void SetDouble(Element e, string paramName, double value)
        {
            var p = e.LookupParameter(paramName);
            if (p != null && !p.IsReadOnly && p.StorageType == StorageType.Double)
                p.Set(value);
        }
    }
}

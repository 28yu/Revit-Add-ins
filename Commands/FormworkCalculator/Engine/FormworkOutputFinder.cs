using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Output;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 既存の型枠出力 (分析3Dビュー / 集計表 / シート) を検索するユーティリティ。
    ///
    /// 検索戦略 (2段階):
    ///   1. 出力タグ (ParamOutputKind + ParamRelatedSourceView) によるリネーム耐性検索
    ///   2. 名前パターンによるフォールバック検索 (旧出力 / タグ未付与のもの)
    /// </summary>
    internal static class FormworkOutputFinder
    {
        /// <summary>指定ソースビュー名に対応する分析3Dビューを返す (タグ→名前 の順で検索)。</summary>
        internal static View3D FindAnalysisView(Document doc, string sourceViewName)
        {
            if (doc == null) return null;
            var all = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate)
                .ToList();

            // 1. タグ検索
            var tagged = all.FirstOrDefault(v =>
                FormworkParameterManager.GetOutputKind(v) == FormworkParameterManager.OutputKindAnalysisView
                && FormworkParameterManager.GetRelatedSourceView(v) == (sourceViewName ?? string.Empty));
            if (tagged != null) return tagged;

            // 2. 名前フォールバック (新旧全パターン)
            string analysisName = FormworkVisualizer.BuildAnalysisViewName(sourceViewName);
            string legacyName = string.IsNullOrEmpty(sourceViewName)
                ? FormworkVisualizer.LegacyAnalysisViewName
                : FormworkVisualizer.LegacyAnalysisViewPrefix + sourceViewName;
            string legacy2Name = string.IsNullOrEmpty(sourceViewName)
                ? FormworkVisualizer.Legacy2AnalysisViewName
                : FormworkVisualizer.Legacy2AnalysisViewPrefix + sourceViewName;
            return all.FirstOrDefault(v => v.Name == analysisName || v.Name == legacyName || v.Name == legacy2Name);
        }

        /// <summary>プロジェクト内の全分析3Dビュー (タグ + 名前パターン) を返す。</summary>
        internal static List<View3D> FindAllAnalysisViews(Document doc)
        {
            if (doc == null) return new List<View3D>();
            var all = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate)
                .ToList();
            return all.Where(v =>
                FormworkParameterManager.GetOutputKind(v) == FormworkParameterManager.OutputKindAnalysisView
                || FormworkVisualizer.IsAnalysisViewName(v.Name))
                .OrderBy(v => v.Name)
                .ToList();
        }

        /// <summary>指定ソースビュー名に対応するビュー別集計表を返す。</summary>
        internal static ViewSchedule FindMainSchedule(Document doc, string sourceViewName)
        {
            if (doc == null) return null;
            var all = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(v => !v.IsTemplate)
                .ToList();

            var tagged = all.FirstOrDefault(v =>
                FormworkParameterManager.GetOutputKind(v) == FormworkParameterManager.OutputKindMainSchedule
                && FormworkParameterManager.GetRelatedSourceView(v) == (sourceViewName ?? string.Empty));
            if (tagged != null) return tagged;

            string name = string.IsNullOrEmpty(sourceViewName)
                ? ScheduleCreator.ScheduleName
                : $"{ScheduleCreator.ScheduleName} - {sourceViewName}";
            return all.FirstOrDefault(v => v.Name == name);
        }

        /// <summary>合計集計表 (型枠数量集計_合計) を返す。</summary>
        internal static ViewSchedule FindSummarySchedule(Document doc)
        {
            if (doc == null) return null;
            var all = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(v => !v.IsTemplate)
                .ToList();

            var tagged = all.FirstOrDefault(v =>
                FormworkParameterManager.GetOutputKind(v) == FormworkParameterManager.OutputKindSummarySchedule);
            if (tagged != null) return tagged;

            return all.FirstOrDefault(v => v.Name == ScheduleCreator.SummaryScheduleName);
        }

        /// <summary>型枠集計シートを返す。</summary>
        internal static ViewSheet FindFormworkSheet(Document doc)
        {
            if (doc == null) return null;
            var all = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => !s.IsTemplate && !s.IsPlaceholder)
                .ToList();

            var tagged = all.FirstOrDefault(s =>
                FormworkParameterManager.GetOutputKind(s) == FormworkParameterManager.OutputKindSheet);
            if (tagged != null) return tagged;

            return all.FirstOrDefault(s =>
                s.Name == FormworkSheetCreator.SheetName
                || s.Name.StartsWith(FormworkSheetCreator.SheetName + " "));
        }
    }
}

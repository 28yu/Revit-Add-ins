using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.GenericModelMerge.Models;
using Tools28.Commands.GenericModelMerge.Services;
using Tools28.Commands.GenericModelMerge.Views;
using Tools28.Localization;

namespace Tools28.Commands.GenericModelMerge
{
    /// <summary>
    /// アクティブ 3D ビューに表示されている複数カテゴリの要素から形状を読み取り、
    /// 一つの一般モデル（ダイレクトシェイプ または 一般モデルファミリ）を作成するコマンド。
    ///
    /// ビュー内の要素を対象にするため、セクションボックス・ビューフィルタ・要素の非表示で
    /// 事前に範囲を絞れる。さらにダイアログのカテゴリチェックリストで絞り込む。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class GenericModelMergeCommand : IExternalCommand
    {
        /// <summary>これを超える要素数のときは、時間がかかる旨を確認する。</summary>
        private const int ConfirmThreshold = 3000;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp?.ActiveUIDocument;
            Document doc = uidoc?.Document;

            if (doc == null)
            {
                message = Loc.S("GmMerge.Err.NoDoc");
                return Result.Cancelled;
            }

            // 立体形状を扱う機能なので、範囲が直感的に決まる 3D ビュー限定にする
            var view = doc.ActiveView as View3D;
            if (view == null || view.IsTemplate)
            {
                TaskDialog.Show(Loc.S("GmMerge.Title"), Loc.S("GmMerge.Err.Not3DView"));
                return Result.Cancelled;
            }

            try
            {
                DiagLog.Cmd("GenericModelMerge", $"開始 view={view.Name}");

                var categories = ViewElementScanner.ScanCategories(doc, view);
                if (categories.Count == 0)
                {
                    TaskDialog.Show(Loc.S("GmMerge.Title"), Loc.S("GmMerge.Err.NoElement"));
                    return Result.Cancelled;
                }

                var materials = ViewElementScanner.ListMaterials(doc);

                // 既定名は「モデルに保存されるが検索キーではない」文字列（分類B）なので、
                // アドインの言語設定ではなく Revit 本体の UI 言語に合わせる。
                string modelLang = RevitUiLanguage.Resolve(uiapp);

                var dialog = new GenericModelMergeDialog(
                    categories, materials, GetProjectDirectory(doc), modelLang);
                dialog.SetRevitOwner(commandData);

                if (dialog.ShowDialog() != true) return Result.Cancelled;

                var options = dialog.Options;
                var targetIds = dialog.TargetElementIds;

                if (targetIds.Count > ConfirmThreshold)
                {
                    var confirm = TaskDialog.Show(
                        Loc.S("GmMerge.Title"),
                        string.Format(Loc.S("GmMerge.Confirm.Many"), targetIds.Count),
                        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
                    if (confirm != TaskDialogResult.Yes) return Result.Cancelled;
                }

                return Run(commandData, doc, view, options, targetIds);
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[CMD:GenericModelMerge] 例外: {ex}");
                message = string.Format(Loc.S("Common.ErrorWithManual"), ex.Message);
                return Result.Failed;
            }
        }

        private Result Run(
            ExternalCommandData commandData, Document doc, View3D view,
            MergeOptions options, List<ElementId> targetIds)
        {
            // --- 形状の読み取りと結合（読み取りのみ。トランザクションは不要） ---
            var merged = SolidMerger.Build(doc, targetIds, options.CombineMode);
            DiagLog.Write($"[CMD:GenericModelMerge] 形状 {merged.Solids.Count} 個 / " +
                          $"元要素 {merged.SourceElementIds.Count} 個 / " +
                          $"形状なし {merged.NoGeometryCount} 個 / " +
                          $"結合失敗 {merged.UnionFailedCount} 個");

            if (merged.Solids.Count == 0)
            {
                TaskDialog.Show(Loc.S("GmMerge.Title"), Loc.S("GmMerge.Err.NoSolid"));
                return Result.Cancelled;
            }

            // --- ファミリ出力の場合は .rfa をここで作る ---
            // Document.SaveAs はトランザクションが開いていると失敗するため、
            // プロジェクトのトランザクションを開く前に済ませる。
            FamilyBuilder.FamilyFileOutcome famOutcome = null;
            if (options.OutputKind == MergeOutputKind.Family)
            {
                string template = FamilyBuilder.FindGenericModelTemplate(commandData.Application.Application);
                if (string.IsNullOrEmpty(template))
                {
                    template = AskTemplatePath(commandData.Application.Application);
                    if (string.IsNullOrEmpty(template))
                    {
                        TaskDialog.Show(Loc.S("GmMerge.Title"), Loc.S("GmMerge.Err.NoTemplate"));
                        return Result.Cancelled;
                    }
                }

                famOutcome = FamilyBuilder.CreateFamilyFile(
                    commandData.Application.Application, doc, merged.Solids,
                    template, options.FamilyPath, options.MaterialId);

                if (!famOutcome.Success)
                {
                    DiagLog.Write($"[CMD:GenericModelMerge] ファミリ作成失敗: {famOutcome.ErrorMessage}");
                    TaskDialog.Show(Loc.S("GmMerge.Title"),
                        string.Format(Loc.S("GmMerge.Err.FamilyFailed"), famOutcome.ErrorMessage));
                    return Result.Failed;
                }
            }

            // --- プロジェクトへの反映 ---
            ElementId createdId = ElementId.InvalidElementId;
            DirectShapeBuilder.BuildOutcome dsOutcome = null;
            int hiddenCount = 0;

            using (var t = new Transaction(doc, Loc.S("GmMerge.Txn.Create")))
            {
                t.Start();

                if (options.OutputKind == MergeOutputKind.Family)
                {
                    createdId = FamilyBuilder.LoadAndPlace(doc, options.FamilyPath);
                    if (createdId == ElementId.InvalidElementId)
                    {
                        t.RollBack();
                        TaskDialog.Show(Loc.S("GmMerge.Title"), Loc.S("GmMerge.Err.LoadFailed"));
                        return Result.Failed;
                    }
                }
                else
                {
                    dsOutcome = DirectShapeBuilder.Create(
                        doc, merged.Solids, options.Name, options.MaterialId);
                    createdId = dsOutcome.ElementId;
                }

                if (options.HideSourceElements)
                    hiddenCount = HideSourceElements(doc, view, merged.SourceElementIds, createdId);

                t.Commit();
            }

            ShowResult(merged, options, dsOutcome, famOutcome, hiddenCount);
            DiagLog.Cmd("GenericModelMerge", "完了");
            return Result.Succeeded;
        }

        /// <summary>
        /// 元要素をアクティブビューで非表示にする。
        /// 非表示にできない要素（ビューで隠せない種類）は Revit が例外を投げるため、
        /// CanBeHidden で事前にふるい落としてから渡す。生成した要素は当然除外する。
        /// </summary>
        private static int HideSourceElements(
            Document doc, View view, List<ElementId> sourceIds, ElementId createdId)
        {
            var hideable = new List<ElementId>();
            foreach (var id in sourceIds)
            {
                if (id == createdId) continue;
                var e = doc.GetElement(id);
                if (e == null) continue;
                try { if (e.CanBeHidden(view)) hideable.Add(id); }
                catch { }
            }

            if (hideable.Count == 0) return 0;

            try
            {
                view.HideElements(hideable);
                return hideable.Count;
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[CMD:GenericModelMerge] 非表示に失敗: {ex.Message}");
                return 0;
            }
        }

        private static void ShowResult(
            MergeResult merged, MergeOptions options,
            DirectShapeBuilder.BuildOutcome dsOutcome,
            FamilyBuilder.FamilyFileOutcome famOutcome,
            int hiddenCount)
        {
            var sb = new StringBuilder();
            sb.AppendLine(string.Format(Loc.S("GmMerge.Result.Created"),
                merged.SourceElementIds.Count, merged.Solids.Count));

            if (options.OutputKind == MergeOutputKind.Family && famOutcome != null)
                sb.AppendLine(string.Format(Loc.S("GmMerge.Result.SavedTo"), famOutcome.SavedPath));

            if (hiddenCount > 0)
                sb.AppendLine(string.Format(Loc.S("GmMerge.Result.Hidden"), hiddenCount));

            if (merged.NoGeometryCount > 0)
                sb.AppendLine(string.Format(Loc.S("GmMerge.Result.NoGeometry"), merged.NoGeometryCount));

            if (merged.UnionFailedCount > 0)
                sb.AppendLine(string.Format(Loc.S("GmMerge.Result.UnionFailed"), merged.UnionFailedCount));

            if (famOutcome != null && famOutcome.FailedSolidCount > 0)
                sb.AppendLine(string.Format(Loc.S("GmMerge.Result.ShapeFailed"), famOutcome.FailedSolidCount));

            bool materialRequested = options.MaterialId != ElementId.InvalidElementId;
            if (materialRequested)
            {
                if (dsOutcome != null && dsOutcome.MaterialSkippedTooManyFaces)
                    sb.AppendLine(Loc.S("GmMerge.Result.MaterialSkipped"));
                else if (dsOutcome != null && !dsOutcome.MaterialApplied)
                    sb.AppendLine(Loc.S("GmMerge.Result.MaterialFailed"));
                else if (famOutcome != null && !famOutcome.MaterialApplied)
                    sb.AppendLine(Loc.S("GmMerge.Result.MaterialFailed"));
            }

            TaskDialog.Show(Loc.S("GmMerge.Title"), sb.ToString().TrimEnd());
        }

        /// <summary>ファミリテンプレートが自動で見つからないときに手動で選んでもらう。</summary>
        private static string AskTemplatePath(Autodesk.Revit.ApplicationServices.Application app)
        {
            TaskDialog.Show(Loc.S("GmMerge.Title"), Loc.S("GmMerge.SelectTemplate.Prompt"));

            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = Loc.S("GmMerge.SelectTemplate.Title"),
                Filter = Loc.S("GmMerge.SelectTemplate.Filter") + " (*.rft)|*.rft",
                CheckFileExists = true,
            };

            try
            {
                string root = app.FamilyTemplatePath;
                if (!string.IsNullOrEmpty(root) && Directory.Exists(root))
                    dlg.InitialDirectory = root;
            }
            catch { }

            return dlg.ShowDialog() == true ? dlg.FileName : null;
        }

        /// <summary>現在のプロジェクトファイルのフォルダ（クラウドモデルなどでは null）。</summary>
        private static string GetProjectDirectory(Document doc)
        {
            try
            {
                string path = doc.PathName;
                if (string.IsNullOrEmpty(path)) return null;
                return Path.GetDirectoryName(path);
            }
            catch { return null; }
        }
    }
}

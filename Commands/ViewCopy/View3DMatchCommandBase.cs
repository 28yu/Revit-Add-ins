using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Localization;

namespace Tools28.Commands.ViewCopy
{
    /// <summary>
    /// 3Dビューの設定（視点・切断ボックス）を「コピー」「ペースト」の1ボタンで扱う共通の流れ。
    ///
    ///   - プロジェクトブラウザで3Dビューを選んでいる → 選んだビューすべてに貼る（従来のペーストと同じ）
    ///   - 開いている3Dビューで、まだ何もコピーしていない → このビューをコピー元として覚える
    ///   - 開いている3Dビューで、コピー済み → 「このビューに貼る／このビューをコピー元にする」を選ぶ
    ///
    /// 視点・切断ボックスで違うのは「何を覚えて、どう当てるか」と画面の文言だけなので、
    /// それを派生クラスで書く。文言のキーは "{KeyPrefix}.xxx"。
    /// </summary>
    public abstract class View3DMatchCommandBase : IExternalCommand
    {
        /// <summary>文言のキーの先頭（例: "ViewMatch"）。</summary>
        protected abstract string KeyPrefix { get; }

        protected abstract bool HasCopied { get; }

        /// <summary>覚えているコピー元の説明（ビュー名など）。</summary>
        protected abstract string SourceDescription { get; }

        /// <summary>3Dビューの設定を覚える。覚えられなかったときは理由を返す（成功なら null）。</summary>
        protected abstract string CopyFrom(View3D view);

        /// <summary>覚えた設定を3Dビューに当てる（トランザクション内で呼ばれる。失敗は例外）。</summary>
        protected abstract void ApplyTo(View3D view);

        private string S(string key) => Loc.S(KeyPrefix + "." + key);

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                View3D active = doc.ActiveView as View3D;
                if (active != null && active.IsTemplate) active = null;

                // プロジェクトブラウザで選んだ3Dビュー。開いているビュー自身は、ダブルクリックで開いた直後に
                // 選択に残ることがあるため、それ「だけ」の選択はプロジェクトブラウザでの指定とみなさない
                var selectedViews = uidoc.Selection.GetElementIds()
                    .Select(id => doc.GetElement(id))
                    .OfType<View3D>()
                    .Where(v => !v.IsTemplate)
                    .ToList();
                if (selectedViews.Any(v => active == null || v.Id != active.Id))
                {
                    if (!HasCopied)
                    {
                        message = S("NothingCopied");
                        return Result.Failed;
                    }
                    Paste(uidoc, doc, selectedViews, active);
                    return Result.Succeeded;
                }

                if (active == null)
                {
                    message = Loc.S("Common.Need3DViewOrSelect");
                    return Result.Failed;
                }

                if (!HasCopied)
                {
                    string error = CopyFrom(active);
                    if (error != null)
                    {
                        message = error;
                        return Result.Failed;
                    }
                    TaskDialog.Show(S("Title"), string.Format(S("Copied"), active.Name));
                    return Result.Succeeded;
                }

                switch (AskPasteOrCopy())
                {
                    case TaskDialogResult.CommandLink1:
                        Paste(uidoc, doc, new List<View3D> { active }, active);
                        return Result.Succeeded;
                    case TaskDialogResult.CommandLink2:
                        string error = CopyFrom(active);
                        if (error != null)
                        {
                            message = error;
                            return Result.Failed;
                        }
                        return Result.Succeeded;
                    default:
                        return Result.Cancelled;
                }
            }
            catch (Exception ex)
            {
                message = string.Format(S("Error"), ex.Message);
                return Result.Failed;
            }
        }

        /// <summary>コピー済みのとき、「このビューに貼る」か「このビューをコピー元にする」かを選んでもらう。</summary>
        private TaskDialogResult AskPasteOrCopy()
        {
            var dialog = new TaskDialog(S("Title"))
            {
                MainInstruction = S("Choose"),
                CommonButtons = TaskDialogCommonButtons.Cancel,
                DefaultButton = TaskDialogResult.CommandLink1,
            };
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                S("Paste"), string.Format(S("PasteSub"), SourceDescription));
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                S("Recopy"), S("RecopySub"));
            return dialog.Show();
        }

        private void Paste(UIDocument uidoc, Document doc, List<View3D> targets, View3D active)
        {
            int success = 0;
            int failed = 0;

            using (Transaction trans = new Transaction(doc, S("Txn")))
            {
                trans.Start();
                foreach (View3D view in targets)
                {
                    try
                    {
                        ApplyTo(view);
                        success++;
                    }
                    catch (Exception ex)
                    {
                        DiagLog.Write($"[{KeyPrefix}] 適用に失敗 view={view.Name}: {ex.Message}");
                        failed++;
                    }
                }
                trans.Commit();
            }

            // 非アクティブなビューは次に開いたときに最新で描かれるので、開いているビューだけ再描画する
            if (active != null && targets.Any(v => v.Id == active.Id))
            {
                try { uidoc.RefreshActiveView(); } catch { }
            }

            // Revit は Result.Succeeded のとき message を表示しないため、部分失敗は明示的に伝える
            if (failed > 0)
            {
                TaskDialog.Show(Loc.S("Common.Warning"),
                    string.Format(Loc.S("Common.PartialFail"), success, failed));
            }
        }
    }
}

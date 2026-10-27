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
    /// Revit の「タイプ プロパティを一致」(MA) に合わせ、コピーした状態はカーソルの横の印（<see cref="CopyBadge"/>）で示す。
    ///
    ///   - 印なし（未コピー）で、開いている3Dビューで押す → このビューをコピー元として覚え、印を出す
    ///   - 印あり で、別の3Dビューで押す → すぐに貼る（印は残るので、続けてほかのビューにも貼れる）
    ///   - 印あり で、コピー元のビューで押す → コピー元を覚え直す
    ///   - プロジェクトブラウザで3Dビューを選んで押す → 選んだビューすべてに貼る
    ///   - Esc → コピーを取りやめ、印を消す
    ///
    /// 視点・切断ボックスで違うのは「何を覚えて、どう当てるか」と文言・印の絵だけなので、それを派生クラスで書く。
    /// 文言のキーは "{KeyPrefix}.xxx"。
    /// </summary>
    public abstract class View3DMatchCommandBase : IExternalCommand
    {
        /// <summary>文言のキーの先頭（例: "ViewMatch"）。印の識別にも使う。</summary>
        protected abstract string KeyPrefix { get; }

        /// <summary>カーソルの横に出す印の画像（Resources/Icons のファイル名）。</summary>
        protected abstract string BadgeIcon { get; }

        protected abstract bool HasCopied { get; }

        /// <summary>覚えているコピー元のビュー名。</summary>
        protected abstract string SourceViewName { get; }

        /// <summary>覚えているコピー元の説明（結果の表示用）。</summary>
        protected abstract string SourceDescription { get; }

        /// <summary>3Dビューの設定を覚える。覚えられなかったときは理由を返す（成功なら null）。</summary>
        protected abstract string CopyFrom(View3D view);

        /// <summary>覚えた設定を3Dビューに当てる（トランザクション内で呼ばれる。失敗は例外）。</summary>
        protected abstract void ApplyTo(View3D view);

        /// <summary>覚えた設定を消す（Esc で取りやめたとき）。</summary>
        protected abstract void ClearCopy();

        private string S(string key) => Loc.S(KeyPrefix + "." + key);

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;
                View3D active = doc.ActiveView as View3D;
                if (active != null && active.IsTemplate) active = null;

                // プロジェクトブラウザで選んだ3Dビュー。開いているビュー自身は、ダブルクリックで開いた直後に
                // 選択に残ることがあるため、それ「だけ」の選択はプロジェクトブラウザでの指定とみなさない
                var selected = uidoc.Selection.GetElementIds().Select(id => doc.GetElement(id)).Where(e => e != null).ToList();
                var selectedViews = selected.OfType<View3D>().Where(v => !v.IsTemplate).ToList();
                DiagLog.Write($"[{KeyPrefix}] 実行 active={active?.Name ?? "(3Dビュー以外)"} " +
                              $"選択={selected.Count}件（3Dビュー {selectedViews.Count}件） コピー済み={HasCopied}");

                if (selectedViews.Any(v => active == null || v.Id != active.Id))
                {
                    if (!HasCopied)
                    {
                        message = S("NothingCopied");
                        return Result.Failed;
                    }
                    int pasted = Paste(uiapp, doc, selectedViews);
                    TaskDialog.Show(S("Title"), string.Format(S("Pasted"), pasted, SourceDescription));
                    return Result.Succeeded;
                }

                if (active == null)
                {
                    message = Loc.S("Common.Need3DViewOrSelect");
                    return Result.Failed;
                }

                // 未コピー、またはコピー元のビューで押した → このビューを（改めて）コピー元にする
                if (!HasCopied || string.Equals(SourceViewName, active.Name, StringComparison.Ordinal))
                {
                    string error = CopyFrom(active);
                    if (error != null)
                    {
                        message = error;
                        return Result.Failed;
                    }
                    CopyBadge.Show(KeyPrefix, BadgeIcon, ClearCopy, uiapp.MainWindowHandle);
                    return Result.Succeeded;
                }

                // コピー済みで別のビュー → すぐに貼る（MA でクリックした要素に写すのと同じ）
                Paste(uiapp, doc, new List<View3D> { active });
                CopyBadge.Show(KeyPrefix, BadgeIcon, ClearCopy, uiapp.MainWindowHandle);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = string.Format(S("Error"), ex.Message);
                return Result.Failed;
            }
        }

        /// <summary>覚えた設定を当てる。当てられた数を返す。</summary>
        private int Paste(UIApplication uiapp, Document doc, List<View3D> targets)
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

            RedrawOpenViews(uiapp.ActiveUIDocument, targets);

            // Revit は Result.Succeeded のとき message を表示しないため、部分失敗は明示的に伝える
            if (failed > 0)
            {
                TaskDialog.Show(Loc.S("Common.Warning"),
                    string.Format(Loc.S("Common.PartialFail"), success, failed));
            }
            return success;
        }

        /// <summary>
        /// 貼った先のうち、画面に開いているビューを描き直す。
        /// ⚠ RefreshActiveView は操作中のビューしか描き直さないため、並べて表示している別のビューは
        ///   貼れていても見た目が変わらず「貼れていない」ように見えた。ズーム範囲を同じ値で設定し直して描き直させる
        /// </summary>
        private void RedrawOpenViews(UIDocument uidoc, List<View3D> targets)
        {
            try
            {
                uidoc.RefreshActiveView();
                var ids = new HashSet<ElementId>(targets.Select(v => v.Id));
                foreach (UIView uiView in uidoc.GetOpenUIViews())
                {
                    if (!ids.Contains(uiView.ViewId)) continue;
                    try
                    {
                        IList<XYZ> corners = uiView.GetZoomCorners();
                        if (corners != null && corners.Count >= 2)
                            uiView.ZoomAndCenterRectangle(corners[0], corners[1]);
                    }
                    catch (Exception ex)
                    {
                        DiagLog.Write($"[{KeyPrefix}] 描き直しに失敗: {ex.Message}");
                    }
                }
            }
            catch
            {
                // 画面の描き直しに失敗しても、設定は貼れている
            }
        }
    }
}

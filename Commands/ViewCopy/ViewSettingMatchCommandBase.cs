using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Localization;

namespace Tools28.Commands.ViewCopy
{
    /// <summary>
    /// ビューの設定（視点・切断ボックス・トリミング領域）を「コピー」「ペースト」の1ボタンで扱う共通の流れ。
    /// Revit の「タイプ プロパティを一致」(MA) に合わせ、コピーした状態はカーソルの横の印（<see cref="CopyBadge"/>）で示す。
    ///
    ///   - 印なし（未コピー）で、開いているビューで押す → このビューをコピー元として覚え、印を出す
    ///   - 印あり で、別のビューで押す → すぐに貼る（印は残るので、続けてほかのビューにも貼れる）
    ///   - 印あり で、コピー元のビューで押す → コピー元を覚え直す
    ///   - プロジェクトブラウザでビューを選んで押す → 選んだビューすべてに貼る
    ///   - Esc → コピーを取りやめ、印を消す
    ///   - 視点と切断ボックス（3Dビュー用）の両方をコピーしている（印が2つ）ときは、どちらのボタンで貼っても両方を貼る。
    ///     トリミング領域は主に3Dビュー以外で使うので、この組み合わせには入れない（<see cref="CombinesWithOthers"/>）
    ///
    /// 種類ごとに違うのは「使えるビュー」「何を覚えて、どう当てるか」と文言・印の絵だけなので、それを派生クラスで書く。
    /// 文言のキーは "{KeyPrefix}.xxx"。
    /// </summary>
    public abstract class ViewSettingMatchCommandBase : IExternalCommand
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

        /// <summary>この種類を使えるビューか（既定は3Dビュー）。</summary>
        protected virtual bool Supports(View view) => view is View3D;

        /// <summary>開いているビューで使えないときの案内。</summary>
        protected virtual string NeedViewMessage(View activeView) => Loc.S("Common.Need3DViewOrSelect");

        /// <summary>視点・切断ボックスのように、ほかの種類と一緒に貼る仲間か。</summary>
        protected virtual bool CombinesWithOthers => true;

        /// <summary>ビューの設定を覚える。覚えられなかったときは理由を返す（成功なら null）。</summary>
        protected abstract string CopyFrom(View view);

        /// <summary>覚えた設定をビューに当てる（トランザクション内で呼ばれる。失敗は例外）。</summary>
        protected abstract void ApplyTo(View view);

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
                View activeView = doc.ActiveView;
                View active = activeView != null && !activeView.IsTemplate && Supports(activeView) ? activeView : null;

                // プロジェクトブラウザで選んだビュー。開いているビュー自身は、ダブルクリックで開いた直後に
                // 選択に残ることがあるため、それ「だけ」の選択はプロジェクトブラウザでの指定とみなさない
                var selected = uidoc.Selection.GetElementIds().Select(id => doc.GetElement(id)).Where(e => e != null).ToList();
                var selectedViews = selected.OfType<View>().Where(v => !v.IsTemplate && Supports(v)).ToList();
                DiagLog.Write($"[{KeyPrefix}] 実行 active={activeView?.Name} 使える={active != null} " +
                              $"選択={selected.Count}件（使えるビュー {selectedViews.Count}件） コピー済み={HasCopied}");

                if (selectedViews.Any(v => active == null || v.Id != active.Id))
                {
                    if (!HasCopied)
                    {
                        message = S("NothingCopied");
                        return Result.Failed;
                    }
                    var kinds = KindsToPaste();
                    int pasted = Paste(uiapp, doc, selectedViews, kinds);
                    TaskDialog.Show(S("Title"), kinds.Count > 1
                        ? string.Format(Loc.S("View3DMatch.PastedBoth"), pasted)
                        : string.Format(S("Pasted"), pasted, SourceDescription));
                    return Result.Succeeded;
                }

                if (active == null)
                {
                    message = NeedViewMessage(activeView);
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
                Paste(uiapp, doc, new List<View> { active }, KindsToPaste());
                CopyBadge.Show(KeyPrefix, BadgeIcon, ClearCopy, uiapp.MainWindowHandle);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = string.Format(S("Error"), ex.Message);
                return Result.Failed;
            }
        }

        /// <summary>
        /// 貼る種類。組み合わせの仲間なら、仲間のうちコピーしている（印が出ている）ものすべて。
        /// そうでなければ押したボタンの種類だけ。押したボタンの種類は必ずコピー済みの状態で呼ぶ。
        /// </summary>
        private List<ViewSettingMatchCommandBase> KindsToPaste()
        {
            if (!CombinesWithOthers) return new List<ViewSettingMatchCommandBase> { this };
            return new List<ViewSettingMatchCommandBase>
            {
                new ViewMatchCommand(),
                new SectionBoxCopy.SectionBoxMatchCommand(),
            }.Where(k => k.CombinesWithOthers && k.HasCopied).ToList();
        }

        /// <summary>
        /// コピーしている種類すべてを当てる（1つの操作なので Ctrl+Z 1回で戻る）。
        /// すべての種類を当てられたビューの数を返す。
        /// </summary>
        private int Paste(UIApplication uiapp, Document doc, List<View> targets, List<ViewSettingMatchCommandBase> kinds)
        {
            int success = 0;
            int failed = 0;
            string txnName = kinds.Count > 1 ? Loc.S("View3DMatch.TxnBoth") : S("Txn");

            // 視点は Revit の Ctrl+Z で戻らないため、貼る前の視点を覚えておく（ViewOrientationUndo）
            if (kinds.Any(k => k is ViewMatchCommand))
            {
                ViewOrientationUndo.EnsureReady(uiapp);
                ViewOrientationUndo.Remember(doc, txnName, targets.OfType<View3D>());
            }

            using (Transaction trans = new Transaction(doc, txnName))
            {
                trans.Start();
                foreach (View view in targets)
                {
                    bool ok = true;
                    foreach (var kind in kinds)
                    {
                        try
                        {
                            kind.ApplyTo(view);
                        }
                        catch (Exception ex)
                        {
                            DiagLog.Write($"[{kind.KeyPrefix}] 適用に失敗 view={view.Name}: {ex.Message}");
                            ok = false;
                        }
                    }
                    if (ok) success++;
                    else failed++;
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
        internal static void RedrawOpenViews(UIDocument uidoc, IEnumerable<View> targets)
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
                        DiagLog.Write($"[ViewSettingMatch] 描き直しに失敗: {ex.Message}");
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

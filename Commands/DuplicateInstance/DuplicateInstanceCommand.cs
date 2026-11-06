using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Tools28.Commands.DuplicateInstance.Models;
using Tools28.Commands.DuplicateInstance.Views;
using Tools28.Localization;

namespace Tools28.Commands.DuplicateInstance
{
    /// <summary>
    /// 「同じ場所に同一のインスタンスがあります」警告の要素を一覧にし、
    /// パラメータを見比べて片方を削除するコマンド。
    /// 「画面で確認」を押すとダイアログを一度閉じ、Revit 上で要素を選択・ズームして、
    /// Esc（またはクリック）でダイアログを開き直す。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DuplicateInstanceCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application?.ActiveUIDocument;
            Document doc = uidoc?.Document;
            if (doc == null)
            {
                message = Loc.S("DupInst.NoDoc");
                return Result.Cancelled;
            }

            try
            {
                var state = new DuplicateDialogState();
                if (!AskScope(uidoc, state)) return Result.Cancelled;
                while (true)
                {
                    state.ShowRequest = null;
                    var dialog = new DuplicateInstanceDialog(doc, state);
                    dialog.SetRevitOwner(commandData);
                    dialog.ShowDialog();

                    if (state.ShowRequest == null || state.ShowRequest.Count == 0) break;
                    ShowInRevit(uidoc, state.ShowRequest);
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = string.Format(Loc.S("Common.ErrorWithManual"), ex.Message);
                return Result.Failed;
            }
        }

        /// <summary>
        /// 最初に、どの範囲の重複を調べるか（モデル全体／現在のビュー／選択している要素）を選んでもらう。
        /// キャンセルなら false。
        /// </summary>
        private static bool AskScope(UIDocument uidoc, DuplicateDialogState state)
        {
            var doc = uidoc.Document;
            var view = doc.ActiveView;
            ICollection<ElementId> selected = null;
            try { selected = uidoc.Selection.GetElementIds(); } catch { }
            int selCount = selected?.Count ?? 0;

            var td = new TaskDialog(Loc.S("DupInst.Title"))
            {
                MainInstruction = Loc.S("DupInst.Scope.Main"),
                // ※ DefaultButton に CommandLink を指定すると、リンクを追加する前の時点で
                //   「対応するボタンがない」例外になる。指定しなくても先頭のリンクが既定になる
                CommonButtons = TaskDialogCommonButtons.Cancel
            };
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, Loc.S("DupInst.Scope.Model"), Loc.S("DupInst.Scope.Model.Desc"));
            if (view != null)
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                    string.Format(Loc.S("DupInst.Scope.View"), view.Name), Loc.S("DupInst.Scope.View.Desc"));
            if (selCount > 0)
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                    string.Format(Loc.S("DupInst.Scope.Selection"), selCount), Loc.S("DupInst.Scope.Selection.Desc"));
            else
                td.FooterText = Loc.S("DupInst.Scope.NoSelection");

            switch (td.Show())
            {
                case TaskDialogResult.CommandLink1:
                    state.ScopeIds = null;
                    state.ScopeLabel = Loc.S("DupInst.Scope.Model");
                    return true;

                case TaskDialogResult.CommandLink2:
                    try
                    {
                        var ids = new HashSet<long>();
                        foreach (var id in new FilteredElementCollector(doc, view.Id)
                                     .WhereElementIsNotElementType().ToElementIds())
                            ids.Add(id.LongValue());
                        state.ScopeIds = ids;
                        state.ScopeLabel = string.Format(Loc.S("DupInst.Scope.View"), view.Name);
                        return true;
                    }
                    catch
                    {
                        // 集計表など、要素を取り出せないビュー
                        TaskDialog.Show(Loc.S("DupInst.Title"), Loc.S("DupInst.Scope.ViewNotSupported"));
                        return false;
                    }

                case TaskDialogResult.CommandLink3:
                    state.ScopeIds = new HashSet<long>(selected.Select(id => id.LongValue()));
                    state.ScopeLabel = string.Format(Loc.S("DupInst.Scope.Selection"), selCount);
                    return true;

                default:
                    return false;
            }
        }

        /// <summary>要素を選択してズームし、利用者が Esc かクリックするまで待つ。</summary>
        private static void ShowInRevit(UIDocument uidoc, List<ElementId> ids)
        {
            var doc = uidoc.Document;
            var alive = new List<ElementId>();
            foreach (var id in ids) if (doc.GetElement(id) != null) alive.Add(id);
            if (alive.Count == 0) return;

            try { uidoc.Selection.SetElementIds(alive); } catch { }

            bool shown = true;
            try { uidoc.ShowElements(alive); }
            catch { shown = false; }

            if (!shown)
            {
                TaskDialog.Show(Loc.S("DupInst.Title"), Loc.S("DupInst.Show.CannotShow"));
                return;
            }

            try
            {
                // 画面の操作（ズーム・回転）は自由にできる。Esc か要素のクリックで戻る
                uidoc.Selection.PickObject(ObjectType.Element, Loc.S("DupInst.Show.Prompt"));
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // Esc で戻った（正常）
            }
            catch
            {
                // 選択待ちができないビュー（集計表など）では、確認用のダイアログで待つ
                TaskDialog.Show(Loc.S("DupInst.Title"), Loc.S("DupInst.Show.Return"));
            }

            // 戻ったあとも、確認した要素は選択したままにする
            try { uidoc.Selection.SetElementIds(alive); } catch { }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Tools28.Localization;

namespace Tools28.Commands.RoomTagCreator
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RoomTagAutoCreatorCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                View activeView = doc.ActiveView;

                // シートビューか確認
                if (activeView.ViewType != ViewType.DrawingSheet)
                {
                    TaskDialog.Show(Loc.S("Common.Error"),
                        Loc.S("RoomTag.OpenSheet"));
                    return Result.Cancelled;
                }

                // ビューポートを選択
                Reference pickedRef;
                try
                {
                    pickedRef = uidoc.Selection.PickObject(
                        ObjectType.Element,
                        new ViewportSelectionFilter(),
                        "シート上のビューポートを選択してください");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (pickedRef == null)
                    return Result.Cancelled;

                Viewport viewport = doc.GetElement(pickedRef.ElementId) as Viewport;
                if (viewport == null)
                {
                    TaskDialog.Show(Loc.S("Common.Error"), Loc.S("RoomTag.NoViewport"));
                    return Result.Cancelled;
                }

                // ビューポートからビューを取得
                View sourceView = doc.GetElement(viewport.ViewId) as View;
                if (sourceView == null)
                {
                    TaskDialog.Show(Loc.S("Common.Error"), Loc.S("RoomTag.NoView"));
                    return Result.Cancelled;
                }

                // ビューポート内の部屋を取得
                var rooms = RoomTagService.GetViewportRooms(doc, viewport);
                if (rooms.Count == 0)
                {
                    TaskDialog.Show(Loc.S("Common.Warning"), Loc.S("RoomTag.NoRooms"));
                    return Result.Cancelled;
                }

                // WPFダイアログ表示
                var dialog = new RoomTagUI(doc, rooms, sourceView.Name);
                bool? dialogResult = dialog.ShowDialog();

                if (dialogResult != true)
                    return Result.Cancelled;

                // ダイアログ結果の検証
                if (string.IsNullOrEmpty(dialog.SelectedViewFamilyTypeName))
                {
                    TaskDialog.Show(Loc.S("Common.Error"), Loc.S("RoomTag.NoViewType"));
                    return Result.Failed;
                }

                if (dialog.SelectedTagTypeId == null ||
                    dialog.SelectedTagTypeId == ElementId.InvalidElementId)
                {
                    TaskDialog.Show(Loc.S("Common.Error"), Loc.S("RoomTag.NoTagType"));
                    return Result.Failed;
                }

                if (dialog.Layout == null)
                {
                    TaskDialog.Show("エラー", "配置設定が取得できませんでした。");
                    return Result.Failed;
                }

                // トランザクショングループで処理を実行
                using (TransactionGroup tg = new TransactionGroup(doc, "部屋タグ自動配置"))
                {
                    tg.Start();

                    ViewPlan newView = null;

                    // 1. 新規ビュー作成
                    using (Transaction t1 = new Transaction(doc, "新規ビュー作成"))
                    {
                        t1.Start();
                        try
                        {
                            newView = RoomTagService.CreateNewView(
                                doc, sourceView, dialog.NewViewName,
                                dialog.SelectedViewFamilyTypeName);
                        }
                        catch (InvalidOperationException ex)
                        {
                            t1.RollBack();
                            tg.RollBack();
                            TaskDialog.Show("エラー", ex.Message);
                            return Result.Failed;
                        }
                        t1.Commit();
                    }

                    // 2. 部屋タグ生成
                    using (Transaction t2 = new Transaction(doc, "部屋タグ生成"))
                    {
                        t2.Start();
                        try
                        {
                            var createdTags = RoomTagService.CreateRoomTags(
                                doc, newView, dialog.SelectedRooms,
                                dialog.SelectedTagTypeId, dialog.Layout);

                            if (createdTags.Count == 0)
                            {
                                t2.RollBack();
                                tg.RollBack();
                                TaskDialog.Show("エラー", "部屋タグの生成に失敗しました。");
                                return Result.Failed;
                            }
                        }
                        catch (Exception ex)
                        {
                            t2.RollBack();
                            tg.RollBack();
                            TaskDialog.Show("エラー",
                                $"部屋タグの生成中にエラーが発生しました。\n{ex.Message}");
                            return Result.Failed;
                        }
                        t2.Commit();
                    }

                    tg.Assimilate();

                    // 3. 新規ビューに切り替え
                    uidoc.ActiveView = newView;

                    TaskDialog.Show("完了",
                        $"部屋タグの自動配置が完了しました。\n\n" +
                        $"ビュー名: {dialog.NewViewName}\n" +
                        $"配置タグ数: {dialog.SelectedRooms.Count}");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"部屋タグ自動配置中にエラーが発生しました。\n\n{ex.Message}" +
                    $"\n\n--- スタックトレース ---\n{ex.StackTrace}" +
                    "\n\nマニュアル: https://28tools.com/addins.html" +
                    "\n配布サイト: https://28yu.github.io/28tools-download/";
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// ビューポートのみ選択可能にするフィルタ
    /// </summary>
    public class ViewportSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Viewport;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}

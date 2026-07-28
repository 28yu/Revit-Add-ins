using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Tools28.Localization;

namespace Tools28.Commands.Room3DColor
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Room3DColorCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 部屋の収集
                List<Room> rooms = RoomSolidGenerator.CollectRooms(doc);
                if (rooms.Count == 0)
                {
                    TaskDialog.Show(Loc.S("Common.Error"), Loc.S("Room3D.NoRooms"));
                    return Result.Cancelled;
                }

                var roomById = rooms.ToDictionary(r => r.Id, r => r);

                // ダイアログ用情報の構築
                List<RoomEntry> entries = RoomSolidGenerator.BuildRoomEntries(doc, rooms);
                List<string> paramNames = RoomSolidGenerator.GetCommonParameterNames(entries);

                // ダイアログ表示
                var dialog = new Room3DColorDialog(entries, paramNames);
                dialog.SetRevitOwner(commandData);
                if (dialog.ShowDialog() != true)
                    return Result.Cancelled;

                RoomColorResult result = dialog.GetResult();

                int shapeSuccess = 0;
                int shapeFailure = 0;
                bool volumeEnabled = false;
                bool workshared = doc.IsWorkshared;
                ElementId legendViewId = null;
                View3D colorView = null;

                using (TransactionGroup tg = new TransactionGroup(doc, "部屋3D色分け"))
                {
                    tg.Start();

                    // 体積計算の有効化（別トランザクションでジオメトリ再計算を確定）
                    using (Transaction t1 = new Transaction(doc, "体積計算を有効化"))
                    {
                        t1.Start();
                        volumeEnabled = RoomSolidGenerator.EnsureVolumeComputation(doc);
                        t1.Commit();
                    }

                    using (Transaction t2 = new Transaction(doc, "部屋ソリッド生成・色分け"))
                    {
                        t2.Start();

                        if (result.DeleteExisting)
                            RoomSolidGenerator.DeleteExistingShapes(doc);

                        // 専用ワークセット（ワークシェアモデルのみ）を取得・作成
                        WorksetId roomWorksetId = workshared
                            ? RoomSolidGenerator.EnsureRoomWorkset(doc, Loc.S("Room3D.WorksetName"))
                            : null;

                        // 専用3Dビューを作成
                        colorView = RoomSolidGenerator.CreateColorView(doc, result.ViewName);

                        ElementId solidFillPatternId = RoomColorManager.GetSolidFillPatternId(doc);
                        var allShapeIds = new List<ElementId>();

                        foreach (var group in result.Groups)
                        {
                            OverrideGraphicSettings ogs =
                                RoomColorManager.CreateColorOverrides(group.Color, solidFillPatternId);

                            foreach (ElementId roomId in group.RoomIds)
                            {
                                if (!roomById.TryGetValue(roomId, out Room room))
                                {
                                    shapeFailure++;
                                    continue;
                                }

                                ElementId shapeId =
                                    RoomSolidGenerator.CreateRoomShape(doc, room, roomWorksetId);
                                if (shapeId == null)
                                {
                                    shapeFailure++;
                                    continue;
                                }

                                try { colorView.SetElementOverrides(shapeId, ogs); } catch { }
                                allShapeIds.Add(shapeId);
                                shapeSuccess++;
                            }
                        }

                        // 部屋ソリッドのみ表示に絞り込み
                        RoomSolidGenerator.IsolateRoomShapes(doc, colorView, allShapeIds);

                        // 他ビューでの非表示処理
                        if (workshared && roomWorksetId != null)
                        {
                            // 専用ビューのみワークセットを表示（既定は非表示に設定済み）
                            RoomSolidGenerator.ShowWorksetInView(colorView, roomWorksetId);
                        }
                        else if (!workshared)
                        {
                            // ワークセット未使用モデルは要素単位で他ビューを非表示に
                            RoomSolidGenerator.HideShapesInOtherViews(doc, colorView.Id, allShapeIds);
                        }

                        // 凡例作成
                        if (result.CreateLegend)
                        {
                            legendViewId = RoomColorLegendManager.CreateLegendDraftingView(
                                doc, Loc.S("Room3D.LegendViewName"), result.Groups);
                        }

                        t2.Commit();
                    }

                    tg.Assimilate();
                }

                // 作成したビューをアクティブに
                if (colorView != null)
                {
                    try { uidoc.ActiveView = colorView; } catch { }
                }

                // 完了メッセージ
                string legendInfo = legendViewId != null
                    ? Loc.S("Room3D.DoneLegendCreated")
                    : Loc.S("Room3D.DoneLegendSkipped");
                string volumeInfo = volumeEnabled
                    ? "\n" + Loc.S("Room3D.VolumeEnabledNote")
                    : "";

                string visibilityInfo = workshared
                    ? "\n" + string.Format(Loc.S("Room3D.DoneWorkset"), Loc.S("Room3D.WorksetName"))
                    : "\n" + Loc.S("Room3D.DoneHiddenOtherViews");

                string doneMessage = string.Format(Loc.S("Room3D.DoneMessage"),
                    result.ViewName, shapeSuccess, result.Groups.Count) +
                    "\n" + legendInfo + visibilityInfo + volumeInfo;

                if (shapeFailure > 0)
                    doneMessage += "\n" + string.Format(Loc.S("Room3D.DoneFailure"), shapeFailure);

                TaskDialog.Show(Loc.S("Room3D.DoneTitle"), doneMessage);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = string.Format(Loc.S("Room3D.Error"), ex.Message) +
                    "\n\nマニュアル: https://28tools.com/addins.html" +
                    "\n配布サイト: https://28yu.github.io/28tools-download/";
                return Result.Failed;
            }
        }
    }
}

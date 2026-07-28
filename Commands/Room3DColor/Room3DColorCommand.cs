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
                                    RoomSolidGenerator.CreateRoomShape(doc, room, group.Label);
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

                string doneMessage = string.Format(Loc.S("Room3D.DoneMessage"),
                    result.ViewName, shapeSuccess, result.Groups.Count) +
                    "\n" + legendInfo + volumeInfo;

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

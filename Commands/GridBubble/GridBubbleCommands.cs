using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.GridBubble
{
    [Transaction(TransactionMode.Manual)]
    public class ExecuteGridBubbleBothCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                using (Transaction trans = new Transaction(doc, "符号両端表示切替"))
                {
                    trans.Start();

                    // 選択要素を取得
                    var selectedIds = uidoc.Selection.GetElementIds();
                    var selectedGrids = new List<Grid>();
                    var selectedLevels = new List<Level>();

                    // 選択要素から通り芯とレベルを抽出
                    foreach (ElementId id in selectedIds)
                    {
                        Element elem = doc.GetElement(id);
                        if (elem is Grid grid)
                            selectedGrids.Add(grid);
                        else if (elem is Level level)
                            selectedLevels.Add(level);
                    }

                    List<Grid> targetGrids;
                    List<Level> targetLevels;

                    if (selectedGrids.Count > 0 || selectedLevels.Count > 0)
                    {
                        // 選択要素がある場合：選択した通り芯・レベルのみ処理
                        targetGrids = selectedGrids;
                        targetLevels = selectedLevels;
                    }
                    else
                    {
                        // 選択要素がない場合：アクティブビューの全通り芯・レベルを処理
                        targetGrids = new FilteredElementCollector(doc, doc.ActiveView.Id)
                            .OfClass(typeof(Grid))
                            .Cast<Grid>()
                            .ToList();

                        targetLevels = new FilteredElementCollector(doc, doc.ActiveView.Id)
                            .OfClass(typeof(Level))
                            .Cast<Level>()
                            .ToList();
                    }

                    // 現在の状態を確認（最初の要素で判定）
                    bool isCurrentlyVisible = false;
                    if (targetGrids.Count > 0)
                    {
                        isCurrentlyVisible = targetGrids[0].IsBubbleVisibleInView(DatumEnds.End0, doc.ActiveView) ||
                                           targetGrids[0].IsBubbleVisibleInView(DatumEnds.End1, doc.ActiveView);
                    }
                    else if (targetLevels.Count > 0)
                    {
                        isCurrentlyVisible = targetLevels[0].IsBubbleVisibleInView(DatumEnds.End0, doc.ActiveView) ||
                                           targetLevels[0].IsBubbleVisibleInView(DatumEnds.End1, doc.ActiveView);
                    }

                    if (isCurrentlyVisible)
                    {
                        // 現在表示中 → 非表示にする
                        foreach (Grid grid in targetGrids)
                        {
                            try
                            {
                                grid.HideBubbleInView(DatumEnds.End0, doc.ActiveView);
                                grid.HideBubbleInView(DatumEnds.End1, doc.ActiveView);
                            }
                            catch { }
                        }

                        foreach (Level level in targetLevels)
                        {
                            try
                            {
                                level.HideBubbleInView(DatumEnds.End0, doc.ActiveView);
                                level.HideBubbleInView(DatumEnds.End1, doc.ActiveView);
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // 現在非表示 → 両端表示にする
                        foreach (Grid grid in targetGrids)
                        {
                            try
                            {
                                grid.ShowBubbleInView(DatumEnds.End0, doc.ActiveView);
                                grid.ShowBubbleInView(DatumEnds.End1, doc.ActiveView);
                            }
                            catch { }
                        }

                        foreach (Level level in targetLevels)
                        {
                            try
                            {
                                level.ShowBubbleInView(DatumEnds.End0, doc.ActiveView);
                                level.ShowBubbleInView(DatumEnds.End1, doc.ActiveView);
                            }
                            catch { }
                        }
                    }

                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ExecuteGridBubbleLeftCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                using (Transaction trans = new Transaction(doc, "符号左端表示切替"))
                {
                    trans.Start();

                    // 選択要素を取得
                    var selectedIds = uidoc.Selection.GetElementIds();
                    var selectedGrids = new List<Grid>();
                    var selectedLevels = new List<Level>();

                    // 選択要素から通り芯とレベルを抽出
                    foreach (ElementId id in selectedIds)
                    {
                        Element elem = doc.GetElement(id);
                        if (elem is Grid grid)
                            selectedGrids.Add(grid);
                        else if (elem is Level level)
                            selectedLevels.Add(level);
                    }

                    List<Grid> targetGrids;
                    List<Level> targetLevels;

                    if (selectedGrids.Count > 0 || selectedLevels.Count > 0)
                    {
                        // 選択要素がある場合：選択した通り芯・レベルのみ処理
                        targetGrids = selectedGrids;
                        targetLevels = selectedLevels;
                    }
                    else
                    {
                        // 選択要素がない場合：アクティブビューの全通り芯・レベルを処理
                        targetGrids = new FilteredElementCollector(doc, doc.ActiveView.Id)
                            .OfClass(typeof(Grid))
                            .Cast<Grid>()
                            .ToList();

                        targetLevels = new FilteredElementCollector(doc, doc.ActiveView.Id)
                            .OfClass(typeof(Level))
                            .Cast<Level>()
                            .ToList();
                    }

                    // 現在の左端表示状態を確認
                    bool isLeftCurrentlyVisible = false;
                    if (targetGrids.Count > 0)
                    {
                        isLeftCurrentlyVisible = targetGrids[0].IsBubbleVisibleInView(DatumEnds.End0, doc.ActiveView);
                    }
                    else if (targetLevels.Count > 0)
                    {
                        isLeftCurrentlyVisible = targetLevels[0].IsBubbleVisibleInView(DatumEnds.End0, doc.ActiveView);
                    }

                    if (isLeftCurrentlyVisible)
                    {
                        // 現在左端表示中 → 左端を非表示にする
                        foreach (Grid grid in targetGrids)
                        {
                            try
                            {
                                grid.HideBubbleInView(DatumEnds.End0, doc.ActiveView);
                            }
                            catch { }
                        }

                        foreach (Level level in targetLevels)
                        {
                            try
                            {
                                level.HideBubbleInView(DatumEnds.End0, doc.ActiveView);
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // 現在左端非表示 → 左端のみ表示にする
                        foreach (Grid grid in targetGrids)
                        {
                            try
                            {
                                grid.ShowBubbleInView(DatumEnds.End0, doc.ActiveView);  // 左端表示
                                grid.HideBubbleInView(DatumEnds.End1, doc.ActiveView); // 右端非表示
                            }
                            catch { }
                        }

                        foreach (Level level in targetLevels)
                        {
                            try
                            {
                                level.ShowBubbleInView(DatumEnds.End0, doc.ActiveView);  // 左端表示
                                level.HideBubbleInView(DatumEnds.End1, doc.ActiveView); // 右端非表示
                            }
                            catch { }
                        }
                    }

                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ExecuteGridBubbleRightCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                using (Transaction trans = new Transaction(doc, "符号右端表示切替"))
                {
                    trans.Start();

                    // 選択要素を取得
                    var selectedIds = uidoc.Selection.GetElementIds();
                    var selectedGrids = new List<Grid>();
                    var selectedLevels = new List<Level>();

                    // 選択要素から通り芯とレベルを抽出
                    foreach (ElementId id in selectedIds)
                    {
                        Element elem = doc.GetElement(id);
                        if (elem is Grid grid)
                            selectedGrids.Add(grid);
                        else if (elem is Level level)
                            selectedLevels.Add(level);
                    }

                    List<Grid> targetGrids;
                    List<Level> targetLevels;

                    if (selectedGrids.Count > 0 || selectedLevels.Count > 0)
                    {
                        // 選択要素がある場合：選択した通り芯・レベルのみ処理
                        targetGrids = selectedGrids;
                        targetLevels = selectedLevels;
                    }
                    else
                    {
                        // 選択要素がない場合：アクティブビューの全通り芯・レベルを処理
                        targetGrids = new FilteredElementCollector(doc, doc.ActiveView.Id)
                            .OfClass(typeof(Grid))
                            .Cast<Grid>()
                            .ToList();

                        targetLevels = new FilteredElementCollector(doc, doc.ActiveView.Id)
                            .OfClass(typeof(Level))
                            .Cast<Level>()
                            .ToList();
                    }

                    // 現在の右端表示状態を確認
                    bool isRightCurrentlyVisible = false;
                    if (targetGrids.Count > 0)
                    {
                        isRightCurrentlyVisible = targetGrids[0].IsBubbleVisibleInView(DatumEnds.End1, doc.ActiveView);
                    }
                    else if (targetLevels.Count > 0)
                    {
                        isRightCurrentlyVisible = targetLevels[0].IsBubbleVisibleInView(DatumEnds.End1, doc.ActiveView);
                    }

                    if (isRightCurrentlyVisible)
                    {
                        // 現在右端表示中 → 右端を非表示にする
                        foreach (Grid grid in targetGrids)
                        {
                            try
                            {
                                grid.HideBubbleInView(DatumEnds.End1, doc.ActiveView);
                            }
                            catch { }
                        }

                        foreach (Level level in targetLevels)
                        {
                            try
                            {
                                level.HideBubbleInView(DatumEnds.End1, doc.ActiveView);
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // 現在右端非表示 → 右端のみ表示にする
                        foreach (Grid grid in targetGrids)
                        {
                            try
                            {
                                grid.HideBubbleInView(DatumEnds.End0, doc.ActiveView); // 左端非表示
                                grid.ShowBubbleInView(DatumEnds.End1, doc.ActiveView); // 右端表示
                            }
                            catch { }
                        }

                        foreach (Level level in targetLevels)
                        {
                            try
                            {
                                level.HideBubbleInView(DatumEnds.End0, doc.ActiveView); // 左端非表示
                                level.ShowBubbleInView(DatumEnds.End1, doc.ActiveView); // 右端表示
                            }
                            catch { }
                        }
                    }

                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Localization;

namespace Tools28.Commands.FillPatternIO
{
    /// <summary>
    /// 開いている rvt ファイル内の塗り潰しパターンを一覧表示し、
    /// .pat ファイルへの書き出し／読み込みを行うコマンド。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class FillPatternIOCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            if (uidoc == null || uidoc.Document == null)
            {
                TaskDialog.Show(Loc.S("FillPatternIO.Title"), Loc.S("FillPatternIO.NoDocument"));
                return Result.Cancelled;
            }

            try
            {
                var dialog = new FillPatternIODialog(uidoc);
                dialog.SetRevitOwner(commandData);
                dialog.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show(Loc.S("Common.Error"),
                    string.Format(Loc.S("FillPatternIO.ProcessError"), ex.Message));
                return Result.Failed;
            }
        }
    }
}

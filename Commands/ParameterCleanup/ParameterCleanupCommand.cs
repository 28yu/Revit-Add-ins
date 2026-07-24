using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.ParameterCleanup.Views;
using Tools28.Localization;

namespace Tools28.Commands.ParameterCleanup
{
    /// <summary>
    /// 同名パラメータの特定・値の有無表示・不要（未使用）パラメータ削除を行うコマンド。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ParameterCleanupCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp?.ActiveUIDocument;
            Document doc = uidoc?.Document;

            if (doc == null)
            {
                message = Loc.S("ParamCleanup.NoDoc");
                return Result.Cancelled;
            }

            try
            {
                var dialog = new ParameterCleanupDialog(doc);
                dialog.SetRevitOwner(commandData);
                dialog.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message + "\n\nマニュアル: https://28tools.com/addins.html";
                return Result.Failed;
            }
        }
    }
}

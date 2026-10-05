using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.ParameterValueReplace.Views;
using Tools28.Localization;

namespace Tools28.Commands.ParameterValueReplace
{
    /// <summary>
    /// 同じ名前でパラメータグループが違う文字パラメータの間で、値を要素ごとに移動するコマンド
    /// （例: モデル プロパティの「☆財産区分」→ セットの「☆財産区分」）。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ParameterValueMoveCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp?.ActiveUIDocument;
            Document doc = uidoc?.Document;

            if (doc == null)
            {
                message = Loc.S("ValueReplace.NoDoc");
                return Result.Cancelled;
            }

            try
            {
                ICollection<ElementId> selection = null;
                try { selection = uidoc.Selection.GetElementIds(); } catch { }

                var dialog = new ParameterValueMoveDialog(doc, doc.ActiveView, selection);
                dialog.SetRevitOwner(commandData);
                dialog.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[ValueMove] エラー: {ex}");
                message = ex.Message + "\n\nマニュアル: https://28tools.com/addins.html";
                return Result.Failed;
            }
        }
    }
}

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
    /// モデル内の文字パラメータの値を一括で置き換える（空欄にして削除もできる）コマンド。
    /// Excel 連携を使わずに「未使用」などの値をモデル全体からまとめて消すために使う。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ParameterValueReplaceCommand : IExternalCommand
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

                var dialog = new ParameterValueReplaceDialog(doc, doc.ActiveView, selection);
                dialog.SetRevitOwner(commandData);
                dialog.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[ValueReplace] エラー: {ex}");
                message = ex.Message + "\n\nマニュアル: https://28tools.com/addins.html";
                return Result.Failed;
            }
        }
    }
}

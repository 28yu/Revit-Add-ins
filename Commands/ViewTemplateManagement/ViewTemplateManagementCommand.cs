using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.ViewTemplateManagement.Views;
using Tools28.Localization;

namespace Tools28.Commands.ViewTemplateManagement
{
    /// <summary>
    /// プロジェクト内の全ビューテンプレートを一覧表示し、
    /// 適用先ビューの確認・検索・名前変更・不要テンプレートの削除を行うコマンド。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ViewTemplateManagementCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp?.ActiveUIDocument;
            Document doc = uidoc?.Document;

            if (doc == null)
            {
                message = Loc.S("TemplateMgmt.NoDoc");
                return Result.Cancelled;
            }

            try
            {
                var dialog = new ViewTemplateManagementDialog(doc);
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

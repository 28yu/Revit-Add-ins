using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.TestButton
{
    [Transaction(TransactionMode.Manual)]
    public class TestButtonCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            TaskDialog.Show("自動デプロイ確認 v32", "バックグラウンド自動デプロイ成功！\n\n(v32: 通知メッセージ日本語化)");
            return Result.Succeeded;
        }
    }
}

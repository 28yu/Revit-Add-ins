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
            TaskDialog.Show("自動デプロイ確認 v46", "バックグラウンド自動デプロイ成功！\n\n(v46: 日本語コミットメッセージ通知テスト)");
            return Result.Succeeded;
        }
    }
}

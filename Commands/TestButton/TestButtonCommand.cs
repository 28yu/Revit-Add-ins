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
            TaskDialog.Show("自動デプロイ確認 v7", "AutoBuild完全動作テスト成功！\n\npush → auto-merge → AutoBuild検知 → ビルド → デプロイ\nの全自動パイプラインが正常に動作しています。");
            return Result.Succeeded;
        }
    }
}

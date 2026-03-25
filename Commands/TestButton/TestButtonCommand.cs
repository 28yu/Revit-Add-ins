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
            TaskDialog.Show("自動デプロイ確認 v18", "バックグラウンド自動デプロイのテスト\n\nAutoBuild.ps1ウィンドウを閉じた状態で\nタスクスケジューラ経由の自動デプロイを確認\n\n(v18: バックグラウンド動作テスト)");
            return Result.Succeeded;
        }
    }
}

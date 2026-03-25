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
            TaskDialog.Show("自動デプロイ確認 v21", "バックグラウンド自動デプロイ成功！\n\nAutoBuild自動検知 → ビルド → デプロイ → 通知\n\n(v21: 自動デプロイ動作確認テスト)");
            return Result.Succeeded;
        }
    }
}

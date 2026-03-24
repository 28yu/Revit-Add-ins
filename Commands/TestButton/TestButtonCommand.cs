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
            TaskDialog.Show("自動デプロイ確認 v2", "自動デプロイパイプライン動作確認OK！\n\n1. Claude Code でコード変更\n2. push → GitHub auto-merge\n3. AutoBuild.ps1 が検知 → pull → ビルド → デプロイ\n\nこのメッセージが見えていれば全自動化成功です！");
            return Result.Succeeded;
        }
    }
}

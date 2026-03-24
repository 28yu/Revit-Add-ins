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
            TaskDialog.Show("自動デプロイ確認", "自動デプロイが正常に動作しています！\n\nClaude Code → push → auto-merge → AutoBuild検知 → ビルド＆デプロイ\nこの一連の流れが全て自動で完了しました。");
            return Result.Succeeded;
        }
    }
}

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
            TaskDialog.Show("自動デプロイ確認 v3", "全自動パイプライン動作確認OK！\n\n手動操作なしで、このメッセージが表示されていれば\nClaude Code → push → merge → pull → ビルド → デプロイ\nの完全自動化が成功しています。");
            return Result.Succeeded;
        }
    }
}

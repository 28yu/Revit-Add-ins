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
            TaskDialog.Show("自動デプロイ確認 v22", "バックグラウンド自動デプロイ成功！\n\n(v22: ビルドエラー詳細ログ追加)");
            return Result.Succeeded;
        }
    }
}

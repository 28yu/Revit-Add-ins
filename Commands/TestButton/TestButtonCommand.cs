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
            TaskDialog.Show("自動デプロイ確認 v5", "AutoBuild改善テスト成功！\n\ngit clean -fd 追加 + ビルドエラーログ記録\nによる自動ビルド安定化を確認。");
            return Result.Succeeded;
        }
    }
}

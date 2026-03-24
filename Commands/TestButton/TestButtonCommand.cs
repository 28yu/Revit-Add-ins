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
            TaskDialog.Show("自動デプロイ確認 v6", "AutoBuild再起動後のテスト成功！\n\ngit clean -fd + ビルドエラーログ記録が有効な状態で\n自動ビルド＆デプロイが正常動作しています。");
            return Result.Succeeded;
        }
    }
}

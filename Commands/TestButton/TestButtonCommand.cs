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
            TaskDialog.Show("自動デプロイ確認 v17", "AutoBuild 全面修正後のテスト\n\n自動ビルド & デプロイが正常に\n動作することを確認します。\n\n(v17: AutoBuild英語化+終了コード修正+ゾンビプロセス対策)");
            return Result.Succeeded;
        }
    }
}

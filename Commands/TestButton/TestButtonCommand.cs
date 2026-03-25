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
            TaskDialog.Show("自動デプロイ確認 v12", "AutoBuild 起動後の初回テスト\n\n自動ビルド & デプロイが正常に\n動作することを確認します。\n\n(v12: AutoBuild初回テスト)");
            return Result.Succeeded;
        }
    }
}

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
            TaskDialog.Show("自動デプロイ確認 v11", "Revit閉じてから再ビルドテスト\n\nClosedXML.dllのロック解除後の\n自動デプロイ確認です。\n\n(v11: DLLロック解消後のリトライ)");
            return Result.Succeeded;
        }
    }
}

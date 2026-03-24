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
            TaskDialog.Show("テストボタン", "テストボタンが正常に動作しました。\nこのダイアログが表示されていれば、ビルド＆デプロイは成功です。");
            return Result.Succeeded;
        }
    }
}

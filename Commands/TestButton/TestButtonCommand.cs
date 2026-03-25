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
            TaskDialog.Show("自動デプロイ確認 v41", "バックグラウンド自動デプロイ成功！\n\n(v41: UTF-8 BOM追加)");
            return Result.Succeeded;
        }
    }
}

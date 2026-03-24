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
            TaskDialog.Show("自動デプロイ確認 v8", "バックグラウンド自動デプロイ成功！\n\nウィンドウなしのバックグラウンド監視から\n自動ビルド＆デプロイが正常に動作しています。");
            return Result.Succeeded;
        }
    }
}

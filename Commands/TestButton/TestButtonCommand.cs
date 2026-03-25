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
            TaskDialog.Show("自動デプロイ確認 v19", "バックグラウンド自動デプロイ成功！\n\nウィンドウなしでバックグラウンド監視 → 自動ビルド → デプロイ完了\n\n(v19: バックグラウンド自動デプロイ環境構築完了)");
            return Result.Succeeded;
        }
    }
}

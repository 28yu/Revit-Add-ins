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
            TaskDialog.Show("自動デプロイ確認 v10", "通知ダイアログ変更テスト\n\nビルド完了通知がMessageBoxに変更され、\n手動で閉じるまで表示されるようになりました。\n\n(v10: ビルドリトライ)");
            return Result.Succeeded;
        }
    }
}

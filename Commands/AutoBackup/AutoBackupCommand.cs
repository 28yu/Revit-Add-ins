using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.AutoBackup.Services;
using Tools28.Commands.AutoBackup.Views;

namespace Tools28.Commands.AutoBackup
{
    /// <summary>
    /// 自動バックアップの設定ダイアログを開くコマンド。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AutoBackupCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            var dialog = new AutoBackupDialog(AutoBackupService.Instance.CurrentSettings, commandData.Application);
            dialog.SetRevitOwner(commandData);
            dialog.ShowDialog();
            return Result.Succeeded;
        }
    }
}

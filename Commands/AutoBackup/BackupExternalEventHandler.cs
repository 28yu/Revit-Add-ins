using Autodesk.Revit.UI;
using Tools28.Commands.AutoBackup.Services;

namespace Tools28.Commands.AutoBackup
{
    /// <summary>
    /// タイマー（別スレッド）から Raise され、Revit がアイドル状態（=ユーザーが操作していない瞬間）
    /// になったときにメインスレッド上で実行される。ここで実際のバックアップ処理を行うことで、
    /// ユーザーの作業を割り込まずに自動保存できる。
    /// </summary>
    public class BackupExternalEventHandler : IExternalEventHandler
    {
        public void Execute(UIApplication app)
        {
            AutoBackupService.Instance.PerformBackup(app);
        }

        public string GetName() => "Tools28 AutoBackup";
    }
}

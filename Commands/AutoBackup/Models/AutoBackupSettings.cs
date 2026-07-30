namespace Tools28.Commands.AutoBackup.Models
{
    /// <summary>
    /// 自動バックアップ機能の設定。
    /// %AppData%\Tools28\AutoBackupSettings.json に JSON で永続化する。
    /// </summary>
    public class AutoBackupSettings
    {
        /// <summary>自動バックアップを有効にする。</summary>
        public bool Enabled { get; set; } = false;

        /// <summary>バックアップ間隔（分）。</summary>
        public int IntervalMinutes { get; set; } = 15;

        /// <summary>
        /// バックアップ先をモデルと同じフォルダにする。
        /// true の場合、モデルフォルダ直下の "Tools28_Backups" サブフォルダに保存する。
        /// </summary>
        public bool UseModelFolder { get; set; } = true;

        /// <summary>UseModelFolder が false のときに使うバックアップ先フォルダ。</summary>
        public string BackupFolder { get; set; } = string.Empty;

        /// <summary>モデルごとに保存する世代数（これを超えた古いバックアップは自動削除）。</summary>
        public int MaxGenerations { get; set; } = 10;

        /// <summary>
        /// バックアップ前に未保存の変更をモデルに保存する（方式A: 最新状態を保存）。
        /// false にすると、ディスク上の最後に保存されたファイルをコピーするだけになる（方式B）。
        /// </summary>
        public bool SaveBeforeBackup { get; set; } = true;

        public AutoBackupSettings Clone()
        {
            return new AutoBackupSettings
            {
                Enabled = Enabled,
                IntervalMinutes = IntervalMinutes,
                UseModelFolder = UseModelFolder,
                BackupFolder = BackupFolder,
                MaxGenerations = MaxGenerations,
                SaveBeforeBackup = SaveBeforeBackup,
            };
        }
    }
}

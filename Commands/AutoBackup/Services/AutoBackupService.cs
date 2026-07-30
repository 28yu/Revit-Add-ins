using System;
using System.IO;
using System.Linq;
using System.Timers;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.AutoBackup.Models;
using Tools28.Localization;

namespace Tools28.Commands.AutoBackup.Services
{
    /// <summary>
    /// 自動バックアップの中枢。タイマーで一定間隔ごとに ExternalEvent を Raise し、
    /// Revit がアイドルになった瞬間にメインスレッドでバックアップを実行する。
    /// Revit API はシングルスレッドのため保存自体はメインスレッドで行うが、
    /// アイドル時に限定することでユーザーの操作を割り込まない。
    /// </summary>
    public sealed class AutoBackupService
    {
        private static readonly AutoBackupService _instance = new AutoBackupService();
        public static AutoBackupService Instance => _instance;

        private const string BackupSubFolder = "Tools28_Backups";
        private const string BackupMarker = "_backup_";

        private ExternalEvent _externalEvent;
        private Timer _timer;
        private bool _isBackingUp;
        private readonly object _sync = new object();

        private AutoBackupSettings _settings = new AutoBackupSettings();

        /// <summary>最終バックアップ日時（未実行なら null）。</summary>
        public DateTime? LastBackupTime { get; private set; }

        /// <summary>最終バックアップの結果メッセージ（UI 表示用）。</summary>
        public string LastBackupMessage { get; private set; }

        public AutoBackupSettings CurrentSettings => _settings.Clone();

        private AutoBackupService() { }

        /// <summary>
        /// OnStartup から呼ぶ。ExternalEvent を生成（メインスレッド必須）し、設定を読み込んで
        /// 有効ならタイマーを開始する。
        /// </summary>
        public void Initialize()
        {
            try
            {
                _externalEvent = ExternalEvent.Create(new BackupExternalEventHandler());
                _settings = AutoBackupSettingsService.Load();
                _timer = new Timer { AutoReset = true };
                _timer.Elapsed += OnTimerElapsed;
                ApplyTimerState();
                DiagLog.Write($"AutoBackup 初期化完了 (有効={_settings.Enabled}, 間隔={_settings.IntervalMinutes}分)");
            }
            catch (Exception ex)
            {
                DiagLog.Write($"AutoBackup 初期化失敗: {ex.Message}");
            }
        }

        /// <summary>OnShutdown から呼ぶ。</summary>
        public void Shutdown()
        {
            try
            {
                _timer?.Stop();
                _timer?.Dispose();
                _timer = null;
                _externalEvent?.Dispose();
                _externalEvent = null;
            }
            catch { }
        }

        /// <summary>設定ダイアログからの適用。永続化してタイマーを再構成する。</summary>
        public void ApplySettings(AutoBackupSettings settings)
        {
            _settings = settings.Clone();
            AutoBackupSettingsService.Save(_settings);
            ApplyTimerState();
            DiagLog.Write($"AutoBackup 設定適用 (有効={_settings.Enabled}, 間隔={_settings.IntervalMinutes}分)");
        }

        /// <summary>「今すぐバックアップ」ボタン用。有効/無効に関わらず 1 回実行する。</summary>
        public void RequestManualBackup()
        {
            try
            {
                _externalEvent?.Raise();
            }
            catch (Exception ex)
            {
                DiagLog.Write($"AutoBackup 手動要求失敗: {ex.Message}");
            }
        }

        private void ApplyTimerState()
        {
            if (_timer == null) return;
            _timer.Stop();
            if (_settings.Enabled)
            {
                int minutes = Math.Max(1, _settings.IntervalMinutes);
                _timer.Interval = minutes * 60_000.0;
                _timer.Start();
            }
        }

        private void OnTimerElapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                _externalEvent?.Raise();
            }
            catch (Exception ex)
            {
                DiagLog.Write($"AutoBackup タイマー Raise 失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// ExternalEventHandler から呼ばれる。メインスレッド上・API コンテキスト内で実行される。
        /// </summary>
        public void PerformBackup(UIApplication app)
        {
            lock (_sync)
            {
                if (_isBackingUp) return;
                _isBackingUp = true;
            }

            try
            {
                Document doc = app?.ActiveUIDocument?.Document;
                if (doc == null)
                {
                    SetStatus(false, Loc.S("AutoBackup.Status.NoDoc"), updateTime: false);
                    return;
                }

                if (doc.IsLinked)
                {
                    SetStatus(false, Loc.S("AutoBackup.Status.Linked"), updateTime: false);
                    return;
                }

                string sourcePath = doc.PathName;
                if (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath))
                {
                    SetStatus(false, Loc.S("AutoBackup.Status.NotSaved"), updateTime: false);
                    return;
                }

                // 方式A: 未保存の変更をディスクに書き出してから、その最新ファイルをコピーする。
                if (_settings.SaveBeforeBackup && doc.IsModified)
                {
                    doc.Save();
                }

                string backupFolder = ResolveBackupFolder(sourcePath);
                Directory.CreateDirectory(backupFolder);

                string baseName = Path.GetFileNameWithoutExtension(sourcePath);
                string ext = Path.GetExtension(sourcePath);
                string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string destName = $"{baseName}{BackupMarker}{stamp}{ext}";
                string destPath = Path.Combine(backupFolder, destName);

                File.Copy(sourcePath, destPath, overwrite: true);

                RotateBackups(backupFolder, baseName, ext);

                SetStatus(true, Loc.S("AutoBackup.Status.Success"), updateTime: true);
                DiagLog.Write($"AutoBackup 成功: {destPath}");
            }
            catch (Exception ex)
            {
                SetStatus(false, string.Format(Loc.S("AutoBackup.Status.Failed"), ex.Message), updateTime: false);
                DiagLog.Write($"AutoBackup 失敗: {ex}");
            }
            finally
            {
                lock (_sync) { _isBackingUp = false; }
            }
        }

        private string ResolveBackupFolder(string sourcePath)
        {
            if (_settings.UseModelFolder || string.IsNullOrWhiteSpace(_settings.BackupFolder))
            {
                string modelDir = Path.GetDirectoryName(sourcePath) ?? string.Empty;
                return Path.Combine(modelDir, BackupSubFolder);
            }
            return _settings.BackupFolder;
        }

        /// <summary>
        /// 同一モデルのバックアップが MaxGenerations を超えた分（古い順）を削除する。
        /// </summary>
        private void RotateBackups(string folder, string baseName, string ext)
        {
            try
            {
                int keep = Math.Max(1, _settings.MaxGenerations);
                string pattern = $"{baseName}{BackupMarker}*{ext}";
                var files = new DirectoryInfo(folder)
                    .GetFiles(pattern)
                    .OrderByDescending(f => f.LastWriteTime)
                    .ToList();

                foreach (var old in files.Skip(keep))
                {
                    try { old.Delete(); }
                    catch (Exception ex) { DiagLog.Write($"AutoBackup 旧世代削除失敗: {ex.Message}"); }
                }
            }
            catch (Exception ex)
            {
                DiagLog.Write($"AutoBackup ローテーション失敗: {ex.Message}");
            }
        }

        private void SetStatus(bool success, string message, bool updateTime)
        {
            LastBackupMessage = message;
            if (updateTime) LastBackupTime = DateTime.Now;
        }
    }
}

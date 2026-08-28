using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Timers;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
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

        private UIControlledApplication _uiCtrlApp;
        // 中央同期の実行中だけ true。この間に出るクリック要求ダイアログを自動処理する。
        private volatile bool _isSyncing;

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
        public void Initialize(UIControlledApplication uiCtrlApp)
        {
            try
            {
                _uiCtrlApp = uiCtrlApp;
                _externalEvent = ExternalEvent.Create(new BackupExternalEventHandler());
                _settings = AutoBackupSettingsService.Load();
                _timer = new Timer { AutoReset = true };
                _timer.Elapsed += OnTimerElapsed;

                if (_uiCtrlApp != null)
                {
                    // 同期中に出たダイアログを記録し、許可リストに載っているものだけ自動処理する。
                    _uiCtrlApp.DialogBoxShowing += OnDialogBoxShowing;
                    // モデルを開いた直後にバックアップが走らないよう、カウントをリセットする。
                    _uiCtrlApp.ControlledApplication.DocumentOpened += OnDocumentOpened;
                }

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
                if (_uiCtrlApp != null)
                {
                    try { _uiCtrlApp.DialogBoxShowing -= OnDialogBoxShowing; } catch { }
                    try { _uiCtrlApp.ControlledApplication.DocumentOpened -= OnDocumentOpened; } catch { }
                    _uiCtrlApp = null;
                }
                _timer?.Stop();
                _timer?.Dispose();
                _timer = null;
                _externalEvent?.Dispose();
                _externalEvent = null;
            }
            catch { }
        }

        /// <summary>
        /// モデルを開いた直後にバックアップ/同期が走らないよう、間隔カウントをリセットする。
        /// これにより「開いた瞬間に同期が始まる」現象を防ぎ、開いてから1間隔ぶんは作業に集中できる。
        /// </summary>
        private void OnDocumentOpened(object sender, Autodesk.Revit.DB.Events.DocumentOpenedEventArgs e)
        {
            try { ApplyTimerState(); } catch { }
        }

        /// <summary>
        /// 自動的に OK 応答してよいと確認できたダイアログの DialogId（許可リスト）。
        ///
        /// ここに載っていないダイアログは自動処理せず、Revit に通常どおり表示させて
        /// ユーザーに判断させる。実運用ログ（C:\temp\Tools28_debug.txt の
        /// 「AutoBackup 同期中ダイアログ」行）で DialogId を確認し、
        /// 安全と判断できたものだけをここに追加すること。
        /// </summary>
        private static readonly HashSet<string> AutoDismissDialogIds
            = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                // 例: "TaskDialog_Local_Changes_Not_Synchronized"
            };

        /// <summary>
        /// 中央同期の実行中に出たダイアログを記録し、許可リストにあるものだけ自動処理する。
        /// 同期中以外は何もしない。
        ///
        /// ⚠ 以前は同期中の「すべての」ダイアログを OK で自動処理していたが、
        ///    競合・権限・保存先の警告といった、本来ユーザーが判断すべき確認まで
        ///    勝手に承認してしまう状態だった。
        ///    実運用ログから「正常に完了する同期ではダイアログが出ない」ことが確認できたため、
        ///    許可リスト方式に切り替えた。普段の自動同期は従来どおり止まらずに完了し、
        ///    例外時だけ手動同期と同じようにユーザーが判断できる。
        ///
        /// ※ 同期の進捗バー自体はダイアログではないため抑制できず、その間は編集不可（Revit 仕様）。
        /// </summary>
        private void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs e)
        {
            if (!_isSyncing) return;

            string dialogId = null;
            try { dialogId = e.DialogId; }
            catch { }

            if (!string.IsNullOrEmpty(dialogId) && AutoDismissDialogIds.Contains(dialogId))
            {
                try
                {
                    e.OverrideResult((int)TaskDialogResult.Ok);
                    DiagLog.Write($"AutoBackup 同期中ダイアログを自動処理: DialogId={dialogId}");
                }
                catch (Exception ex)
                {
                    DiagLog.Write($"AutoBackup 同期中ダイアログの自動処理に失敗: DialogId={dialogId} / {ex.Message}");
                }
                return;
            }

            // 許可リストに無い＝ユーザーに判断させる。OverrideResult は呼ばない。
            string shownId = string.IsNullOrEmpty(dialogId) ? "(取得不可)" : dialogId;
            DiagLog.Write($"AutoBackup 同期中ダイアログ（自動処理せず・ユーザー判断）: DialogId={shownId}");
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

                // ワークシェアモデルで「中央モデルへ自動同期」が有効な場合は、
                // ファイルコピーではなく Synchronize with Central を実行する。
                // クラウド中央モデル（BIM 360 / ACC）はファイルコピーできないため、この経路が必須。
                if (doc.IsWorkshared && _settings.SyncWorksharedToCentral)
                {
                    SyncWithCentral(doc);
                    return;
                }

                // クラウドモデルはローカルにコピー可能な実ファイルを持たない。
                // 自動同期が無効のままではバックアップできないため、明示的に案内する。
                if (doc.IsModelInCloud)
                {
                    SetStatus(false, Loc.S("AutoBackup.Status.CloudNeedsSync"), updateTime: false);
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
                // 次回は「今回の完了時点」から1間隔ぶん空ける（連続発火を防ぐ）。
                try { ApplyTimerState(); } catch { }
            }
        }

        /// <summary>
        /// ワークシェア中央モデルへ同期する。クラウド中央モデルの場合、
        /// これにより ACC / BIM 360 側にバージョン履歴（＝バックアップ）が蓄積される。
        /// 継続作業を妨げないよう権利の返却（Relinquish）は行わない。
        /// </summary>
        private void SyncWithCentral(Document doc)
        {
            var transactOptions = new TransactWithCentralOptions();
            // 中央がロック中（他ユーザーが同期中など）の場合は待たずにスキップし、
            // 待機ダイアログで手が止まらないようにする。
            transactOptions.SetLockCallback(new SkipIfLockedCallback());

            var syncOptions = new SynchronizeWithCentralOptions();
            // 権利を返却しない（ユーザーが編集中の要素の所有権を保持したまま同期）。
            syncOptions.SetRelinquishOptions(new RelinquishOptions(false));
            // 同期コメントは中央モデルの履歴に残り、他のメンバー全員が見る文字列。
            // アドインの言語設定ではなく Revit 本体の UI 言語に合わせる。
            syncOptions.Comment = Loc.S("AutoBackup.Sync.Comment", RevitUiLanguage.Resolve(doc));
            syncOptions.Compact = false;

            _isSyncing = true;
            try
            {
                doc.SynchronizeWithCentral(transactOptions, syncOptions);
                SetStatus(true, Loc.S("AutoBackup.Status.Synced"), updateTime: true);
                DiagLog.Write($"AutoBackup 中央同期 成功: {doc.Title}");
            }
            catch (Exception ex)
            {
                // ロック未取得や一時的な状態は「失敗」ではなくスキップ扱いにして次回に委ねる。
                // （自動バックアップが赤いエラーを出し続けないようにするため）
                SetStatus(false, Loc.S("AutoBackup.Status.SyncSkipped"), updateTime: false);
                DiagLog.Write($"AutoBackup 中央同期 スキップ/失敗: {ex.Message}");
            }
            finally
            {
                _isSyncing = false;
            }
        }

        /// <summary>
        /// 中央モデルがロックされている場合、待たずに即スキップさせるコールバック。
        /// これにより他メンバーの同期中でも待機ダイアログで固まらない。
        /// </summary>
        private sealed class SkipIfLockedCallback : ICentralLockedCallback
        {
            public bool ShouldWaitForLockAvailability() => false;
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
                    // ⚠ LastWriteTime は使えない。File.Copy はコピー元の更新日時を引き継ぐため、
                    //    モデルが未変更のまま複数回バックアップすると全ファイルが同じ日時になり、
                    //    並びが不定になって新しい世代を削除してしまうことがある。
                    //    自分で付けたファイル名のタイムスタンプで並べる
                    //    （yyyyMMdd_HHmmss は辞書順＝時系列順）。
                    .OrderByDescending(f => ExtractBackupStamp(f.Name), StringComparer.Ordinal)
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

        /// <summary>
        /// "モデル名_backup_20260828_143000.rvt" から "20260828_143000" を取り出す。
        /// 取り出せない場合は空文字（＝最も古い扱い）。
        /// </summary>
        private static string ExtractBackupStamp(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return string.Empty;

            int i = fileName.LastIndexOf(BackupMarker, StringComparison.Ordinal);
            if (i < 0) return string.Empty;

            string rest = fileName.Substring(i + BackupMarker.Length);
            int dot = rest.LastIndexOf('.');
            return dot >= 0 ? rest.Substring(0, dot) : rest;
        }

        private void SetStatus(bool success, string message, bool updateTime)
        {
            LastBackupMessage = message;
            if (updateTime) LastBackupTime = DateTime.Now;
        }
    }
}

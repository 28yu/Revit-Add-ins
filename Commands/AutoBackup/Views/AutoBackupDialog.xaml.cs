using System;
using System.Windows;
using Autodesk.Revit.UI;
using Tools28.Commands.AutoBackup.Models;
using Tools28.Commands.AutoBackup.Services;
using Tools28.Localization;
using FolderBrowserDialog = System.Windows.Forms.FolderBrowserDialog;
using WinFormsDialogResult = System.Windows.Forms.DialogResult;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace Tools28.Commands.AutoBackup.Views
{
    public partial class AutoBackupDialog : Window
    {
        private readonly UIApplication _uiapp;

        public AutoBackupDialog(AutoBackupSettings settings, UIApplication uiapp)
        {
            _uiapp = uiapp;
            InitializeComponent();
            ApplyLocalization();
            Load(settings ?? new AutoBackupSettings());
            UpdateLastBackupLabel();
        }

        private void ApplyLocalization()
        {
            Title = Loc.S("AutoBackup.Title");
            txtEnabled.Text = Loc.S("AutoBackup.Enable");
            grpInterval.Header = Loc.S("AutoBackup.Interval.Header");
            txtIntervalUnit.Text = Loc.S("AutoBackup.Interval.Unit");
            grpFolder.Header = Loc.S("AutoBackup.Folder.Header");
            txtModelFolder.Text = Loc.S("AutoBackup.Folder.Model");
            txtCustomFolder.Text = Loc.S("AutoBackup.Folder.Custom");
            grpAdvanced.Header = Loc.S("AutoBackup.Advanced.Header");
            txtGenerations.Text = Loc.S("AutoBackup.Generations");
            txtSaveBeforeBackup.Text = Loc.S("AutoBackup.SaveBeforeBackup");
            txtSyncWorkshared.Text = Loc.S("AutoBackup.SyncWorkshared");
            txtSyncNote.Text = Loc.S("AutoBackup.SyncWorkshared.Note");
            BtnBackupNow.Content = Loc.S("AutoBackup.BackupNow");
            BtnOK.Content = Loc.S("Common.OK");
            BtnCancel.Content = Loc.S("Common.Cancel");
        }

        private void Load(AutoBackupSettings s)
        {
            ChkEnabled.IsChecked = s.Enabled;
            TxtInterval.Text = s.IntervalMinutes.ToString();
            RadioModelFolder.IsChecked = s.UseModelFolder;
            RadioCustomFolder.IsChecked = !s.UseModelFolder;
            TxtFolder.Text = s.BackupFolder ?? string.Empty;
            TxtGenerations.Text = s.MaxGenerations.ToString();
            ChkSaveBeforeBackup.IsChecked = s.SaveBeforeBackup;
            ChkSyncWorkshared.IsChecked = s.SyncWorksharedToCentral;
        }

        /// <summary>UI から設定を構築する。検証エラー時は null を返す。</summary>
        private AutoBackupSettings BuildSettings()
        {
            if (!int.TryParse(TxtInterval.Text.Trim(), out int interval) || interval < 1)
            {
                TaskDialog.Show(Loc.S("Common.InputError"), Loc.S("AutoBackup.Error.Interval"));
                return null;
            }

            if (!int.TryParse(TxtGenerations.Text.Trim(), out int generations) || generations < 1)
            {
                TaskDialog.Show(Loc.S("Common.InputError"), Loc.S("AutoBackup.Error.Generations"));
                return null;
            }

            bool useModelFolder = RadioModelFolder.IsChecked == true;
            string folder = TxtFolder.Text.Trim();
            if (!useModelFolder && string.IsNullOrWhiteSpace(folder))
            {
                TaskDialog.Show(Loc.S("Common.InputError"), Loc.S("AutoBackup.Error.Folder"));
                return null;
            }

            return new AutoBackupSettings
            {
                Enabled = ChkEnabled.IsChecked == true,
                IntervalMinutes = interval,
                UseModelFolder = useModelFolder,
                BackupFolder = folder,
                MaxGenerations = generations,
                SaveBeforeBackup = ChkSaveBeforeBackup.IsChecked == true,
                SyncWorksharedToCentral = ChkSyncWorkshared.IsChecked == true,
            };
        }

        private void UpdateLastBackupLabel()
        {
            var svc = AutoBackupService.Instance;
            if (svc.LastBackupTime == null && string.IsNullOrEmpty(svc.LastBackupMessage))
            {
                txtLastBackup.Text = Loc.S("AutoBackup.Last.Never");
                return;
            }

            string time = svc.LastBackupTime?.ToString("yyyy-MM-dd HH:mm:ss")
                          ?? Loc.S("AutoBackup.Last.NeverShort");
            txtLastBackup.Text = string.Format(Loc.S("AutoBackup.Last"), time, svc.LastBackupMessage);
        }

        private void OnBrowse(object sender, RoutedEventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                if (!string.IsNullOrWhiteSpace(TxtFolder.Text))
                    dlg.SelectedPath = TxtFolder.Text.Trim();

                if (dlg.ShowDialog() == WinFormsDialogResult.OK)
                {
                    TxtFolder.Text = dlg.SelectedPath;
                    RadioCustomFolder.IsChecked = true;
                }
            }
        }

        private void OnBackupNow(object sender, RoutedEventArgs e)
        {
            var settings = BuildSettings();
            if (settings == null) return;

            // 設定を先に適用してからバックアップを実行する。
            AutoBackupService.Instance.ApplySettings(settings);

            // モーダルダイアログ表示中は ExternalEvent が発火しないため、
            // ここではメインスレッド上（コマンドの API コンテキスト内）で直接実行して
            // 即座に結果を反映する。
            AutoBackupService.Instance.PerformBackup(_uiapp);
            UpdateLastBackupLabel();
        }

        private void OnOK(object sender, RoutedEventArgs e)
        {
            var settings = BuildSettings();
            if (settings == null) return;

            AutoBackupService.Instance.ApplySettings(settings);
            DialogResult = true;
            Close();
        }
    }
}

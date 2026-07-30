using System;
using System.IO;
using System.Runtime.Serialization.Json;
using Tools28.Commands.AutoBackup.Models;

namespace Tools28.Commands.AutoBackup.Services
{
    /// <summary>
    /// 自動バックアップ設定の保存/読込サービス。
    /// </summary>
    public static class AutoBackupSettingsService
    {
        private static string GetSettingsPath()
        {
            string appDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Tools28");
            Directory.CreateDirectory(appDataFolder);
            return Path.Combine(appDataFolder, "AutoBackupSettings.json");
        }

        public static AutoBackupSettings Load()
        {
            try
            {
                string path = GetSettingsPath();
                if (!File.Exists(path))
                    return new AutoBackupSettings();

                var serializer = new DataContractJsonSerializer(typeof(AutoBackupSettings));
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    var settings = serializer.ReadObject(stream) as AutoBackupSettings;
                    return settings ?? new AutoBackupSettings();
                }
            }
            catch (Exception ex)
            {
                DiagLog.Write($"AutoBackupSettings 読込失敗: {ex.Message}");
                return new AutoBackupSettings();
            }
        }

        public static void Save(AutoBackupSettings settings)
        {
            try
            {
                var serializer = new DataContractJsonSerializer(typeof(AutoBackupSettings));
                using (var stream = new FileStream(GetSettingsPath(), FileMode.Create, FileAccess.Write))
                {
                    serializer.WriteObject(stream, settings);
                }
            }
            catch (Exception ex)
            {
                DiagLog.Write($"AutoBackupSettings 保存失敗: {ex.Message}");
            }
        }
    }
}

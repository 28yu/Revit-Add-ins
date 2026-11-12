using System;
using System.IO;

namespace Tools28.Commands.PrintHiddenText
{
    /// <summary>
    /// 「印刷しない注釈」機能の ON/OFF。既定は ON。
    /// %AppData%\Tools28\PrintHidden.setting に "on" / "off" を保存する（パソコン・Windows ユーザーごと）。
    /// </summary>
    internal static class PrintHiddenSettings
    {
        private static bool? _enabled;

        private static string FilePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Tools28", "PrintHidden.setting");

        public static bool Enabled
        {
            get
            {
                if (_enabled == null)
                {
                    bool value = true;   // ファイルが無い・読めないときは ON
                    try
                    {
                        if (File.Exists(FilePath))
                            value = !string.Equals(File.ReadAllText(FilePath).Trim(), "off", StringComparison.OrdinalIgnoreCase);
                    }
                    catch { }
                    _enabled = value;
                }
                return _enabled.Value;
            }
            set
            {
                _enabled = value;
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                    File.WriteAllText(FilePath, value ? "on" : "off");
                }
                catch (Exception ex)
                {
                    DiagLog.Write($"PrintHiddenText 設定の保存に失敗: {ex.Message}");
                }
            }
        }
    }
}

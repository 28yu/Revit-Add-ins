using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace Tools28.Localization
{
    internal static class Loc
    {
        private static string _lang = "JP";
        private static readonly Dictionary<string, Dictionary<string, string>> _strings
            = new Dictionary<string, Dictionary<string, string>>();

        public static event Action LanguageChanged;

        public static string CurrentLang => _lang;

        public static void SetLanguage(string lang)
        {
            if (_lang == lang) return;
            _lang = lang;
            SaveLanguageSetting(lang);
            LanguageChanged?.Invoke();
        }

        /// <summary>アドイン UI 用。ユーザーが選んだ言語（CurrentLang）で解決する。</summary>
        public static string S(string key)
        {
            if (_strings.TryGetValue(_lang, out var dict) && dict.TryGetValue(key, out var val))
                return val;
            if (_strings.TryGetValue("JP", out var jpDict) && jpDict.TryGetValue(key, out var jpVal))
                return jpVal;
            return key;
        }

        /// <summary>
        /// 言語を明示して解決する。モデルに保存される文字列（シート名・同期コメント等）用。
        /// アドインの言語設定ではなく <see cref="RevitUiLanguage"/> で求めた
        /// Revit 本体の言語コードを渡すこと。
        /// </summary>
        public static string S(string key, string lang)
        {
            if (!string.IsNullOrEmpty(lang) &&
                _strings.TryGetValue(lang, out var dict) && dict.TryGetValue(key, out var val))
                return val;
            return S(key);   // 見つからなければ CurrentLang → JP の順でフォールバック
        }

        static Loc()
        {
            _strings["JP"] = StringsJP.All;
            _strings["US"] = StringsEN.All;
            _strings["CN"] = StringsCN.All;
            _lang = LoadLanguageSetting();
        }

        /// <summary>
        /// 言語設定の保存先。%AppData%\Tools28\ に置く。
        /// 旧版は DLL と同じフォルダ（C:\ProgramData\Autodesk\Revit\Addins\...）に保存していたが、
        /// このフォルダは管理者権限の install.bat が作るため一般ユーザーの Revit からは書き込めず、
        /// 言語切替が再起動のたびに失われていた。
        /// AutoBackup / ExcelExportImport の設定と同じ per-user の場所に統一する。
        /// </summary>
        private static string GetSettingsPath()
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Tools28");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "28tools_lang.txt");
        }

        /// <summary>旧保存先（DLL と同じフォルダ）。既存ユーザーの設定を引き継ぐためだけに読む。</summary>
        private static string GetLegacySettingsPath()
        {
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            return Path.Combine(dir ?? string.Empty, "28tools_lang.txt");
        }

        private static string LoadLanguageSetting()
        {
            // 新しい保存先 → 旧保存先 の順に読み、最初に見つかった有効な値を採用する。
            foreach (var getPath in new Func<string>[] { GetSettingsPath, GetLegacySettingsPath })
            {
                try
                {
                    string path = getPath();
                    if (string.IsNullOrEmpty(path) || !File.Exists(path)) continue;

                    string lang = File.ReadAllText(path).Trim();
                    if (lang == "JP" || lang == "US" || lang == "CN")
                        return lang;
                }
                catch { }
            }
            return "JP";
        }

        private static void SaveLanguageSetting(string lang)
        {
            try
            {
                File.WriteAllText(GetSettingsPath(), lang);
            }
            catch (Exception ex)
            {
                // 保存失敗を握り潰すと「切り替えたのに再起動で戻る」原因が追えなくなるため記録する。
                DiagLog.Write($"言語設定の保存に失敗: {ex.Message}");
            }
        }
    }
}

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

        public static string S(string key)
        {
            if (_strings.TryGetValue(_lang, out var dict) && dict.TryGetValue(key, out var val))
                return val;
            if (_strings.TryGetValue("JP", out var jpDict) && jpDict.TryGetValue(key, out var jpVal))
                return jpVal;
            return key;
        }

        static Loc()
        {
            _strings["JP"] = StringsJP.All;
            _strings["US"] = StringsEN.All;
            _strings["CN"] = StringsCN.All;
            _lang = LoadLanguageSetting();
        }

        private static string GetSettingsPath()
        {
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            return Path.Combine(dir, "28tools_lang.txt");
        }

        private static string LoadLanguageSetting()
        {
            try
            {
                string path = GetSettingsPath();
                if (File.Exists(path))
                {
                    string lang = File.ReadAllText(path).Trim();
                    if (lang == "JP" || lang == "US" || lang == "CN")
                        return lang;
                }
            }
            catch { }
            return "JP";
        }

        private static void SaveLanguageSetting(string lang)
        {
            try
            {
                File.WriteAllText(GetSettingsPath(), lang);
            }
            catch { }
        }
    }
}

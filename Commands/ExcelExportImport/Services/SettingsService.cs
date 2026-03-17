using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using Tools28.Commands.ExcelExportImport.Models;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// エクスポート設定の保存/読込サービス
    /// </summary>
    public static class SettingsService
    {
        /// <summary>
        /// 設定をJSONファイルに保存
        /// </summary>
        public static void SaveSettings(string filePath, ExportSettings settings)
        {
            var serializer = new DataContractJsonSerializer(typeof(ExportSettings));
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                serializer.WriteObject(stream, settings);
            }
        }

        /// <summary>
        /// JSONファイルから設定を読込
        /// </summary>
        public static ExportSettings LoadSettings(string filePath)
        {
            var serializer = new DataContractJsonSerializer(typeof(ExportSettings));
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                return (ExportSettings)serializer.ReadObject(stream);
            }
        }

        /// <summary>
        /// デフォルト設定ファイルのパスを取得
        /// </summary>
        public static string GetDefaultSettingsPath()
        {
            string appDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Tools28");
            Directory.CreateDirectory(appDataFolder);
            return Path.Combine(appDataFolder, "ExcelExportSettings.json");
        }

        /// <summary>
        /// 現在のダイアログ状態からExportSettingsを生成
        /// </summary>
        public static ExportSettings CreateFromSelection(
            List<CategoryInfo> selectedCategories,
            List<ParameterInfo> outputParameters)
        {
            var settings = new ExportSettings();

            settings.SelectedCategories = selectedCategories
                .Select(c => c.Name)
                .ToList();

            settings.OutputParameters = outputParameters
                .Select(p => new ExportParameterEntry
                {
                    RawName = p.RawName,
                    IsTypeParameter = p.IsTypeParameter,
                    CategoryName = p.CategoryName
                })
                .ToList();

            return settings;
        }
    }
}

using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace Tools28.Commands.LanguageSwitch
{
    internal static class LanguageHelper
    {
        internal static void SetLanguage(string langCode, string flagIconFileName)
        {
            var pulldown = Application.LanguagePulldown;
            if (pulldown == null) return;

            pulldown.ItemText = langCode;

            try
            {
                var image = LoadImage(flagIconFileName);
                if (image != null)
                    pulldown.Image = image;
            }
            catch { }
        }

        private static BitmapImage LoadImage(string fileName)
        {
            try
            {
                string packUri = $"pack://application:,,,/Tools28;component/Resources/Icons/{fileName}";
                var image = new BitmapImage(new Uri(packUri));
                image.Freeze();
                return image;
            }
            catch { }

            try
            {
                string assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string filePath = Path.Combine(assemblyDir, "Icons", fileName);
                if (File.Exists(filePath))
                {
                    var image = new BitmapImage();
                    image.BeginInit();
                    image.UriSource = new Uri(filePath, UriKind.Absolute);
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.EndInit();
                    image.Freeze();
                    return image;
                }
            }
            catch { }

            return null;
        }
    }
}

using System;
using System.Windows;
using System.Windows.Interop;
using Autodesk.Revit.UI;

namespace Tools28
{
    /// <summary>
    /// WPF ダイアログを Revit のメインウィンドウを Owner として表示するヘルパー。
    /// Revit 2025/2026 (.NET 8) では Owner 未設定の ShowDialog がメインウィンドウの
    /// 背面に隠れ、Revit がフリーズしたように見える問題があるため必須。
    /// </summary>
    internal static class RevitDialogHelper
    {
        public static void SetRevitOwner(this Window dialog, ExternalCommandData commandData)
        {
            if (dialog == null || commandData == null) return;
            try
            {
                IntPtr hwnd = commandData.Application.MainWindowHandle;
                if (hwnd != IntPtr.Zero)
                    new WindowInteropHelper(dialog).Owner = hwnd;
            }
            catch { }
        }

        public static void SetRevitOwner(this Window dialog, UIApplication uiApp)
        {
            if (dialog == null || uiApp == null) return;
            try
            {
                IntPtr hwnd = uiApp.MainWindowHandle;
                if (hwnd != IntPtr.Zero)
                    new WindowInteropHelper(dialog).Owner = hwnd;
            }
            catch { }
        }
    }
}

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Interop;
using Autodesk.Revit.UI;

namespace Tools28
{
    /// <summary>
    /// WPF ダイアログを Revit メインウィンドウ配下で表示するヘルパー。
    /// Revit 2025/2026 (.NET 8) では Owner 未設定の ShowDialog がメインウィンドウの
    /// 背面に隠れる問題があるため、Owner 設定は必須。
    ///
    /// 重要: Owner は SourceInitialized イベントで設定する。
    /// 表示前に WindowInteropHelper.Owner を設定すると .NET 8 WPF で
    /// HwndSource 早期生成によるデッドロックを起こすことがある。
    /// </summary>
    internal static class RevitDialogHelper
    {
        public static void SetRevitOwner(this Window dialog, ExternalCommandData commandData)
        {
            if (dialog == null) return;

            dialog.SourceInitialized += (sender, args) =>
            {
                try
                {
                    IntPtr hwnd = GetRevitMainHandle(commandData);
                    if (hwnd != IntPtr.Zero)
                        new WindowInteropHelper(dialog).Owner = hwnd;
                }
                catch (Exception ex)
                {
                    DiagLog.Write($"SetRevitOwner 例外: {ex.Message}");
                }
            };
        }

        public static IntPtr GetRevitMainHandle(ExternalCommandData commandData)
        {
            try
            {
                if (commandData != null && commandData.Application != null)
                {
                    var h = commandData.Application.MainWindowHandle;
                    if (h != IntPtr.Zero) return h;
                }
            }
            catch { }

            try
            {
                return Process.GetCurrentProcess().MainWindowHandle;
            }
            catch { }

            return IntPtr.Zero;
        }
    }
}

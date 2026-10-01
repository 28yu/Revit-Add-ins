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

        /// <summary>
        /// <see cref="SetRevitOwner"/> に加えて、ダイアログを <b>Revit のウィンドウの中央</b>に表示し、前面に出す。
        /// WPF の CenterScreen は Revit と別のモニターに出ることがあり、2画面環境では
        /// 他のソフトのウィンドウの後ろに隠れて見つけにくくなるため、Revit の位置を基準にする。
        /// </summary>
        public static void SetRevitOwnerCentered(this Window dialog, ExternalCommandData commandData)
        {
            if (dialog == null) return;

            dialog.SetRevitOwner(commandData);
            dialog.WindowStartupLocation = WindowStartupLocation.Manual;

            // 表示前（SourceInitialized）に一度合わせ、サイズが確定した Loaded で微調整する
            dialog.SourceInitialized += (sender, args) => CenterOverRevit(dialog, commandData, measure: true);
            dialog.Loaded += (sender, args) => CenterOverRevit(dialog, commandData, measure: false);

            // 他のソフトの後ろに隠れないよう、表示されたら一度だけ前面に出す
            EventHandler onRendered = null;
            onRendered = (sender, args) =>
            {
                dialog.ContentRendered -= onRendered;
                try
                {
                    dialog.Activate();
                    dialog.Topmost = true;
                    dialog.Topmost = false;
                }
                catch (Exception ex)
                {
                    DiagLog.Write($"SetRevitOwnerCentered 前面化 例外: {ex.Message}");
                }
            };
            dialog.ContentRendered += onRendered;
        }

        private static void CenterOverRevit(Window dialog, ExternalCommandData commandData, bool measure)
        {
            try
            {
                IntPtr revit = GetRevitMainHandle(commandData);
                if (revit == IntPtr.Zero || IsIconic(revit) || !GetWindowRect(revit, out RECT r))
                    return;

                var source = PresentationSource.FromVisual(dialog);
                if (source?.CompositionTarget == null)
                    return;

                // Revit の位置（物理ピクセル）を WPF の単位に変換
                var toDip = source.CompositionTarget.TransformFromDevice;
                System.Windows.Point topLeft = toDip.Transform(new System.Windows.Point(r.Left, r.Top));
                System.Windows.Point bottomRight = toDip.Transform(new System.Windows.Point(r.Right, r.Bottom));

                double width = dialog.ActualWidth;
                double height = dialog.ActualHeight;
                if (measure || width <= 0 || height <= 0)
                {
                    // 表示前はサイズが未確定なので、指定サイズまたは内容から見積もる
                    double w = double.IsNaN(dialog.Width) ? double.PositiveInfinity : dialog.Width;
                    double h = double.IsNaN(dialog.Height) ? double.PositiveInfinity : dialog.Height;
                    dialog.Measure(new System.Windows.Size(w, h));
                    width = double.IsNaN(dialog.Width) ? dialog.DesiredSize.Width : dialog.Width;
                    height = double.IsNaN(dialog.Height) ? dialog.DesiredSize.Height : dialog.Height;
                }
                if (width <= 0 || height <= 0)
                    return;

                double left = topLeft.X + (bottomRight.X - topLeft.X - width) / 2;
                double top = topLeft.Y + (bottomRight.Y - topLeft.Y - height) / 2;

                // Revit より大きいダイアログでも上端・左端が Revit の外へはみ出さないようにする
                dialog.Left = Math.Max(left, topLeft.X);
                dialog.Top = Math.Max(top, topLeft.Y);
            }
            catch (Exception ex)
            {
                DiagLog.Write($"CenterOverRevit 例外: {ex.Message}");
            }
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool IsIconic(IntPtr hWnd);

        /// <summary>
        /// 削除トランザクションや Revit の TaskDialog 表示後に、WPF ダイアログが
        /// Revit 本体ウィンドウの背面へ隠れるのを防ぐため、前面へ復帰させる。
        /// Revit 側のウィンドウアクティブ化が処理された後に実行されるよう、
        /// Background 優先度で遅延実行する。
        /// </summary>
        public static void BringToFrontDeferred(this Window dialog)
        {
            if (dialog == null) return;
            dialog.Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    if (!dialog.IsVisible) return;
                    if (dialog.WindowState == WindowState.Minimized)
                        dialog.WindowState = WindowState.Normal;
                    dialog.Activate();
                    // 一瞬だけ最前面に上げて Revit 本体より前へ出し、すぐ通常へ戻す
                    dialog.Topmost = true;
                    dialog.Topmost = false;
                }
                catch (Exception ex)
                {
                    DiagLog.Write($"BringToFrontDeferred 例外: {ex.Message}");
                }
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>
        /// 一覧の読み込み・スキャンなど処理中の間、Revit 本体ウィンドウの操作を
        /// 一時的に無効化する。using ブロックを抜けると自動的に元の状態へ戻る。
        ///
        /// 安全上の要点: 開始時の有効/無効状態を記録し、<b>元が有効だった場合のみ</b>復帰させる。
        /// モーダル表示 (ShowDialog) によって WPF が既に Revit を無効化している場合は
        /// 何も変更しないため、モーダル中に誤って Revit を操作可能にしてしまうことがない。
        /// </summary>
        public static IDisposable BlockRevitInput(this Window dialog)
            => new RevitInputBlock(dialog);

        private sealed class RevitInputBlock : IDisposable
        {
            private readonly IntPtr _hwnd;
            private readonly bool _wasEnabled;
            private readonly Window _dialog;
            private readonly System.Windows.Input.Cursor _prevCursor;
            private bool _disposed;

            public RevitInputBlock(Window dialog)
            {
                _dialog = dialog;
                try
                {
                    if (dialog != null)
                    {
                        // ダイアログ側は待機カーソルにして処理中であることを示す
                        _prevCursor = dialog.Cursor;
                        dialog.Cursor = System.Windows.Input.Cursors.Wait;

                        _hwnd = new WindowInteropHelper(dialog).Owner;
                    }

                    if (_hwnd != IntPtr.Zero)
                    {
                        _wasEnabled = IsWindowEnabled(_hwnd);
                        if (_wasEnabled) EnableWindow(_hwnd, false);
                    }
                }
                catch (Exception ex)
                {
                    DiagLog.Write($"BlockRevitInput 開始時例外: {ex.Message}");
                }
            }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;
                try
                {
                    // 元が有効だったときだけ戻す（モーダルによる無効化は維持する）
                    if (_hwnd != IntPtr.Zero && _wasEnabled)
                        EnableWindow(_hwnd, true);
                }
                catch (Exception ex)
                {
                    DiagLog.Write($"BlockRevitInput 復帰時例外: {ex.Message}");
                }

                try
                {
                    if (_dialog != null) _dialog.Cursor = _prevCursor;
                }
                catch { }
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool IsWindowEnabled(IntPtr hWnd);

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

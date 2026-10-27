using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace Tools28.Commands.ViewCopy
{
    /// <summary>
    /// 「コピーした状態」をカーソルの右下に小さな印で示す（Revit の MA のブラシ表示の代わり）。
    ///
    /// Revit API ではカーソルそのものを変えられないため、クリックがすり抜ける小さな透明の窓を出し、
    /// タイマーでカーソルの位置を追いかける。印はコピーの種類（視点・切断ボックス）ごとに横に並べる。
    ///   - Revit が前面にあり、カーソルが Revit のウィンドウの中にあるときだけ見せる
    ///   - Revit が前面のときに Esc が押されたら、すべてのコピーを取りやめて印を消す（MA と同じ）
    /// </summary>
    internal static class CopyBadge
    {
        private class Slot
        {
            public string IconFile;
            public Action OnCancel;
        }

        // 表示順を保つため List で持つ
        private static readonly List<KeyValuePair<string, Slot>> _slots = new List<KeyValuePair<string, Slot>>();
        private static BadgeWindow _window;
        private static DispatcherTimer _timer;
        private static IntPtr _revitWindow;
        private static bool _escWasDown;
        private static readonly uint _revitProcessId = (uint)Process.GetCurrentProcess().Id;

        /// <summary>印を出す（同じキーなら差し替え）。onCancel は Esc で取りやめたときに呼ばれる。</summary>
        public static void Show(string key, string iconFile, Action onCancel, IntPtr revitWindow)
        {
            try
            {
                if (revitWindow != IntPtr.Zero) _revitWindow = revitWindow;
                _slots.RemoveAll(s => s.Key == key);
                _slots.Add(new KeyValuePair<string, Slot>(key, new Slot { IconFile = iconFile, OnCancel = onCancel }));
                Rebuild();
            }
            catch (Exception ex)
            {
                // 印は補助的な表示なので、出せなくても本来の処理は続ける
                DiagLog.Write($"[CopyBadge] 表示に失敗: {ex.Message}");
            }
        }

        /// <summary>印を消す。</summary>
        public static void Hide(string key)
        {
            try
            {
                _slots.RemoveAll(s => s.Key == key);
                Rebuild();
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[CopyBadge] 非表示に失敗: {ex.Message}");
            }
        }

        private static void Rebuild()
        {
            if (_slots.Count == 0)
            {
                _timer?.Stop();
                _window?.Hide();
                return;
            }

            if (_window == null) _window = new BadgeWindow();
            _window.SetIcons(_slots.Select(s => s.Value.IconFile));

            if (_timer == null)
            {
                _timer = new DispatcherTimer(DispatcherPriority.Input) { Interval = TimeSpan.FromMilliseconds(15) };
                _timer.Tick += (s, e) => OnTick();
            }
            _escWasDown = IsEscDown();   // ボタンを押す前から押されていた Esc では取りやめない
            _timer.Start();
            OnTick();
        }

        private static void OnTick()
        {
            try
            {
                bool revitActive = IsRevitForeground();

                // Esc（押した瞬間だけ）: すべてのコピーを取りやめる
                bool escDown = IsEscDown();
                if (revitActive && escDown && !_escWasDown)
                {
                    var cancels = _slots.Select(s => s.Value.OnCancel).ToList();
                    _slots.Clear();
                    foreach (var c in cancels)
                    {
                        try { c?.Invoke(); } catch { }
                    }
                    Rebuild();
                    _escWasDown = true;
                    return;
                }
                _escWasDown = escDown;

                if (!revitActive || !GetCursorPos(out POINT pt) || !IsInsideRevit(pt))
                {
                    if (_window.IsVisible) _window.Hide();
                    return;
                }

                if (!_window.IsVisible) _window.Show();
                // 実際の画面の画素で置く（拡大率の違うモニターでもずれない）
                SetWindowPos(_window.Handle, HWND_TOPMOST, pt.X + 16, pt.Y + 20, 0, 0,
                    SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[CopyBadge] 更新に失敗: {ex.Message}");
                _timer?.Stop();
            }
        }

        private static bool IsEscDown() => (GetAsyncKeyState(VK_ESCAPE) & 0x8000) != 0;

        private static bool IsRevitForeground()
        {
            IntPtr fg = GetForegroundWindow();
            if (fg == IntPtr.Zero) return false;
            GetWindowThreadProcessId(fg, out uint pid);
            return pid == _revitProcessId;
        }

        private static bool IsInsideRevit(POINT pt)
        {
            if (_revitWindow == IntPtr.Zero) return true;
            if (!GetWindowRect(_revitWindow, out RECT r)) return true;
            return pt.X >= r.Left && pt.X < r.Right && pt.Y >= r.Top && pt.Y < r.Bottom;
        }

        /// <summary>印を描く小さな窓。クリックはすり抜け、フォーカスも奪わない。</summary>
        private class BadgeWindow : Window
        {
            private const double IconSize = 20;
            private readonly StackPanel _panel = new StackPanel { Orientation = Orientation.Horizontal };

            public IntPtr Handle { get; private set; }

            public BadgeWindow()
            {
                WindowStyle = WindowStyle.None;
                AllowsTransparency = true;
                Background = Brushes.Transparent;
                Topmost = true;
                ShowInTaskbar = false;
                ShowActivated = false;
                Focusable = false;
                IsHitTestVisible = false;
                ResizeMode = ResizeMode.NoResize;
                SizeToContent = SizeToContent.WidthAndHeight;
                Content = _panel;

                SourceInitialized += (s, e) =>
                {
                    Handle = new WindowInteropHelper(this).Handle;
                    int ex = GetWindowLong(Handle, GWL_EXSTYLE);
                    SetWindowLong(Handle, GWL_EXSTYLE,
                        ex | WS_EX_TRANSPARENT | WS_EX_LAYERED | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW);
                };
                // ハンドルを先に作っておく（最初の位置合わせで使う）
                Handle = new WindowInteropHelper(this).EnsureHandle();
            }

            public void SetIcons(IEnumerable<string> iconFiles)
            {
                _panel.Children.Clear();
                foreach (string file in iconFiles)
                {
                    var image = new Image
                    {
                        Width = IconSize,
                        Height = IconSize,
                        Source = LoadIcon(file),
                    };
                    RenderOptions.SetBitmapScalingMode(image, BitmapScalingMode.HighQuality);
                    _panel.Children.Add(new Border
                    {
                        Child = image,
                        Padding = new Thickness(2),
                        Margin = new Thickness(0, 0, 2, 0),
                        CornerRadius = new CornerRadius(4),
                        Background = new SolidColorBrush(Color.FromArgb(235, 255, 255, 255)),
                        BorderBrush = new SolidColorBrush(Color.FromRgb(90, 90, 90)),
                        BorderThickness = new Thickness(1),
                    });
                }
            }

            private static ImageSource LoadIcon(string file)
            {
                try
                {
                    var image = new BitmapImage(new Uri(
                        $"pack://application:,,,/Tools28;component/Resources/Icons/{file}"));
                    image.Freeze();
                    return image;
                }
                catch
                {
                    return null;
                }
            }
        }

        // ---- Win32 ----
        private const int VK_ESCAPE = 0x1B;
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TRANSPARENT = 0x20;
        private const int WS_EX_LAYERED = 0x80000;
        private const int WS_EX_NOACTIVATE = 0x8000000;
        private const int WS_EX_TOOLWINDOW = 0x80;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_SHOWWINDOW = 0x0040;

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int X; public int Y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

        [DllImport("user32.dll")] private static extern bool GetCursorPos(out POINT pt);
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int vKey);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr hWnd, int index);
        [DllImport("user32.dll")] private static extern int SetWindowLong(IntPtr hWnd, int index, int value);
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
    }
}

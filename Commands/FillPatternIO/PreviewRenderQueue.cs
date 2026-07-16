using System;
using System.Threading;
using System.Windows.Media;
using System.Windows.Threading;

namespace Tools28.Commands.FillPatternIO
{
    /// <summary>
    /// プレビュー画像をバックグラウンドの STA スレッド（専用 Dispatcher）で描画するキュー。
    /// UI スレッドをブロックしないため、大量パターンでもスクロールが軽い。
    /// 描画結果（凍結済み ImageSource）はスレッド間で受け渡し可能。
    /// </summary>
    internal static class PreviewRenderQueue
    {
        private static readonly object _sync = new object();
        private static Dispatcher _dispatcher;

        private static bool EnsureWorker()
        {
            if (_dispatcher != null) return true;
            lock (_sync)
            {
                if (_dispatcher != null) return true;
                try
                {
                    var ready = new ManualResetEventSlim(false);
                    var thread = new Thread(() =>
                    {
                        _dispatcher = Dispatcher.CurrentDispatcher;
                        ready.Set();
                        Dispatcher.Run();
                    })
                    {
                        IsBackground = true,
                        Priority = ThreadPriority.BelowNormal,
                        Name = "Tools28.PreviewRenderer"
                    };
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    ready.Wait(3000);
                    return _dispatcher != null;
                }
                catch
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// バックグラウンドで描画し、完了時に onRendered をバックグラウンドスレッドで呼ぶ。
        /// 呼び出し側で UI スレッドへのマーシャリングを行うこと。
        /// ワーカースレッドを起動できない場合は同期描画にフォールバックする。
        /// </summary>
        public static void Enqueue(PatternData data, int w, int h, Action<ImageSource> onRendered)
        {
            if (EnsureWorker())
            {
                _dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
                {
                    ImageSource img = null;
                    try { img = PatternPreview.Render(data, w, h); } catch { }
                    onRendered(img);
                }));
            }
            else
            {
                ImageSource img = null;
                try { img = PatternPreview.Render(data, w, h); } catch { }
                onRendered(img);
            }
        }
    }
}

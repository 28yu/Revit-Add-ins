using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;
using Tools28.Localization;

namespace Tools28.Commands.DuplicateInstance.Views
{
    /// <summary>
    /// 削除の進み具合（％）を表示し、中止を受け付けるウィンドウ。
    /// Revit API は Revit と同じスレッドでしか呼べないため処理を別スレッドにできない。
    /// 代わりに処理の合間に <see cref="Pump"/> で画面のメッセージを処理させ、
    /// 表示の更新と「中止」ボタンの押下を受け付ける（ExcelExport の進捗ウィンドウと同じ方式）。
    /// </summary>
    public partial class DeleteProgressWindow : Window, Services.IDeleteProgress
    {
        private const long PumpIntervalMs = 150;
        private readonly Stopwatch _elapsed = Stopwatch.StartNew();
        private long _lastPumpMs = -PumpIntervalMs;
        private bool _finished;

        public bool IsCancelRequested { get; private set; }

        public DeleteProgressWindow()
        {
            InitializeComponent();
            Title = Loc.S("DupInst.Progress.Title");
            btnCancel.Content = Loc.S("DupInst.Progress.Cancel");
            Closing += (s, e) =>
            {
                if (_finished) return;
                e.Cancel = true;   // × は中止扱い（処理中にウィンドウだけ消えるのを防ぐ）
                RequestCancel();
            };
        }

        /// <summary>進み具合を更新する（force=false なら一定間隔ごとにだけ画面を描き直す）。</summary>
        public void Report(int done, int total, string status, bool force = false)
        {
            long now = _elapsed.ElapsedMilliseconds;
            if (!force && now - _lastPumpMs < PumpIntervalMs) return;
            _lastPumpMs = now;

            if (status != null) StatusText.Text = status;
            if (total > 0)
            {
                Bar.IsIndeterminate = false;
                double pct = Math.Min(100.0, done * 100.0 / total);
                Bar.Value = pct;
                PercentText.Text = $"{(int)pct}%";
                CountText.Text = string.Format(Loc.S("DupInst.Progress.Count"), done, total);
            }
            else
            {
                Bar.IsIndeterminate = true;
                PercentText.Text = "";
                CountText.Text = "";
            }
            ElapsedText.Text = string.Format(Loc.S("DupInst.Progress.Elapsed"), _elapsed.Elapsed.ToString(@"hh\:mm\:ss"));
            Pump();
        }

        /// <summary>最後の確定中（％が出せない処理）の表示。中止はできない。</summary>
        public void BeginCommit()
        {
            btnCancel.IsEnabled = false;
            Bar.IsIndeterminate = true;
            PercentText.Text = "";
            StatusText.Text = Loc.S("DupInst.Progress.Committing");
            _lastPumpMs = -PumpIntervalMs;
            Pump();
        }

        public void Finish()
        {
            _finished = true;
            try { Close(); } catch { }
        }

        private void Pump()
        {
            try { Dispatcher.Invoke(DispatcherPriority.Background, new Action(() => { })); }
            catch { }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e) => RequestCancel();

        private void RequestCancel()
        {
            if (IsCancelRequested || !btnCancel.IsEnabled) return;
            IsCancelRequested = true;
            btnCancel.IsEnabled = false;
            StatusText.Text = Loc.S("DupInst.Progress.Cancelling");
        }
    }
}

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;
using Tools28.Commands.ExcelExportImport.Services;
using Tools28.Localization;

namespace Tools28.Commands.ExcelExportImport.Views
{
    /// <summary>
    /// エクスポート中の進み具合を表示し、キャンセルを受け付けるウィンドウ。
    ///
    /// Revit API は Revit 本体と同じスレッドでしか呼べないため、処理を別スレッドに逃がせない。
    /// 代わりに処理の合間（一定時間ごと）に画面のメッセージを処理させ（<see cref="Pump"/>）、
    /// 表示の更新とキャンセルボタンの押下を受け付ける。これにより「応答なし」にならず、
    /// 極端に時間のかかるパラメータがあっても途中で止められる。
    /// </summary>
    public partial class ExportProgressWindow : Window, IExportProgress
    {
        /// <summary>画面を更新する間隔（ミリ秒）。短すぎると更新自体が処理を遅くする。</summary>
        private const long PumpIntervalMs = 200;

        /// <summary>この平均時間（ミリ秒/要素）以上のパラメータだけ「時間がかかっている」と表示する</summary>
        private const double SlowThresholdMs = 5.0;

        private readonly Stopwatch _elapsed = Stopwatch.StartNew();
        private long _lastPumpMs = -PumpIntervalMs;

        private int _elementTotal;
        private int _elementDone;
        private string _slowestLabel;
        private double _slowestAverageMs;

        private bool _finished;

        public bool IsCancelRequested { get; private set; }

        public ExportProgressWindow()
        {
            InitializeComponent();
            Title = Loc.S("Export.Title");
            btnCancel.Content = Loc.S("Common.Cancel");

            // 右上の × はキャンセル扱いにする（処理中にウィンドウだけ消えるのを防ぐ）
            Closing += (s, e) =>
            {
                if (_finished) return;
                e.Cancel = true;
                RequestCancel();
            };
        }

        /// <summary>処理が終わったらウィンドウを閉じる（× によるキャンセル扱いを解除して閉じる）。</summary>
        public void Finish()
        {
            _finished = true;
            try { Close(); } catch { }
        }

        public void BeginCategory(int index, int total, string categoryName, int elementCount)
        {
            _elementTotal = elementCount;
            _elementDone = 0;
            CategoryText.Text = string.Format(Loc.S("Export.Progress.Category"), index, total, categoryName);
            Pump(force: true);
        }

        public void ElementDone()
        {
            _elementDone++;
            Pump(force: false);
        }

        public void Tick()
        {
            Pump(force: false);
        }

        public void SetSlowest(string parameterLabel, double averageMs)
        {
            // 文字列の組み立ては表示更新時（Pump）にまとめて行う（毎要素の負荷を避ける）
            _slowestLabel = parameterLabel;
            _slowestAverageMs = averageMs;
        }

        public void BeginSaving()
        {
            CategoryText.Text = Loc.S("Export.Progress.Saving");
            _elementTotal = 0;
            btnCancel.IsEnabled = false; // 保存中は止められない
            Pump(force: true);
        }

        /// <summary>
        /// 前回の更新から一定時間たっていれば表示を更新し、溜まった画面メッセージ
        /// （描画・キャンセルボタンのクリック）を処理する。
        /// </summary>
        private void Pump(bool force)
        {
            long now = _elapsed.ElapsedMilliseconds;
            if (!force && now - _lastPumpMs < PumpIntervalMs)
                return;
            _lastPumpMs = now;

            if (_elementTotal > 0)
            {
                ElementProgress.IsIndeterminate = false;
                ElementProgress.Maximum = _elementTotal;
                ElementProgress.Value = Math.Min(_elementDone, _elementTotal);
                ElementText.Text = string.Format(Loc.S("Export.Progress.Elements"), _elementDone, _elementTotal);
            }
            else
            {
                ElementProgress.IsIndeterminate = true;
                ElementText.Text = " ";
            }

            ElapsedText.Text = string.Format(Loc.S("Export.Progress.Elapsed"),
                _elapsed.Elapsed.ToString(@"hh\:mm\:ss"));

            if (!string.IsNullOrEmpty(_slowestLabel) && _slowestAverageMs >= SlowThresholdMs)
            {
                SlowestText.Text = string.Format(Loc.S("Export.Progress.Slowest"),
                    _slowestLabel, _slowestAverageMs.ToString("0.0"));
                SlowestText.Visibility = Visibility.Visible;
            }

            // Background 優先度の空処理を同期実行 → それより優先度の高い入力・描画が先に処理される
            try
            {
                Dispatcher.Invoke(DispatcherPriority.Background, new Action(() => { }));
            }
            catch
            {
                // 表示更新の失敗で処理全体を止めない
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            RequestCancel();
        }

        private void RequestCancel()
        {
            if (IsCancelRequested || !btnCancel.IsEnabled) return;
            IsCancelRequested = true;
            btnCancel.IsEnabled = false;
            CategoryText.Text = Loc.S("Export.Progress.Cancelling");
        }
    }
}

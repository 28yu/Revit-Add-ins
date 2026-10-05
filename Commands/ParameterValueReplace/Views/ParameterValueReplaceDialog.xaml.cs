using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.DB;
using Tools28.Commands.ParameterValueReplace.Services;
using Tools28.Localization;

namespace Tools28.Commands.ParameterValueReplace.Views
{
    /// <summary>
    /// パラメータ値の一括置換ダイアログ。
    /// 「検索」で一致するパラメータをカテゴリ × パラメータごとに一覧し、
    /// チェックしたものだけを「置換を実行」で書き換える。
    /// 走査・書き込みとも一定時間ごとに画面へ制御を返すため、大容量モデルでも「応答なし」にならず中止できる。
    /// </summary>
    public partial class ParameterValueReplaceDialog : Window
    {
        private readonly Document _doc;
        private readonly View _activeView;
        private readonly ICollection<ElementId> _selectionIds;

        private List<ReplaceGroup> _groups = new List<ReplaceGroup>();

        // 検索したときの条件（置換はこの条件で行う。検索後に入力欄を変えても食い違わないように保持）
        private string _searchedFind;
        private string _searchedReplace;
        private MatchMode _searchedMode;

        private CancellationTokenSource _cts;
        private bool _busy;
        private bool _closed;

        public ParameterValueReplaceDialog(Document doc, View activeView, ICollection<ElementId> selectionIds)
        {
            _doc = doc;
            _activeView = activeView;
            _selectionIds = selectionIds;
            InitializeComponent();
            ApplyLocalization();

            cmbScope.Items.Add(Loc.S("ValueReplace.Scope.Project"));
            cmbScope.Items.Add(Loc.S("ValueReplace.Scope.ActiveView"));
            cmbScope.Items.Add(Loc.S("ValueReplace.Scope.Selection"));
            // 要素を選択してから起動した場合は「選択要素」を初期値にする
            cmbScope.SelectedIndex = (selectionIds != null && selectionIds.Count > 0) ? 2 : 0;

            // 処理中は閉じさせない（書き込みの途中で閉じると、元に戻す処理が Revit の操作可能な時間外になるため）。
            // 閉じようとしたら中止だけ受け付ける。
            Closing += (s, e) =>
            {
                if (_busy)
                {
                    e.Cancel = true;
                    _cts?.Cancel();
                }
            };
            Closed += (s, e) => _closed = true;
            UpdateCount();
        }

        private void ApplyLocalization()
        {
            Title = Loc.S("ValueReplace.Title");
            txtDescription.Text = Loc.S("ValueReplace.Description");
            lblFind.Text = Loc.S("ValueReplace.Find");
            lblReplace.Text = Loc.S("ValueReplace.Replace");
            txtReplace.ToolTip = Loc.S("ValueReplace.Replace.Tip");
            rbExact.Content = Loc.S("ValueReplace.Match.Exact");
            rbExact.ToolTip = Loc.S("ValueReplace.Match.Exact.Tip");
            rbContains.Content = Loc.S("ValueReplace.Match.Contains");
            rbContains.ToolTip = Loc.S("ValueReplace.Match.Contains.Tip");
            chkIncludeTypes.Content = Loc.S("ValueReplace.IncludeTypes");
            lblScope.Text = Loc.S("ValueReplace.Scope");
            btnSearch.Content = Loc.S("ValueReplace.Btn.Search");

            colSelect.Header = Loc.S("ValueReplace.Col.Select");
            colCategory.Header = Loc.S("ValueReplace.Col.Category");
            colParameter.Header = Loc.S("ValueReplace.Col.Parameter");
            colKind.Header = Loc.S("ValueReplace.Col.Kind");
            colScope.Header = Loc.S("ValueReplace.Col.Scope");
            colCount.Header = Loc.S("ValueReplace.Col.Count");
            colSample.Header = Loc.S("ValueReplace.Col.Sample");

            btnSelectAll.Content = Loc.S("Common.SelectAll");
            btnDeselectAll.Content = Loc.S("Common.SelectNone");
            btnReplace.Content = Loc.S("ValueReplace.Btn.Replace");
            btnClose.Content = Loc.S("ValueReplace.Btn.Close");
        }

        // ===================== 検索 =====================

        private async void Search_Click(object sender, RoutedEventArgs e)
        {
            if (_busy)
            {
                _cts?.Cancel();   // 実行中は「中止」ボタンとして働く
                return;
            }

            string find = txtFind.Text ?? "";
            if (string.IsNullOrWhiteSpace(find))
            {
                MessageBox.Show(this, Loc.S("ValueReplace.FindEmpty"), Loc.S("Common.Confirm"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var scope = (ReplaceScope)Math.Max(0, cmbScope.SelectedIndex);
            if (scope == ReplaceScope.Selection && (_selectionIds == null || _selectionIds.Count == 0))
            {
                MessageBox.Show(this, Loc.S("ValueReplace.NoSelection"), Loc.S("Common.Confirm"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _searchedFind = find;
            _searchedReplace = txtReplace.Text ?? "";
            _searchedMode = rbContains.IsChecked == true ? MatchMode.Contains : MatchMode.Exact;

            var groups = new Dictionary<string, ReplaceGroup>();
            ResultGrid.ItemsSource = null;
            _groups = new List<ReplaceGroup>();

            await RunBusyAsync(Loc.S("ValueReplace.Status.Searching"), ct =>
                ValueReplaceService.Scan(_doc, find, _searchedMode, scope, chkIncludeTypes.IsChecked == true,
                    _activeView, _selectionIds, groups, ct),
                total: 0,
                progressFormat: Loc.S("ValueReplace.Status.SearchingCount"));

            if (_closed) return;

            _groups = groups.Values
                .OrderBy(g => g.CategoryName, StringComparer.CurrentCulture)
                .ThenBy(g => g.ParameterName, StringComparer.CurrentCulture)
                .ThenBy(g => g.IsType)
                .ToList();
            foreach (var g in _groups)
                g.PropertyChanged += (s2, e2) => UpdateCount();
            ResultGrid.ItemsSource = _groups;

            int total = _groups.Sum(g => g.Count);
            txtStatus.Text = string.Format(Loc.S("ValueReplace.Status.Found"), total, _groups.Count);
            DiagLog.Write($"[ValueReplace] 検索 '{find}' ({_searchedMode}, 範囲={scope}) → {total} 件 / {_groups.Count} 種類");
            UpdateCount();
        }

        // ===================== 置換 =====================

        private async void Replace_Click(object sender, RoutedEventArgs e)
        {
            if (_busy) return;

            var selected = _groups.Where(g => g.IsSelected).ToList();
            int count = selected.Sum(g => g.Count);
            if (count == 0) return;

            // 検索後に入力欄を変えていたら、検索し直してもらう（違う条件で書き換えないため）
            if (!string.Equals(txtFind.Text ?? "", _searchedFind, StringComparison.Ordinal)
                || !string.Equals(txtReplace.Text ?? "", _searchedReplace, StringComparison.Ordinal)
                || (rbContains.IsChecked == true ? MatchMode.Contains : MatchMode.Exact) != _searchedMode)
            {
                MessageBox.Show(this, Loc.S("ValueReplace.ConditionChanged"), Loc.S("Common.Confirm"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string newText = string.IsNullOrEmpty(_searchedReplace) ? Loc.S("ValueReplace.EmptyValue") : _searchedReplace;
            var answer = MessageBox.Show(this,
                string.Format(Loc.S("ValueReplace.ConfirmMsg"), count, selected.Count, _searchedFind, newText),
                Loc.S("Common.Confirm"), MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (answer != MessageBoxResult.Yes) return;

            var result = new ReplaceResult();
            var sw = Stopwatch.StartNew();

            // 全体を1つのまとまりにし、中止した場合はすべて元に戻す
            using (var tg = new TransactionGroup(_doc, Loc.S("ValueReplace.Txn")))
            {
                tg.Start();

                await RunBusyAsync(Loc.S("ValueReplace.Status.Replacing"), ct =>
                    ValueReplaceService.Apply(_doc, selected, _searchedFind, _searchedReplace, _searchedMode, result, ct),
                    total: count,
                    progressFormat: Loc.S("ValueReplace.Status.ReplacingCount"));

                if (result.Cancelled || _closed)
                {
                    tg.RollBack();
                    DiagLog.Write($"[ValueReplace] 中止して元に戻しました（{result.Success} 件処理済み）");
                    if (!_closed)
                    {
                        txtStatus.Text = Loc.S("ValueReplace.Status.Cancelled");
                        MessageBox.Show(this, Loc.S("ValueReplace.CancelledMsg"), Loc.S("ValueReplace.Title"),
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    return;
                }

                tg.Assimilate();
            }

            DiagLog.Write($"[ValueReplace] 置換 '{_searchedFind}'→'{_searchedReplace}' 成功={result.Success} 失敗={result.Failed} " +
                          $"対象外={result.Skipped} {sw.ElapsedMilliseconds}ms");

            MessageBox.Show(this,
                string.Format(Loc.S("ValueReplace.DoneMsg"), result.Success, result.Failed, result.Skipped),
                Loc.S("ValueReplace.Title"), MessageBoxButton.OK,
                result.Failed > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);

            // 置換済みの一覧は古くなるので消す（続けて作業する場合は検索し直す）
            _groups = new List<ReplaceGroup>();
            ResultGrid.ItemsSource = null;
            txtStatus.Text = Loc.S("ValueReplace.Status.Done");
            UpdateCount();
        }

        /// <summary>
        /// 反復子の処理を、約50ms ごとに画面へ制御を返しながら最後まで進める。
        /// 実行中は「検索」ボタンが「中止」に変わり、Revit 本体の操作は止める。
        /// </summary>
        private async Task RunBusyAsync(string statusText, Func<CancellationToken, IEnumerable<int>> work,
                                        int total, string progressFormat)
        {
            _busy = true;
            _cts = new CancellationTokenSource();
            var ct = _cts.Token;
            SetBusy(true, total);
            txtStatus.Text = statusText;

            var revitBlock = this.BlockRevitInput();
            try
            {
                await Task.Delay(1);   // 「処理中」の表示を先に描画させる
                var sw = Stopwatch.StartNew();
                foreach (int n in work(ct))
                {
                    if (sw.ElapsedMilliseconds >= 50)
                    {
                        if (total > 0) Progress.Value = Math.Min(n, total);
                        txtStatus.Text = string.Format(progressFormat, n, total);
                        await Task.Delay(1);   // 描画更新・中止の受付
                        sw.Restart();
                    }
                    if (_closed) _cts.Cancel();
                }
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[ValueReplace] 処理中にエラー: {ex}");
                if (!_closed)
                {
                    MessageBox.Show(this, string.Format(Loc.S("ValueReplace.ErrorMsg"), ex.Message),
                        Loc.S("Common.Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            finally
            {
                revitBlock.Dispose();
                _busy = false;
                if (!_closed) SetBusy(false, 0);
            }
        }

        private void SetBusy(bool busy, int total)
        {
            btnSearch.Content = busy ? Loc.S("ValueReplace.Btn.Cancel") : Loc.S("ValueReplace.Btn.Search");
            btnReplace.IsEnabled = !busy && SelectedCount() > 0;
            btnSelectAll.IsEnabled = !busy;
            btnDeselectAll.IsEnabled = !busy;
            btnClose.IsEnabled = !busy;
            txtFind.IsEnabled = !busy;
            txtReplace.IsEnabled = !busy;
            ResultGrid.IsEnabled = !busy;

            Progress.Visibility = busy ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            Progress.IsIndeterminate = busy && total <= 0;
            Progress.Maximum = Math.Max(1, total);
            Progress.Value = 0;
        }

        // ===================== 選択 =====================

        private int SelectedCount() => _groups.Where(g => g.IsSelected).Sum(g => g.Count);

        private void UpdateCount()
        {
            int total = _groups.Sum(g => g.Count);
            int selected = SelectedCount();
            txtCount.Text = string.Format(Loc.S("ValueReplace.Count"), total, selected);
            if (!_busy) btnReplace.IsEnabled = selected > 0;
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var g in _groups) g.IsSelected = true;
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var g in _groups) g.IsSelected = false;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Tools28.Commands.ParameterValueReplace.Services;
using Tools28.Localization;

namespace Tools28.Commands.ParameterValueReplace.Views
{
    /// <summary>
    /// パラメータ値の移動ダイアログ。
    /// ① 「パラメータを読み込む」: 同じ名前でグループ違いの文字パラメータを一覧
    /// ② 移動元/移動先のグループを選び「確認」: 要素ごとの移動内容を数える（まだ書き込まない）
    /// ③ 「移動を実行」: 確認した内容で書き込む（移動先へ書き、移動元を空欄に）
    /// 条件を変えたら確認からやり直す（確認した内容と違う条件で書き込まないため）。
    /// </summary>
    public partial class ParameterValueMoveDialog : Window
    {
        private readonly Document _doc;
        private readonly View _activeView;
        private readonly ICollection<ElementId> _selectionIds;

        private List<MoveRow> _rows = new List<MoveRow>();
        private MovePlan _plan;          // 確認済みの移動内容（条件を変えたら破棄）

        private CancellationTokenSource _cts;
        private bool _busy;
        private bool _closed;
        private bool _suppressInvalidate;

        public ParameterValueMoveDialog(Document doc, View activeView, ICollection<ElementId> selectionIds)
        {
            _doc = doc;
            _activeView = activeView;
            _selectionIds = selectionIds;
            InitializeComponent();
            ApplyLocalization();

            cmbScope.Items.Add(Loc.S("ValueReplace.Scope.Project"));
            cmbScope.Items.Add(Loc.S("ValueReplace.Scope.ActiveView"));
            cmbScope.Items.Add(Loc.S("ValueReplace.Scope.Selection"));
            cmbScope.SelectedIndex = (selectionIds != null && selectionIds.Count > 0) ? 2 : 0;
            cmbScope.SelectionChanged += (s, e) => ClearLoaded();
            chkIncludeTypes.Click += (s, e) => ClearLoaded();

            // 処理中は閉じさせない（書き込み途中で閉じると元に戻す処理が行えなくなるため）
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
            Title = Loc.S("ValueMove.Title");
            txtDescription.Text = Loc.S("ValueMove.Description");
            lblScope.Text = Loc.S("ValueReplace.Scope");
            chkIncludeTypes.Content = Loc.S("ValueReplace.IncludeTypes");
            btnLoad.Content = Loc.S("ValueMove.Btn.Load");
            lblSource.Text = Loc.S("ValueMove.SourceGroup");
            lblTarget.Text = Loc.S("ValueMove.TargetGroup");
            lblExisting.Text = Loc.S("ValueMove.Existing");
            rbKeep.Content = Loc.S("ValueMove.Existing.Keep");
            rbOverwrite.Content = Loc.S("ValueMove.Existing.Overwrite");

            colSelect.Header = Loc.S("ValueMove.Col.Select");
            colName.Header = Loc.S("ValueMove.Col.Name");
            colSourceValue.Header = Loc.S("ValueMove.Col.SourceValue");
            colTargetElem.Header = Loc.S("ValueMove.Col.TargetElem");
            colMovable.Header = Loc.S("ValueMove.Col.Movable");
            colNoTarget.Header = Loc.S("ValueMove.Col.NoTarget");
            colTargetHas.Header = Loc.S("ValueMove.Col.TargetHasValue");
            colConflict.Header = Loc.S("ValueMove.Col.Conflict");

            btnSelectAll.Content = Loc.S("Common.SelectAll");
            btnDeselectAll.Content = Loc.S("Common.SelectNone");
            btnPreview.Content = Loc.S("ValueMove.Btn.Preview");
            btnMove.Content = Loc.S("ValueMove.Btn.Move");
            btnClose.Content = Loc.S("ValueReplace.Btn.Close");
        }

        private ReplaceScope Scope => (ReplaceScope)Math.Max(0, cmbScope.SelectedIndex);
        private bool IncludeTypes => chkIncludeTypes.IsChecked == true;
        private string SourceGroup => cmbSource.SelectedItem as string;
        private string TargetGroup => cmbTarget.SelectedItem as string;
        private ExistingTargetMode Mode => rbOverwrite.IsChecked == true ? ExistingTargetMode.Overwrite : ExistingTargetMode.Keep;

        // ===================== ① 読み込み =====================

        private async void Load_Click(object sender, RoutedEventArgs e)
        {
            if (_busy)
            {
                _cts?.Cancel();   // 実行中は「中止」ボタンとして働く
                return;
            }
            if (Scope == ReplaceScope.Selection && (_selectionIds == null || _selectionIds.Count == 0))
            {
                MessageBox.Show(this, Loc.S("ValueReplace.NoSelection"), Loc.S("Common.Confirm"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var rows = new Dictionary<string, MoveRow>();
            bool completed = await RunBusyAsync(Loc.S("ValueMove.Status.Loading"), ct =>
                ValueMoveService.Load(_doc, Scope, IncludeTypes, _activeView, _selectionIds, rows, ct),
                total: 0, progressFormat: Loc.S("ValueMove.Status.LoadingCount"));
            if (_closed) return;
            if (!completed)
            {
                ClearLoaded();
                return;
            }

            _rows = rows.Values.OrderBy(r => r.ParameterName, StringComparer.CurrentCulture).ToList();
            foreach (var r in _rows)
            {
                r.PropertyChanged += (s2, e2) =>
                {
                    if (e2.PropertyName == nameof(MoveRow.IsSelected)) InvalidatePlan();
                };
            }

            // グループの候補（同名パラメータが分かれているグループ）
            var groups = _rows.SelectMany(r => r.Groups.Keys).Distinct()
                .OrderBy(g => g, StringComparer.CurrentCulture).ToList();
            _suppressInvalidate = true;
            cmbSource.ItemsSource = groups;
            cmbTarget.ItemsSource = groups;
            cmbSource.SelectedIndex = groups.Count > 0 ? 0 : -1;
            cmbTarget.SelectedIndex = groups.Count > 1 ? 1 : -1;
            _suppressInvalidate = false;

            ApplyGroupFilter();
            txtStatus.Text = string.Format(Loc.S("ValueMove.Status.Loaded"), _rows.Count);
            DiagLog.Write($"[ValueMove] 読み込み: 同名でグループ違いのパラメータ {_rows.Count} 種類 / グループ {groups.Count} 種類");
        }

        private void ClearLoaded()
        {
            _rows = new List<MoveRow>();
            _plan = null;
            ResultGrid.ItemsSource = null;
            cmbSource.ItemsSource = null;
            cmbTarget.ItemsSource = null;
            UpdateCount();
        }

        /// <summary>選んだ移動元/移動先グループの両方にあるパラメータ名だけを一覧に出す</summary>
        private void ApplyGroupFilter()
        {
            string src = SourceGroup, tgt = TargetGroup;
            var visible = new List<MoveRow>();
            if (src != null && tgt != null && src != tgt)
            {
                foreach (var r in _rows)
                {
                    if (!r.Groups.TryGetValue(src, out var s) || !r.Groups.TryGetValue(tgt, out var t))
                        continue;
                    r.SourceValueCount = s.ValueCount;
                    r.TargetElementCount = t.ElementCount;
                    r.ClearPreview();
                    if (s.ValueCount > 0) visible.Add(r);
                }
            }
            ResultGrid.ItemsSource = visible;
            InvalidatePlan();
        }

        private void Group_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressInvalidate) return;
            ApplyGroupFilter();
        }

        private void Option_Changed(object sender, RoutedEventArgs e)
        {
            // XAML の IsChecked 初期値で画面の構築途中にも呼ばれるため、構築完了までは何もしない
            if (ResultGrid == null || txtCount == null || btnMove == null) return;
            InvalidatePlan();
        }

        /// <summary>条件が変わったので、確認済みの内容を破棄する（もう一度「確認」が必要）</summary>
        private void InvalidatePlan()
        {
            if (_suppressInvalidate) return;
            _plan = null;
            foreach (var r in VisibleRows()) r.ClearPreview();
            UpdateCount();
        }

        private IEnumerable<MoveRow> VisibleRows()
            => (ResultGrid.ItemsSource as IEnumerable<MoveRow>) ?? Enumerable.Empty<MoveRow>();

        // ===================== ② 確認 =====================

        private async void Preview_Click(object sender, RoutedEventArgs e)
        {
            if (_busy) return;
            var names = VisibleRows().Where(r => r.IsSelected).Select(r => r.ParameterName).ToList();
            if (names.Count == 0 || SourceGroup == null || TargetGroup == null || SourceGroup == TargetGroup) return;

            var plan = new MovePlan();
            string src = SourceGroup, tgt = TargetGroup;
            var mode = Mode;
            bool completed = await RunBusyAsync(Loc.S("ValueMove.Status.Planning"), ct =>
                ValueMoveService.Plan(_doc, Scope, IncludeTypes, _activeView, _selectionIds,
                    names, src, tgt, mode, plan, ct),
                total: 0, progressFormat: Loc.S("ValueMove.Status.PlanningCount"));
            if (_closed || !completed) return;

            foreach (var r in VisibleRows())
            {
                if (!r.IsSelected) { r.ClearPreview(); continue; }
                r.MovableText = Count(plan.Movable, r.ParameterName);
                r.NoTargetText = Count(plan.NoTarget, r.ParameterName);
                r.TargetHasValueText = Count(plan.TargetHasValue, r.ParameterName)
                    + (plan.TargetHasValue.ContainsKey(r.ParameterName)
                        ? (mode == ExistingTargetMode.Keep ? Loc.S("ValueMove.KeptSuffix") : Loc.S("ValueMove.OverwriteSuffix"))
                        : "");
                r.ConflictText = Count(plan.Conflict, r.ParameterName);
            }
            _plan = plan;
            txtStatus.Text = string.Format(Loc.S("ValueMove.Status.Planned"),
                plan.Ops.Count, plan.NoTarget.Values.Sum(), plan.Conflict.Values.Sum());
            DiagLog.Write($"[ValueMove] 確認 '{src}'→'{tgt}' mode={mode} 移動={plan.Ops.Count} 移動先なし={plan.NoTarget.Values.Sum()} " +
                          $"移動先に値あり={plan.TargetHasValue.Values.Sum()} 食い違い={plan.Conflict.Values.Sum()}");
            UpdateCount();
        }

        private static string Count(Dictionary<string, int> d, string key)
            => d.TryGetValue(key, out int n) ? n.ToString("N0") : "0";

        // ===================== ③ 実行 =====================

        private async void Move_Click(object sender, RoutedEventArgs e)
        {
            if (_busy || _plan == null || _plan.Ops.Count == 0) return;

            var answer = MessageBox.Show(this,
                string.Format(Loc.S("ValueMove.ConfirmMsg"), _plan.Ops.Count, SourceGroup, TargetGroup),
                Loc.S("Common.Confirm"), MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (answer != MessageBoxResult.Yes) return;

            var ops = _plan.Ops;
            var result = new MoveResult();
            var sw = Stopwatch.StartNew();

            // 全体を1つのまとまりにし、中止した場合はすべて元に戻す
            using (var tg = new TransactionGroup(_doc, Loc.S("ValueMove.Txn")))
            {
                tg.Start();
                await RunBusyAsync(Loc.S("ValueMove.Status.Moving"), ct =>
                    ValueMoveService.Apply(_doc, ops, result, ct),
                    total: ops.Count, progressFormat: Loc.S("ValueMove.Status.MovingCount"));

                if (result.Cancelled || _closed)
                {
                    tg.RollBack();
                    DiagLog.Write($"[ValueMove] 中止して元に戻しました（{result.Success} 件処理済み）");
                    if (!_closed)
                    {
                        txtStatus.Text = Loc.S("ValueReplace.Status.Cancelled");
                        MessageBox.Show(this, Loc.S("ValueReplace.CancelledMsg"), Loc.S("ValueMove.Title"),
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    return;
                }
                tg.Assimilate();
            }

            DiagLog.Write($"[ValueMove] 移動 成功={result.Success} 失敗={result.Failed} 対象外={result.Skipped} {sw.ElapsedMilliseconds}ms");
            MessageBox.Show(this,
                string.Format(Loc.S("ValueMove.DoneMsg"), result.Success, result.Failed, result.Skipped),
                Loc.S("ValueMove.Title"), MessageBoxButton.OK,
                result.Failed > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);

            // 値が変わったので一覧は読み込み直しが必要
            ClearLoaded();
            txtStatus.Text = Loc.S("ValueMove.Status.Done");
        }

        // ===================== 共通 =====================

        /// <summary>
        /// 反復子の処理を、約50ms ごとに画面へ制御を返しながら最後まで進める。
        /// 最後まで終わったら true、中止・エラーなら false。
        /// </summary>
        private async Task<bool> RunBusyAsync(string statusText, Func<CancellationToken, IEnumerable<int>> work,
                                              int total, string progressFormat)
        {
            _busy = true;
            _cts = new CancellationTokenSource();
            var ct = _cts.Token;
            SetBusy(true, total);
            txtStatus.Text = statusText;
            bool ok = false;

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
                }
                ok = !ct.IsCancellationRequested;
                if (!ok && !_closed) txtStatus.Text = Loc.S("ValueMove.Status.Stopped");
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[ValueMove] 処理中にエラー: {ex}");
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
            return ok;
        }

        private void SetBusy(bool busy, int total)
        {
            btnLoad.Content = busy ? Loc.S("ValueReplace.Btn.Cancel") : Loc.S("ValueMove.Btn.Load");
            cmbScope.IsEnabled = !busy;
            chkIncludeTypes.IsEnabled = !busy;
            cmbSource.IsEnabled = !busy;
            cmbTarget.IsEnabled = !busy;
            rbKeep.IsEnabled = !busy;
            rbOverwrite.IsEnabled = !busy;
            ResultGrid.IsEnabled = !busy;
            btnSelectAll.IsEnabled = !busy;
            btnDeselectAll.IsEnabled = !busy;
            btnClose.IsEnabled = !busy;

            Progress.Visibility = busy ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            Progress.IsIndeterminate = busy && total <= 0;
            Progress.Maximum = Math.Max(1, total);
            Progress.Value = 0;
            UpdateCount();
        }

        private void UpdateCount()
        {
            var visible = VisibleRows().ToList();
            int selected = visible.Count(r => r.IsSelected);
            txtCount.Text = _plan != null
                ? string.Format(Loc.S("ValueMove.CountPlanned"), selected, _plan.Ops.Count)
                : string.Format(Loc.S("ValueMove.Count"), visible.Count, selected);

            bool groupsOk = SourceGroup != null && TargetGroup != null && SourceGroup != TargetGroup;
            btnPreview.IsEnabled = !_busy && groupsOk && selected > 0;
            btnMove.IsEnabled = !_busy && _plan != null && _plan.Ops.Count > 0;
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in VisibleRows()) r.IsSelected = true;
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in VisibleRows()) r.IsSelected = false;
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}

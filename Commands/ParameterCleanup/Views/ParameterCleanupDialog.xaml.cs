using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.ParameterCleanup.Models;
using Tools28.Commands.ParameterCleanup.Services;
using Tools28.Localization;

namespace Tools28.Commands.ParameterCleanup.Views
{
    public partial class ParameterCleanupDialog : Window
    {
        private readonly Document _doc;
        private readonly ParameterScanner _scanner = new ParameterScanner();

        private List<ParamRow> _rows = new List<ParamRow>();
        private ICollectionView _view;

        private CancellationTokenSource _cts;
        private bool _scanning;
        private bool _valuesChecked;

        public ParameterCleanupDialog(Document doc)
        {
            _doc = doc;
            InitializeComponent();
            ApplyLocalization();
            LoadRows();
        }

        private void ApplyLocalization()
        {
            Title = Loc.S("ParamCleanup.Title");
            txtDescription.Text = Loc.S("ParamCleanup.Description");
            lblSearch.Text = Loc.S("ParamCleanup.Search");
            chkDuplicateOnly.Content = Loc.S("ParamCleanup.DuplicateOnly");
            chkEmptyOnly.Content = Loc.S("ParamCleanup.EmptyOnly");
            btnSelectEmpty.Content = Loc.S("ParamCleanup.SelectEmpty");

            // 列ヘッダー（DataGrid の Header は string 直指定）
            ParamGrid.Columns[0].Header = Loc.S("ParamCleanup.Col.Select");
            colName.Header = Loc.S("ParamCleanup.Col.Name");
            colKind.Header = Loc.S("ParamCleanup.Col.Kind");
            colScope.Header = Loc.S("ParamCleanup.Col.Scope");
            colCategories.Header = Loc.S("ParamCleanup.Col.Categories");
            colState.Header = Loc.S("ParamCleanup.Col.Value");

            btnCheck.Content = Loc.S("ParamCleanup.Btn.Check");
            btnDelete.Content = Loc.S("ParamCleanup.Btn.Delete");
            btnClose.Content = Loc.S("ParamCleanup.Btn.Close");
        }

        private void LoadRows()
        {
            _valuesChecked = false;
            _rows = _scanner.EnumerateParameters(_doc);
            _view = CollectionViewSource.GetDefaultView(_rows);
            _view.Filter = RowFilter;
            ParamGrid.ItemsSource = _view;
            UpdateCount();
        }

        private bool RowFilter(object o)
        {
            if (!(o is ParamRow r)) return false;

            string q = SearchBox?.Text?.Trim();
            if (!string.IsNullOrEmpty(q) &&
                (r.Name == null || r.Name.IndexOf(q, StringComparison.CurrentCultureIgnoreCase) < 0))
                return false;

            if (chkDuplicateOnly?.IsChecked == true && !r.IsDuplicateName)
                return false;

            if (chkEmptyOnly?.IsChecked == true && r.State != ValueState.Empty)
                return false;

            return true;
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            _view?.Refresh();
            UpdateCount();
        }

        private void UpdateCount()
        {
            int total = _rows.Count;
            int shown = _view?.Cast<object>().Count() ?? total;
            int selected = _rows.Count(r => r.IsSelected);
            txtCount.Text = string.Format(Loc.S("ParamCleanup.Count.Summary"), total, shown, selected);
        }

        // ===== 値の有無を確認（スキャン）=====
        private async void Check_Click(object sender, RoutedEventArgs e)
        {
            if (_scanning)
            {
                _cts?.Cancel();   // 実行中はキャンセルボタンとして機能
                return;
            }

            var targets = _rows.Where(r => r.IsScannable).ToList();
            if (targets.Count == 0)
            {
                txtStatus.Text = Loc.S("ParamCleanup.Status.Done");
                return;
            }

            _scanning = true;
            _cts = new CancellationTokenSource();
            var ct = _cts.Token;

            btnCheck.Content = Loc.S("ParamCleanup.Btn.Cancel");
            SetBusy(true);
            Progress.Visibility = Visibility.Visible;
            Progress.Value = 0;
            Progress.Maximum = Math.Max(1, targets.Count);

            int done = 0;
            // 待機オーバーヘッドを抑えるため、一定時間（約50ms）ごとにだけ UI へ制御を返す。
            // Task.Delay は Background 優先度で復帰するため、その間に「中止」ボタンの
            // クリック（Input 優先度）が確実に処理される。
            var sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                foreach (var row in targets)
                {
                    if (ct.IsCancellationRequested) break;

                    row.State = ValueState.Checking;

                    foreach (var _ in _scanner.ScanRow(_doc, row, ct))
                    {
                        if (sw.ElapsedMilliseconds >= 50)
                        {
                            txtStatus.Text = string.Format(Loc.S("ParamCleanup.Status.Scanning"), row.Name);
                            Progress.Value = done;
                            txtCount.Text = string.Format(Loc.S("ParamCleanup.Count.Progress"), done, targets.Count);
                            await Task.Delay(1);   // UI へ制御を返す（描画更新・中止受付）
                            sw.Restart();
                            if (ct.IsCancellationRequested) break;
                        }
                    }

                    if (ct.IsCancellationRequested) break;
                    done++;
                }
            }
            finally
            {
                // キャンセルで中断した行を未確認へ戻す
                foreach (var r in targets)
                    if (r.State == ValueState.Checking) r.State = ValueState.Unchecked;

                _scanner.ClearCache();
                bool cancelled = ct.IsCancellationRequested;
                _valuesChecked = !cancelled;
                _scanning = false;

                btnCheck.Content = Loc.S("ParamCleanup.Btn.Check");
                SetBusy(false);
                Progress.Visibility = Visibility.Collapsed;
                txtStatus.Text = cancelled
                    ? Loc.S("ParamCleanup.Status.Cancelled")
                    : Loc.S("ParamCleanup.Status.Done");

                _view?.Refresh();
                UpdateCount();
            }
        }

        private void SetBusy(bool busy)
        {
            // スキャン中は btnCheck（=中止）と btnClose のみ有効
            btnDelete.IsEnabled = !busy;
            btnSelectEmpty.IsEnabled = !busy;
        }

        private void SelectEmpty_Click(object sender, RoutedEventArgs e)
        {
            if (!_valuesChecked)
            {
                TaskDialog.Show(Loc.S("ParamCleanup.Title"), Loc.S("ParamCleanup.NotChecked.Msg"));
                return;
            }

            foreach (var r in _rows)
                r.IsSelected = r.State == ValueState.Empty;

            _view?.Refresh();
            UpdateCount();
        }

        // ===== 削除 =====
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            var selected = _rows.Where(r => r.IsSelected).ToList();
            if (selected.Count == 0)
            {
                TaskDialog.Show(Loc.S("ParamCleanup.Title"), Loc.S("ParamCleanup.NoSelection.Msg"));
                return;
            }

            var confirm = new TaskDialog(Loc.S("ParamCleanup.Confirm.Title"))
            {
                MainInstruction = string.Format(Loc.S("ParamCleanup.Confirm.Main"), selected.Count),
                MainContent = Loc.S("ParamCleanup.Confirm.Content"),
                CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                DefaultButton = TaskDialogResult.No
            };
            if (confirm.Show() != TaskDialogResult.Yes)
                return;

            var ids = selected.Select(r => r.Id).Where(id => id != null && id != ElementId.InvalidElementId).ToList();
            int ok = 0, fail = 0;

            try
            {
                using (var t = new Transaction(_doc, Loc.S("ParamCleanup.Txn")))
                {
                    t.Start();
                    ICollection<ElementId> deleted = null;
                    try
                    {
                        deleted = _doc.Delete(ids);
                    }
                    catch
                    {
                        // 一括削除に失敗した場合は個別に試行
                        deleted = new List<ElementId>();
                        foreach (var id in ids)
                        {
                            try
                            {
                                var r = _doc.Delete(id);
                                if (r != null)
                                    foreach (var did in r) deleted.Add(did);
                            }
                            catch { }
                        }
                    }
                    t.Commit();

                    var deletedSet = new HashSet<ElementId>(deleted ?? new List<ElementId>());
                    ok = ids.Count(id => deletedSet.Contains(id));
                    fail = ids.Count - ok;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show(Loc.S("Common.Error"),
                    string.Format(Loc.S("ParamCleanup.Result.Error"), ex.Message));
                return;
            }

            TaskDialog.Show(Loc.S("ParamCleanup.Result.Title"),
                string.Format(Loc.S("ParamCleanup.Result.Msg"), ok, fail));

            LoadRows();   // 一覧を再構築
            txtStatus.Text = "";
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (_scanning)
                _cts?.Cancel();
            base.OnClosing(e);
        }
    }
}

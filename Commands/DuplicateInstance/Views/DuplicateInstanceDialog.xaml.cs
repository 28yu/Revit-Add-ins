using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.DuplicateInstance.Models;
using Tools28.Commands.DuplicateInstance.Services;
using Tools28.Localization;

namespace Tools28.Commands.DuplicateInstance.Views
{
    /// <summary>
    /// 「同じ場所に同一のインスタンスがあります」警告の要素を一覧にし、
    /// 組ごとのパラメータを見比べて、不要な方を削除するダイアログ。
    /// </summary>
    public partial class DuplicateInstanceDialog : Window
    {
        private readonly Document _doc;
        private readonly DuplicateDialogState _state;
        private readonly DuplicateScanner _scanner = new DuplicateScanner();

        private List<DuplicateRow> _rows = new List<DuplicateRow>();
        private ICollectionView _view;
        private List<CompareRow> _compareRows = new List<CompareRow>();
        private ICollectionView _compareView;
        private int _compareGroupNo = -1;
        private bool _ready;

        private string AllCategories => Loc.S("DupInst.Filter.AllCategories");

        public DuplicateInstanceDialog(Document doc, DuplicateDialogState state)
        {
            _doc = doc;
            _state = state ?? new DuplicateDialogState();
            InitializeComponent();
            ApplyLocalization();
            RestoreBounds();

            SearchBox.Text = _state.SearchText ?? "";
            chkDiffOnly.IsChecked = _state.DiffOnly;

            LoadRows();
            _ready = true;
            RefreshView();
            RestoreSelection();

            Closing += (s, e) => SaveState();
        }

        private void ApplyLocalization()
        {
            Title = Loc.S("DupInst.Title");
            txtDescription.Text = Loc.S("DupInst.Description");
            lblSearch.Text = Loc.S("DupInst.Search");
            lblCategory.Text = Loc.S("DupInst.Filter.Category");
            chkCheckedOnly.Content = Loc.S("DupInst.Filter.CheckedOnly");

            colCheck.Header = Loc.S("DupInst.Col.Delete");
            colGroup.Header = Loc.S("DupInst.Col.Group");
            colId.Header = Loc.S("DupInst.Col.Id");
            colCategory.Header = Loc.S("DupInst.Col.Category");
            colFamily.Header = Loc.S("DupInst.Col.Family");
            colType.Header = Loc.S("DupInst.Col.Type");
            colLevel.Header = Loc.S("DupInst.Col.Level");
            colWorkset.Header = Loc.S("DupInst.Col.Workset");
            colMark.Header = Loc.S("DupInst.Col.Mark");
            colComments.Header = Loc.S("DupInst.Col.Comments");
            colHost.Header = Loc.S("DupInst.Col.Host");
            colFilled.Header = Loc.S("DupInst.Col.Filled");
            colDiff.Header = Loc.S("DupInst.Col.Diff");

            lblParamSearch.Text = Loc.S("DupInst.Compare.ParamSearch");
            chkDiffOnly.Content = Loc.S("DupInst.Compare.DiffOnly");

            btnAutoSelect.Content = Loc.S("DupInst.Btn.AutoSelect");
            btnAutoSelect.ToolTip = Loc.S("DupInst.Btn.AutoSelect.Tip");
            btnDeselectAll.Content = Loc.S("DupInst.Btn.DeselectAll");
            btnDeselectAll.ToolTip = Loc.S("DupInst.Btn.DeselectAll.Tip");
            btnShow.Content = Loc.S("DupInst.Btn.Show");
            btnShow.ToolTip = Loc.S("DupInst.Btn.Show.Tip");
            btnReload.Content = Loc.S("DupInst.Btn.Reload");
            btnReload.ToolTip = Loc.S("Common.Reload.Tip");
            btnDelete.Content = Loc.S("DupInst.Btn.Delete");
            btnClose.Content = Loc.S("DupInst.Btn.Close");
        }

        // ===== 読み込み・絞り込み =====

        private void LoadRows()
        {
            using (this.BlockRevitInput())
            {
                _rows = _scanner.Scan(_doc);
            }

            foreach (var r in _rows)
                r.IsSelected = _state.CheckedIds.Contains(r.IdValue);

            _view = CollectionViewSource.GetDefaultView(_rows);
            _view.Filter = RowFilter;
            RowGrid.ItemsSource = _view;

            // カテゴリの選択肢を作り直す（前回の選択を可能なら残す）
            bool wasReady = _ready;
            _ready = false;
            string prev = cmbCategory.SelectedItem as string ?? _state.Category;
            var cats = new List<string> { AllCategories };
            cats.AddRange(_rows.Select(r => r.Category).Where(c => !string.IsNullOrEmpty(c)).Distinct().OrderBy(c => c));
            cmbCategory.ItemsSource = cats;
            cmbCategory.SelectedItem = (prev != null && cats.Contains(prev)) ? prev : AllCategories;
            _ready = wasReady;

            _compareGroupNo = -1;
            UpdateCompare();
            UpdateCount();
        }

        private bool RowFilter(object o)
        {
            if (!(o is DuplicateRow r)) return false;

            string cat = cmbCategory?.SelectedItem as string;
            if (!string.IsNullOrEmpty(cat) && cat != AllCategories && r.Category != cat) return false;

            if (chkCheckedOnly?.IsChecked == true && !r.IsSelected) return false;

            string q = SearchBox?.Text?.Trim();
            if (!string.IsNullOrEmpty(q))
            {
                bool hit = new[] { r.IdValue.ToString(), r.Category, r.Family, r.TypeName, r.Level,
                                   r.Workset, r.Mark, r.Comments, r.Host }
                    .Any(s => s != null && s.IndexOf(q, StringComparison.CurrentCultureIgnoreCase) >= 0);
                if (!hit) return false;
            }
            return true;
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            if (!_ready) return;
            RefreshView();
        }

        private void RefreshView()
        {
            _view?.Refresh();
            UpdateCount();
        }

        private void UpdateCount()
        {
            int groups = _rows.Select(r => r.GroupNo).Distinct().Count();
            int shown = _view?.Cast<object>().Count() ?? _rows.Count;
            int selected = _rows.Count(r => r.IsSelected);
            txtCount.Text = string.Format(Loc.S("DupInst.Count.Summary"), groups, _rows.Count, shown, selected);
            btnDelete.IsEnabled = selected > 0;
        }

        private void RowCheck_Click(object sender, RoutedEventArgs e)
        {
            // チェックの変更を比較表の見出し（削除予定の印）と件数に反映する
            _compareGroupNo = -1;
            UpdateCompare();
            UpdateCount();
        }

        // ===== 比較表 =====

        private void RowGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_ready) return;
            UpdateCompare();
        }

        private DuplicateRow CurrentRow => RowGrid.SelectedItem as DuplicateRow
                                           ?? RowGrid.SelectedItems.OfType<DuplicateRow>().FirstOrDefault();

        /// <summary>一覧で選んでいる行の組について、全パラメータを横に並べた比較表を作る。</summary>
        private void UpdateCompare()
        {
            var cur = CurrentRow;
            if (cur == null)
            {
                _compareGroupNo = -1;
                CompareGrid.Columns.Clear();
                CompareGrid.ItemsSource = null;
                txtCompareTitle.Text = _rows.Count == 0 ? Loc.S("DupInst.NoWarnings") : Loc.S("DupInst.Compare.Hint");
                return;
            }
            if (cur.GroupNo == _compareGroupNo) return;
            _compareGroupNo = cur.GroupNo;

            var members = _rows.Where(r => r.GroupNo == cur.GroupNo).OrderBy(r => r.IdValue).ToList();
            txtCompareTitle.Text = string.Format(Loc.S("DupInst.Compare.Title"), cur.GroupNo, members.Count);

            // 列: パラメータ名 + 要素ごと
            CompareGrid.Columns.Clear();
            CompareGrid.Columns.Add(new DataGridTextColumn
            {
                Header = Loc.S("DupInst.Compare.ColParam"),
                Binding = new Binding(nameof(CompareRow.Name)),
                Width = new DataGridLength(220),
                ElementStyle = (Style)FindResource("CellText")
            });
            for (int i = 0; i < members.Count; i++)
            {
                var m = members[i];
                string header = string.Format(Loc.S("DupInst.Compare.ColElement"), m.IdValue);
                if (m.IsSelected) header += " " + Loc.S("DupInst.Compare.DeleteMark");
                CompareGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = header,
                    Binding = new Binding($"Values[{i}]"),
                    Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                    MinWidth = 140,
                    ElementStyle = (Style)FindResource("CellText")
                });
            }

            // 行: 先頭に基本情報、続いて全パラメータ（名前順）
            var list = new List<CompareRow>
            {
                Header(Loc.S("DupInst.Col.Id"), members.Select(m => m.IdValue.ToString())),
                Header(Loc.S("DupInst.Col.Category"), members.Select(m => m.Category)),
                Header(Loc.S("DupInst.Col.Family"), members.Select(m => m.Family)),
                Header(Loc.S("DupInst.Col.Type"), members.Select(m => m.TypeName)),
                Header(Loc.S("DupInst.Col.Workset"), members.Select(m => m.Workset)),
                Header(Loc.S("DupInst.Col.Host"), members.Select(m => m.Host)),
                Header(Loc.S("DupInst.Col.Filled"), members.Select(m => m.FilledCount.ToString())),
            };
            var keys = members.SelectMany(m => m.Parameters.Keys).Distinct()
                              .OrderBy(k => k, StringComparer.CurrentCulture).ToList();
            foreach (var k in keys)
            {
                var vals = members.Select(m => m.Parameters.TryGetValue(k, out var v) ? v : "").ToArray();
                list.Add(new CompareRow
                {
                    Name = k,
                    Values = vals,
                    IsDiff = vals.Distinct().Count() > 1
                });
            }
            _compareRows = list;
            _compareView = CollectionViewSource.GetDefaultView(_compareRows);
            _compareView.Filter = CompareFilter;
            CompareGrid.ItemsSource = _compareView;
        }

        private static CompareRow Header(string name, IEnumerable<string> values)
        {
            var arr = values.Select(v => v ?? "").ToArray();
            return new CompareRow { Name = name, Values = arr, IsHeader = true, IsDiff = arr.Distinct().Count() > 1 };
        }

        private bool CompareFilter(object o)
        {
            if (!(o is CompareRow r)) return false;
            if (r.IsHeader) return true;   // 基本情報は常に表示
            if (chkDiffOnly?.IsChecked == true && !r.IsDiff) return false;
            string q = ParamSearchBox?.Text?.Trim();
            if (!string.IsNullOrEmpty(q) &&
                (r.Name == null || r.Name.IndexOf(q, StringComparison.CurrentCultureIgnoreCase) < 0))
                return false;
            return true;
        }

        private void Compare_Changed(object sender, RoutedEventArgs e)
        {
            if (!_ready) return;
            _compareView?.Refresh();
        }

        // ===== チェック操作 =====

        /// <summary>
        /// 表示中の組ごとに、値が入っているパラメータが一番多い要素を残し、それ以外にチェックを付ける。
        /// 値の数が同じなら、先に作られた要素（要素IDが小さい方）を残す。
        /// </summary>
        private void AutoSelect_Click(object sender, RoutedEventArgs e)
        {
            var visibleGroups = new HashSet<int>(_view.Cast<object>().OfType<DuplicateRow>().Select(r => r.GroupNo));
            foreach (var g in _rows.Where(r => visibleGroups.Contains(r.GroupNo)).GroupBy(r => r.GroupNo))
            {
                var keep = g.OrderByDescending(r => r.FilledCount).ThenBy(r => r.IdValue).First();
                foreach (var r in g) r.IsSelected = r != keep;
            }
            _compareGroupNo = -1;
            UpdateCompare();
            RefreshView();
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in _rows) r.IsSelected = false;
            _compareGroupNo = -1;
            UpdateCompare();
            RefreshView();
        }

        // ===== Revit 画面で確認 =====

        private void RowGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // チェックボックスのダブルクリックでは画面表示しない
            if (e.OriginalSource is DependencyObject d && FindParent<CheckBox>(d) != null) return;
            if (FindParent<DataGridRow>(e.OriginalSource as DependencyObject) == null) return;
            RequestShow();
        }

        private void Show_Click(object sender, RoutedEventArgs e) => RequestShow();

        /// <summary>選んだ行（選んでいなければその組全体）を Revit 画面で表示するため、ダイアログを一度閉じる。</summary>
        private void RequestShow()
        {
            var sel = RowGrid.SelectedItems.OfType<DuplicateRow>().ToList();
            if (sel.Count == 0)
            {
                TaskDialog.Show(Loc.S("DupInst.Title"), Loc.S("DupInst.Show.NoSelection"));
                this.BringToFrontDeferred();
                return;
            }
            // 1行だけ選んでいるときは、重なっている相手も一緒に選択して見せる
            List<DuplicateRow> targets = sel.Count == 1
                ? _rows.Where(r => r.GroupNo == sel[0].GroupNo).ToList()
                : sel;
            _state.ShowRequest = targets.Select(r => r.Id).ToList();
            Close();
        }

        private static T FindParent<T>(DependencyObject d) where T : DependencyObject
        {
            while (d != null)
            {
                if (d is T t) return t;
                d = System.Windows.Media.VisualTreeHelper.GetParent(d);
            }
            return null;
        }

        // ===== 削除 =====

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            var selected = _rows.Where(r => r.IsSelected).ToList();
            if (selected.Count == 0)
            {
                TaskDialog.Show(Loc.S("DupInst.Title"), Loc.S("DupInst.NoSelection.Msg"));
                this.BringToFrontDeferred();
                return;
            }

            // 組の全員にチェックが付いていると、その場所の要素が1つも残らない
            var allGone = _rows.GroupBy(r => r.GroupNo).Where(g => g.All(r => r.IsSelected)).Select(g => g.Key).ToList();

            var confirm = new TaskDialog(Loc.S("DupInst.Confirm.Title"))
            {
                MainInstruction = string.Format(Loc.S("DupInst.Confirm.Main"), selected.Count),
                MainContent = allGone.Count > 0
                    ? string.Format(Loc.S("DupInst.Confirm.ContentAllGone"), string.Join(", ", allGone.Take(20)))
                    : Loc.S("DupInst.Confirm.Content"),
                CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                DefaultButton = TaskDialogResult.No
            };
            if (confirm.Show() != TaskDialogResult.Yes)
            {
                this.BringToFrontDeferred();
                return;
            }

            var ids = selected.Select(r => r.Id).ToList();
            int ok = 0, fail = 0;
            List<string> warnings = null;
            try
            {
                using (var t = new Transaction(_doc, Loc.S("DupInst.Txn.Delete")))
                {
                    t.Start();
                    var fho = t.GetFailureHandlingOptions();
                    fho = fho.SetForcedModalHandling(false);
                    fho = fho.SetClearAfterRollback(true);
                    var swallower = new WarningCollector();
                    warnings = swallower.Messages;
                    fho = fho.SetFailuresPreprocessor(swallower);
                    t.SetFailureHandlingOptions(fho);

                    // 1件ずつ消す（1つが失敗しても他は消せるように）
                    foreach (var id in ids)
                    {
                        try
                        {
                            if (_doc.GetElement(id) == null) { fail++; continue; }
                            _doc.Delete(id);
                            ok++;
                        }
                        catch { fail++; }
                    }
                    t.CommitOrThrow();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show(Loc.S("Common.Error"), string.Format(Loc.S("DupInst.Result.Error"), ex.Message));
                this.BringToFrontDeferred();
                return;
            }

            string msg = string.Format(Loc.S("DupInst.Result.Msg"), ok, fail);
            if (warnings != null && warnings.Count > 0)
            {
                msg += "\n\n" + Loc.S("Common.RevitWarnings") + "\n"
                     + string.Join("\n", warnings.Take(10).Select(w => "・" + w));
                if (warnings.Count > 10) msg += "\n…";
            }
            TaskDialog.Show(Loc.S("DupInst.Result.Title"), msg);

            foreach (var r in selected) _state.CheckedIds.Remove(r.IdValue);
            foreach (var r in _rows) r.IsSelected = false;   // 再読み込み時に残っている要素へチェックを戻さない
            LoadRows();
            RefreshView();
            this.BringToFrontDeferred();
        }

        private void Reload_Click(object sender, RoutedEventArgs e)
        {
            CaptureChecks();
            LoadRows();
            RefreshView();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            _state.ShowRequest = null;
            Close();
        }

        // ===== 状態の保存・復元（「画面で確認」で開き直すため） =====

        private void CaptureChecks()
        {
            _state.CheckedIds.Clear();
            foreach (var r in _rows.Where(r => r.IsSelected)) _state.CheckedIds.Add(r.IdValue);
        }

        private void SaveState()
        {
            CaptureChecks();
            _state.SearchText = SearchBox.Text ?? "";
            _state.Category = cmbCategory.SelectedItem as string;
            _state.DiffOnly = chkDiffOnly.IsChecked == true;
            _state.SelectedId = CurrentRow?.IdValue ?? -1;
            if (WindowState == WindowState.Normal)
            {
                _state.HasBounds = true;
                _state.Left = Left; _state.Top = Top; _state.Width = Width; _state.Height = Height;
            }
            _state.TopPaneHeight = TopPaneRow.ActualHeight;
        }

        private void RestoreBounds()
        {
            if (!_state.HasBounds) return;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = _state.Left; Top = _state.Top; Width = _state.Width; Height = _state.Height;
            if (_state.TopPaneHeight > 0)
                TopPaneRow.Height = new GridLength(_state.TopPaneHeight, GridUnitType.Pixel);
        }

        private void RestoreSelection()
        {
            var row = _rows.FirstOrDefault(r => r.IdValue == _state.SelectedId);
            if (row == null) { UpdateCompare(); return; }
            RowGrid.SelectedItem = row;
            RowGrid.ScrollIntoView(row);
            UpdateCompare();
        }

        /// <summary>削除時の Revit の警告を画面に出さずに記録する（結果ダイアログでまとめて知らせる）。</summary>
        private class WarningCollector : IFailuresPreprocessor
        {
            public List<string> Messages { get; } = new List<string>();

            public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
            {
                try
                {
                    foreach (var f in a.GetFailureMessages())
                    {
                        var text = f?.GetDescriptionText();
                        if (!string.IsNullOrWhiteSpace(text) && !Messages.Contains(text))
                            Messages.Add(text);
                    }
                }
                catch { }
                a.DeleteAllWarnings();
                return FailureProcessingResult.Continue;
            }
        }
    }
}

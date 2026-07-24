using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
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
        private bool _ready;   // 初期化完了フラグ（InitializeComponent 中のイベント発火を無視）

        // ===== Excel風 列フィルター/並べ替え =====
        private readonly Dictionary<string, HashSet<string>> _columnFilters = new Dictionary<string, HashSet<string>>();
        private readonly Dictionary<string, Func<ParamRow, string>> _colAccessors = new Dictionary<string, Func<ParamRow, string>>();
        private readonly Dictionary<string, string> _colSortPaths = new Dictionary<string, string>();
        private readonly Dictionary<string, DataGridTextColumn> _colByKey = new Dictionary<string, DataGridTextColumn>();
        private readonly Dictionary<string, Button> _filterButtons = new Dictionary<string, Button>();
        private Popup _activePopup;

        public ParameterCleanupDialog(Document doc)
        {
            _doc = doc;
            InitializeComponent();
            ApplyLocalization();
            LoadRows();
            _ready = true;
            Loaded += OnDialogLoaded;
        }

        // ダイアログ表示直後に値の有無を自動確認する
        private async void OnDialogLoaded(object sender, RoutedEventArgs e)
        {
            Loaded -= OnDialogLoaded;   // 初回のみ
            await RunValueCheckAsync();
        }

        private void ApplyLocalization()
        {
            Title = Loc.S("ParamCleanup.Title");
            txtDescription.Text = Loc.S("ParamCleanup.Description");
            lblSearch.Text = Loc.S("ParamCleanup.Search");
            chkDuplicateOnly.Content = Loc.S("ParamCleanup.DuplicateOnly");
            chkEmptyOnly.Content = Loc.S("ParamCleanup.EmptyOnly");
            btnSelectEmpty.Content = Loc.S("ParamCleanup.SelectEmpty");

            lblKindFilter.Text = Loc.S("ParamCleanup.Filter.Kind");
            rbKindAll.Content = Loc.S("ParamCleanup.Filter.All");
            rbKindProject.Content = Loc.S("ParamCleanup.Kind.Project");
            rbKindShared.Content = Loc.S("ParamCleanup.Kind.Shared");
            rbKindGlobal.Content = Loc.S("ParamCleanup.Kind.Global");

            // 列ヘッダー：選択列は文字列、その他は Excel 風の並べ替え/フィルターボタン付き
            ParamGrid.Columns[0].Header = Loc.S("ParamCleanup.Col.Select");
            SetupFilterableColumns();

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

            // 種別ラジオフィルタ
            if (rbKindProject?.IsChecked == true && r.Kind != ParamKind.Project) return false;
            if (rbKindShared?.IsChecked == true && r.Kind != ParamKind.Shared) return false;
            if (rbKindGlobal?.IsChecked == true && r.Kind != ParamKind.Global) return false;

            // Excel風 列フィルタ（列ごとの許可値セット）
            foreach (var kv in _columnFilters)
            {
                if (_colAccessors.TryGetValue(kv.Key, out var acc))
                {
                    var val = acc(r) ?? "";
                    if (!kv.Value.Contains(val)) return false;
                }
            }

            return true;
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            if (!_ready) return;   // 初期化中のイベントは無視
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

        // ===================== Excel風 列メニュー（並べ替え/フィルター） =====================

        /// <summary>フィルター対象列に、並べ替え/フィルターメニュー付きヘッダーを設定する。</summary>
        private void SetupFilterableColumns()
        {
            RegisterColumn("Name", colName, "ParamCleanup.Col.Name", "Name", r => r.Name ?? "");
            RegisterColumn("Kind", colKind, "ParamCleanup.Col.Kind", "KindText", r => r.KindText ?? "");
            RegisterColumn("Scope", colScope, "ParamCleanup.Col.Scope", "ScopeText", r => r.ScopeText ?? "");
            RegisterColumn("Categories", colCategories, "ParamCleanup.Col.Categories", "CategoriesText", r => r.CategoriesText ?? "");
            RegisterColumn("SchedRef", colSchedRef, "ParamCleanup.Col.ScheduleRef", "ScheduleRefText", r => r.ScheduleRefText ?? "");
            RegisterColumn("Value", colState, "ParamCleanup.Col.Value", "StateText", r => r.StateText ?? "");
        }

        private void RegisterColumn(string key, DataGridTextColumn col, string titleKey,
                                    string sortPath, Func<ParamRow, string> accessor)
        {
            _colByKey[key] = col;
            _colSortPaths[key] = sortPath;
            _colAccessors[key] = accessor;
            col.CanUserSort = false;   // 既定のヘッダークリックソートは使わず、独自メニューに統一
            col.Header = BuildFilterHeader(key, Loc.S(titleKey));
        }

        private FrameworkElement BuildFilterHeader(string key, string title)
        {
            var dock = new DockPanel { LastChildFill = true, HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch };

            var btn = new Button
            {
                Content = "▾",
                Width = 20,
                Padding = new Thickness(0),
                Margin = new Thickness(4, 0, 0, 0),
                FontSize = 10,
                Focusable = false,
                Cursor = Cursors.Hand,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground = Brushes.Gray,
                Tag = key,
                ToolTip = Loc.S("ParamCleanup.FilterHint"),
                VerticalAlignment = System.Windows.VerticalAlignment.Center
            };
            btn.Click += ColumnMenuButton_Click;
            DockPanel.SetDock(btn, Dock.Right);
            dock.Children.Add(btn);
            _filterButtons[key] = btn;

            var tb = new TextBlock
            {
                Text = title,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                FontWeight = FontWeights.SemiBold
            };
            dock.Children.Add(tb);
            return dock;
        }

        private void ColumnMenuButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string key)
                ShowColumnMenu(key, btn);
        }

        private void ShowColumnMenu(string key, UIElement anchor)
        {
            CloseActivePopup();
            if (!_colAccessors.TryGetValue(key, out var accessor)) return;

            // 候補値（現在の全行から。空白は (空白) として表示）
            var distinct = _rows.Select(r => accessor(r) ?? "").Distinct().ToList();
            distinct.Sort(StringComparer.CurrentCultureIgnoreCase);
            HashSet<string> allowed = _columnFilters.TryGetValue(key, out var f) ? f : null;

            var root = new StackPanel { Margin = new Thickness(6) };

            // --- 並べ替え ---
            var btnAsc = MenuButton(Loc.S("ParamCleanup.Filter.SortAsc"));
            btnAsc.Click += (s, ev) => { SortByColumn(key, ListSortDirection.Ascending); CloseActivePopup(); };
            var btnDesc = MenuButton(Loc.S("ParamCleanup.Filter.SortDesc"));
            btnDesc.Click += (s, ev) => { SortByColumn(key, ListSortDirection.Descending); CloseActivePopup(); };
            root.Children.Add(btnAsc);
            root.Children.Add(btnDesc);
            root.Children.Add(new Separator { Margin = new Thickness(0, 4, 0, 4) });

            // --- 検索 ---
            var search = new TextBox { Height = 24, Margin = new Thickness(0, 0, 0, 4) };
            root.Children.Add(search);

            // --- (すべて選択) ---
            var selectAll = new CheckBox
            {
                Content = Loc.S("ParamCleanup.Filter.SelectAll"),
                Margin = new Thickness(0, 0, 0, 4),
                FontWeight = FontWeights.SemiBold
            };
            root.Children.Add(selectAll);

            // --- 値リスト ---
            var listPanel = new StackPanel();
            var checks = new List<CheckBox>();
            foreach (var val in distinct)
            {
                var cb = new CheckBox
                {
                    Content = string.IsNullOrEmpty(val) ? Loc.S("ParamCleanup.Filter.Blank") : val,
                    Tag = val,
                    IsChecked = allowed == null || allowed.Contains(val),
                    Margin = new Thickness(0, 1, 0, 1)
                };
                checks.Add(cb);
                listPanel.Children.Add(cb);
            }

            selectAll.IsChecked = checks.All(c => c.IsChecked == true) ? true
                                : checks.Any(c => c.IsChecked == true) ? (bool?)null : false;
            selectAll.Click += (s, ev) =>
            {
                bool on = selectAll.IsChecked == true;
                foreach (var c in checks)
                    if (c.Visibility == System.Windows.Visibility.Visible) c.IsChecked = on;
            };

            search.TextChanged += (s, ev) =>
            {
                string q = search.Text.Trim();
                foreach (var c in checks)
                {
                    string disp = c.Content?.ToString() ?? "";
                    c.Visibility = (q.Length == 0 || disp.IndexOf(q, StringComparison.CurrentCultureIgnoreCase) >= 0)
                        ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                }
            };

            root.Children.Add(new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                MaxHeight = 240,
                Content = listPanel
            });

            // --- OK / クリア / キャンセル ---
            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                Margin = new Thickness(0, 6, 0, 0)
            };
            var btnClear = MenuButton(Loc.S("ParamCleanup.Filter.Clear"));
            btnClear.Width = 64; btnClear.Margin = new Thickness(0, 0, 4, 0);
            btnClear.Click += (s, ev) =>
            {
                _columnFilters.Remove(key);
                UpdateFilterIndicator(key);
                _view?.Refresh(); UpdateCount();
                CloseActivePopup();
            };
            var btnOk = MenuButton(Loc.S("Common.OK"));
            btnOk.Width = 64; btnOk.Margin = new Thickness(0, 0, 4, 0);
            btnOk.Click += (s, ev) =>
            {
                var sel = new HashSet<string>(checks.Where(c => c.IsChecked == true).Select(c => (string)c.Tag));
                if (sel.Count >= distinct.Count) _columnFilters.Remove(key);  // 全選択＝フィルタ無し
                else _columnFilters[key] = sel;
                UpdateFilterIndicator(key);
                _view?.Refresh(); UpdateCount();
                CloseActivePopup();
            };
            var btnCancel = MenuButton(Loc.S("Common.Cancel"));
            btnCancel.Width = 64;
            btnCancel.Click += (s, ev) => CloseActivePopup();
            buttons.Children.Add(btnClear);
            buttons.Children.Add(btnOk);
            buttons.Children.Add(btnCancel);
            root.Children.Add(buttons);

            var border = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xCC, 0xCC, 0xCC)),
                BorderThickness = new Thickness(1),
                Child = root,
                MinWidth = 230
            };

            _activePopup = new Popup
            {
                PlacementTarget = anchor,
                Placement = PlacementMode.Bottom,
                StaysOpen = false,
                AllowsTransparency = true,
                Child = border,
                IsOpen = true
            };
        }

        private static Button MenuButton(string text)
        {
            return new Button
            {
                Content = text,
                Height = 26,
                Margin = new Thickness(0, 1, 0, 1),
                Padding = new Thickness(8, 0, 8, 0),
                HorizontalContentAlignment = System.Windows.HorizontalAlignment.Left,
                Background = Brushes.White,
                Cursor = Cursors.Hand
            };
        }

        private void SortByColumn(string key, ListSortDirection dir)
        {
            if (_view == null || !_colSortPaths.TryGetValue(key, out var path)) return;
            _view.SortDescriptions.Clear();
            _view.SortDescriptions.Add(new SortDescription(path, dir));
            _view.Refresh();
        }

        private void UpdateFilterIndicator(string key)
        {
            if (!_filterButtons.TryGetValue(key, out var btn)) return;
            bool active = _columnFilters.ContainsKey(key);
            btn.Foreground = active ? Brushes.RoyalBlue : Brushes.Gray;
            btn.Content = active ? "▼" : "▾";
            btn.FontWeight = active ? FontWeights.Bold : FontWeights.Normal;
        }

        private void CloseActivePopup()
        {
            if (_activePopup != null)
            {
                _activePopup.IsOpen = false;
                _activePopup = null;
            }
        }

        // ===== 値の有無を確認（スキャン）=====
        // ダイアログ表示時に自動実行され、ボタンでの再確認・中止も同じ処理を使う。
        private void Check_Click(object sender, RoutedEventArgs e)
        {
            if (_scanning)
            {
                _cts?.Cancel();   // 実行中はキャンセルボタンとして機能
                return;
            }
            _ = RunValueCheckAsync();
        }

        /// <summary>値の有無スキャンを実行（自動確認・手動再確認の共通処理）。</summary>
        private async Task RunValueCheckAsync()
        {
            if (_scanning) return;

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
            Progress.Visibility = System.Windows.Visibility.Visible;
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
                Progress.Visibility = System.Windows.Visibility.Collapsed;
                if (cancelled)
                {
                    txtStatus.Text = Loc.S("ParamCleanup.Status.Cancelled");
                }
                else
                {
                    int hasVal = _rows.Count(r => r.State == ValueState.HasValue);
                    int empty = _rows.Count(r => r.State == ValueState.Empty);
                    txtStatus.Text = string.Format(Loc.S("ParamCleanup.Status.DoneSummary"), hasVal, empty);
                }

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

            LoadRows();          // 一覧を再構築
            txtStatus.Text = "";
            _ = RunValueCheckAsync();   // 削除後に自動で再確認
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            CloseActivePopup();
            if (_scanning)
                _cts?.Cancel();
            base.OnClosing(e);
        }
    }
}

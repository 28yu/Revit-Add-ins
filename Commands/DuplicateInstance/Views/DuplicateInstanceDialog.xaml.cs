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

        // 一覧のパラメータ列（全要素のインスタンスパラメータ名の和集合・名前順）
        private List<string> _paramKeys = new List<string>();
        private int _fixedColumnCount;
        // 列ごとの「並べ替えに使う値の取り出し方」と、開き直したときに同じ列を探すための名前
        private readonly Dictionary<DataGridColumn, Func<DuplicateRow, string>> _sortGetters
            = new Dictionary<DataGridColumn, Func<DuplicateRow, string>>();
        private readonly Dictionary<DataGridColumn, string> _columnKeys = new Dictionary<DataGridColumn, string>();
        // 列の名前 → 値の取り出し方 / 見出しの ▼ ボタン（Excel のような絞り込み用）
        private readonly Dictionary<string, Func<DuplicateRow, string>> _gettersByKey
            = new Dictionary<string, Func<DuplicateRow, string>>();
        private readonly Dictionary<string, System.Windows.Controls.Button> _filterButtons
            = new Dictionary<string, System.Windows.Controls.Button>();
        // 開いている絞り込みポップアップの対象列と、その値の一覧
        private string _popupKey;
        private List<FilterItem> _popupItems = new List<FilterItem>();
        private ICollectionView _popupView;
        private bool _ready;

        private string AllCategories => Loc.S("DupInst.Filter.AllCategories");

        public DuplicateInstanceDialog(Document doc, DuplicateDialogState state)
        {
            _doc = doc;
            _state = state ?? new DuplicateDialogState();
            InitializeComponent();
            ApplyLocalization();
            RegisterFixedColumns();
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
            Title = string.IsNullOrEmpty(_state.ScopeLabel)
                ? Loc.S("DupInst.Title")
                : Loc.S("DupInst.Title") + " - " + _state.ScopeLabel;
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
            btnClearFilters.Content = Loc.S("DupInst.Btn.ClearFilters");
            btnClearFilters.ToolTip = Loc.S("DupInst.Btn.ClearFilters.Tip");
            btnPopupSortAsc.Content = Loc.S("DupInst.Popup.SortAsc");
            btnPopupSortDesc.Content = Loc.S("DupInst.Popup.SortDesc");
            btnPopupClear.Content = Loc.S("DupInst.Popup.Clear");
            chkPopupAll.Content = Loc.S("DupInst.Popup.SelectAll");
            btnPopupOk.Content = Loc.S("Common.OK");
            btnPopupCancel.Content = Loc.S("Common.Cancel");
            btnReload.ToolTip = Loc.S("Common.Reload.Tip");
            btnDelete.Content = Loc.S("DupInst.Btn.Delete");
            btnClose.Content = Loc.S("DupInst.Btn.Close");
        }

        // ===== 読み込み・絞り込み =====

        private void LoadRows()
        {
            using (this.BlockRevitInput())
            {
                _rows = _scanner.Scan(_doc, _state.ScopeIds);
            }

            foreach (var r in _rows)
                r.IsSelected = _state.CheckedIds.Contains(r.IdValue);

            // 全要素のパラメータ名を集めて列にし、各行の値をその順番に並べる
            _paramKeys = _rows.SelectMany(r => r.Parameters.Keys).Distinct()
                              .OrderBy(k => k, StringComparer.CurrentCulture).ToList();
            foreach (var r in _rows)
                r.Values = _paramKeys.Select(k => r.Parameters.TryGetValue(k, out var v) ? v ?? "" : "").ToArray();
            BuildParamColumns();

            _view = CollectionViewSource.GetDefaultView(_rows);
            _view.Filter = RowFilter;
            RowGrid.ItemsSource = _view;
            ApplySavedSort();

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

        private bool RowFilter(object o) => o is DuplicateRow r && PassesFilters(r, null);

        /// <summary>絞り込み条件をすべて満たすか。skipKey の列の絞り込みだけは無視する（その列の値一覧を作るため）。</summary>
        private bool PassesFilters(DuplicateRow r, string skipKey)
        {
            foreach (var f in _state.ColumnFilters)
            {
                if (f.Key == skipKey) continue;
                if (!_gettersByKey.TryGetValue(f.Key, out var get)) continue;   // 今は無い列
                if (!f.Value.Contains(get(r) ?? "")) return false;
            }

            string cat = cmbCategory?.SelectedItem as string;
            if (!string.IsNullOrEmpty(cat) && cat != AllCategories && r.Category != cat) return false;

            if (chkCheckedOnly?.IsChecked == true && !r.IsSelected) return false;

            string q = SearchBox?.Text?.Trim();
            if (!string.IsNullOrEmpty(q))
            {
                bool hit = new[] { r.IdValue.ToString(), r.Category, r.Family, r.TypeName, r.Host }
                    .Concat(r.Values)
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

        // ===== Excel のような値の絞り込み =====

        /// <summary>列見出し: 名前と、右端の ▼（絞り込み）ボタン。</summary>
        private object MakeHeader(string text, string key)
        {
            var panel = new DockPanel { LastChildFill = true };
            var btn = new System.Windows.Controls.Button
            {
                Content = "▼",
                FontSize = 8,
                Width = 16,
                Height = 16,
                Padding = new Thickness(0),
                Margin = new Thickness(4, 0, 0, 0),
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand,
                Focusable = false,
                ToolTip = Loc.S("DupInst.Filter.Tip")
            };
            // ボタンが押下を受け取るので、見出しの並べ替えは起きない
            btn.Click += (s, e) => { e.Handled = true; OpenFilterPopup(key, btn); };
            DockPanel.SetDock(btn, Dock.Right);
            panel.Children.Add(btn);
            panel.Children.Add(new TextBlock
            {
                Text = text,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                ToolTip = text
            });
            _filterButtons[key] = btn;
            return panel;
        }

        /// <summary>絞り込み中の列の ▼ を青くする。</summary>
        private void UpdateFilterButtons()
        {
            var on = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0x00, 0x66, 0xCC));
            var offBorder = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0xBB, 0xBB, 0xBB));
            foreach (var kv in _filterButtons)
            {
                bool active = _state.ColumnFilters.ContainsKey(kv.Key);
                kv.Value.Background = active ? on : System.Windows.Media.Brushes.White;
                kv.Value.Foreground = active ? System.Windows.Media.Brushes.White : System.Windows.Media.Brushes.DimGray;
                kv.Value.BorderBrush = active ? on : offBorder;
            }
        }

        private void OpenFilterPopup(string key, UIElement target)
        {
            if (!_gettersByKey.TryGetValue(key, out var get)) return;
            _popupKey = key;

            // 他の絞り込みで残っている行から、この列の値と件数を集める（Excel と同じ）
            var counts = new Dictionary<string, int>();
            foreach (var r in _rows.Where(r => PassesFilters(r, key)))
            {
                string v = get(r) ?? "";
                counts[v] = counts.TryGetValue(v, out var n) ? n + 1 : 1;
            }
            _state.ColumnFilters.TryGetValue(key, out var allowed);

            var values = counts.Keys.ToList();
            values.Sort((a, b) =>
            {
                bool ea = a.Length == 0, eb = b.Length == 0;
                if (ea || eb) return ea == eb ? 0 : (ea ? 1 : -1);   // 空白は最後
                return RowComparer.CompareValues(a, b);
            });
            _popupItems = values.Select(v => new FilterItem
            {
                Value = v,
                Display = (v.Length == 0 ? Loc.S("DupInst.Popup.Blank") : v) + "  (" + counts[v] + ")",
                IsChecked = allowed == null || allowed.Contains(v)
            }).ToList();

            _popupView = CollectionViewSource.GetDefaultView(_popupItems);
            _popupView.Filter = o =>
            {
                string q = PopupSearch.Text?.Trim();
                return string.IsNullOrEmpty(q) ||
                       ((FilterItem)o).Display.IndexOf(q, StringComparison.CurrentCultureIgnoreCase) >= 0;
            };
            PopupSearch.Text = "";
            PopupList.ItemsSource = _popupView;
            UpdatePopupAll();

            FilterPopup.PlacementTarget = target;
            FilterPopup.IsOpen = true;
            PopupSearch.Focus();
        }

        private IEnumerable<FilterItem> VisiblePopupItems => _popupView?.Cast<FilterItem>() ?? Enumerable.Empty<FilterItem>();

        private void UpdatePopupAll()
        {
            var vis = VisiblePopupItems.ToList();
            if (vis.Count == 0) { chkPopupAll.IsChecked = false; return; }
            int on = vis.Count(i => i.IsChecked);
            chkPopupAll.IsChecked = on == vis.Count ? true : (on == 0 ? (bool?)false : null);
        }

        private void PopupSearch_Changed(object sender, TextChangedEventArgs e)
        {
            _popupView?.Refresh();
            UpdatePopupAll();
        }

        private void PopupAll_Click(object sender, RoutedEventArgs e)
        {
            bool on = chkPopupAll.IsChecked != false;
            foreach (var i in VisiblePopupItems) i.IsChecked = on;
            UpdatePopupAll();
        }

        private void PopupItem_Click(object sender, RoutedEventArgs e) => UpdatePopupAll();

        private void PopupOk_Click(object sender, RoutedEventArgs e)
        {
            if (_popupKey == null) return;
            bool searching = !string.IsNullOrWhiteSpace(PopupSearch.Text);
            // 検索中は、見えている項目のうちチェックしたものだけを表示する（Excel と同じ）
            var considered = searching ? VisiblePopupItems.ToList() : _popupItems;
            var allowed = new HashSet<string>(considered.Where(i => i.IsChecked).Select(i => i.Value));
            if (allowed.Count == 0) return;   // 何も選ばないと全部消えてしまうので、確定させない

            if (!searching && allowed.Count == _popupItems.Count)
                _state.ColumnFilters.Remove(_popupKey);
            else
                _state.ColumnFilters[_popupKey] = allowed;

            FilterPopup.IsOpen = false;
            UpdateFilterButtons();
            RefreshView();
        }

        private void PopupCancel_Click(object sender, RoutedEventArgs e) => FilterPopup.IsOpen = false;

        private void PopupClear_Click(object sender, RoutedEventArgs e)
        {
            if (_popupKey != null) _state.ColumnFilters.Remove(_popupKey);
            FilterPopup.IsOpen = false;
            UpdateFilterButtons();
            RefreshView();
        }

        private void PopupSortAsc_Click(object sender, RoutedEventArgs e) => SortFromPopup(ListSortDirection.Ascending);
        private void PopupSortDesc_Click(object sender, RoutedEventArgs e) => SortFromPopup(ListSortDirection.Descending);

        private void SortFromPopup(ListSortDirection dir)
        {
            var col = _columnKeys.FirstOrDefault(kv => kv.Value == _popupKey).Key;
            FilterPopup.IsOpen = false;
            if (col != null) ApplySort(col, dir);
        }

        private void ClearFilters_Click(object sender, RoutedEventArgs e)
        {
            _ready = false;
            _state.ColumnFilters.Clear();
            SearchBox.Text = "";
            cmbCategory.SelectedItem = AllCategories;
            chkCheckedOnly.IsChecked = false;
            _ready = true;
            UpdateFilterButtons();
            RefreshView();
        }

        // ===== パラメータ列と並べ替え =====

        private void RegisterFixedColumns()
        {
            _fixedColumnCount = RowGrid.Columns.Count;
            void Reg(DataGridColumn c, string key, Func<DuplicateRow, string> get)
            {
                string k = "fixed:" + key;
                _sortGetters[c] = get;
                _columnKeys[c] = k;
                if (c == colCheck) return;
                _gettersByKey[k] = get;
                c.Header = MakeHeader(c.Header as string ?? "", k);
            }
            Reg(colCheck, "Check", r => r.IsSelected ? "1" : "0");
            Reg(colGroup, "Group", r => r.GroupNo.ToString());
            Reg(colId, "Id", r => r.IdValue.ToString());
            Reg(colCategory, "Category", r => r.Category);
            Reg(colFamily, "Family", r => r.Family);
            Reg(colType, "Type", r => r.TypeName);
            Reg(colFilled, "Filled", r => r.FilledCount.ToString());
            Reg(colDiff, "Diff", r => r.DiffCount.ToString());
            UpdateFilterButtons();
        }

        /// <summary>固定列の後ろに、パラメータ1つにつき1列を作り直す。</summary>
        private void BuildParamColumns()
        {
            while (RowGrid.Columns.Count > _fixedColumnCount)
            {
                var c = RowGrid.Columns[RowGrid.Columns.Count - 1];
                if (_columnKeys.TryGetValue(c, out var oldKey))
                {
                    _gettersByKey.Remove(oldKey);
                    _filterButtons.Remove(oldKey);
                }
                _sortGetters.Remove(c);
                _columnKeys.Remove(c);
                RowGrid.Columns.RemoveAt(RowGrid.Columns.Count - 1);
            }

            var style = (Style)FindResource("CellText");
            for (int i = 0; i < _paramKeys.Count; i++)
            {
                int idx = i;   // ラムダで使うため、ループ変数を写しておく
                string key = "param:" + _paramKeys[i];
                var col = new DataGridTextColumn
                {
                    Header = MakeHeader(_paramKeys[i], key),
                    Binding = new System.Windows.Data.Binding($"Values[{i}]") { Mode = BindingMode.OneWay },
                    IsReadOnly = true,
                    Width = DataGridLength.Auto,
                    MinWidth = 70,
                    MaxWidth = 260,
                    ElementStyle = style,
                    SortMemberPath = "IdValue"   // 実際の並べ替えは RowGrid_Sorting で行う
                };
                RowGrid.Columns.Add(col);
                Func<DuplicateRow, string> get = r => idx < r.Values.Length ? r.Values[idx] : "";
                _sortGetters[col] = get;
                _gettersByKey[key] = get;
                _columnKeys[col] = key;
            }
            UpdateFilterButtons();
        }

        /// <summary>列見出しのクリック。数値は数として、空欄は常に最後に並べる独自の並べ替えにする。</summary>
        private void RowGrid_Sorting(object sender, DataGridSortingEventArgs e)
        {
            e.Handled = true;
            var dir = e.Column.SortDirection == ListSortDirection.Ascending
                ? ListSortDirection.Descending
                : ListSortDirection.Ascending;
            ApplySort(e.Column, dir);
        }

        private void ApplySort(DataGridColumn col, ListSortDirection dir)
        {
            if (!(_view is ListCollectionView lcv) || !_sortGetters.TryGetValue(col, out var getter)) return;
            foreach (var c in RowGrid.Columns)
                if (c != col) c.SortDirection = null;
            col.SortDirection = dir;
            bool desc = dir == ListSortDirection.Descending;
            lcv.CustomSort = new RowComparer(getter, desc);
            _state.SortKey = _columnKeys.TryGetValue(col, out var k) ? k : null;
            _state.SortDescending = desc;
        }

        /// <summary>前回と同じ列で並べ替える（再読み込み・「画面で確認」から戻ったとき）。</summary>
        private void ApplySavedSort()
        {
            if (string.IsNullOrEmpty(_state.SortKey)) return;
            var col = _columnKeys.FirstOrDefault(kv => kv.Value == _state.SortKey).Key;
            if (col == null) return;
            ApplySort(col, _state.SortDescending ? ListSortDirection.Descending : ListSortDirection.Ascending);
        }

        /// <summary>
        /// 一覧の並べ替え。両方が数で始まる値（例: 1200 mm）は数の大小で比べ、
        /// 空欄は昇順・降順どちらでも最後に置く。同じ値なら組 → 要素ID の順。
        /// </summary>
        private class RowComparer : System.Collections.IComparer
        {
            private static readonly System.Text.RegularExpressions.Regex NumRe =
                new System.Text.RegularExpressions.Regex(@"^\s*[-+]?(\d[\d,]*)?(\.\d+)?");

            private readonly Func<DuplicateRow, string> _get;
            private readonly bool _desc;

            public RowComparer(Func<DuplicateRow, string> get, bool desc) { _get = get; _desc = desc; }

            public int Compare(object x, object y)
            {
                var a = x as DuplicateRow;
                var b = y as DuplicateRow;
                if (a == null || b == null) return 0;
                string va = _get(a) ?? "", vb = _get(b) ?? "";

                bool ea = string.IsNullOrWhiteSpace(va), eb = string.IsNullOrWhiteSpace(vb);
                int c;
                if (ea || eb)
                    c = ea && eb ? 0 : (ea ? 1 : -1);   // 空欄は向きに関係なく最後
                else
                {
                    c = CompareValues(va, vb);
                    if (_desc) c = -c;
                }
                if (c != 0) return c;
                c = a.GroupNo.CompareTo(b.GroupNo);
                return c != 0 ? c : a.IdValue.CompareTo(b.IdValue);
            }

            internal static int CompareValues(string a, string b)
            {
                bool na = TryNumber(a, out double da), nb = TryNumber(b, out double db);
                if (na && nb)
                {
                    int c = da.CompareTo(db);
                    return c != 0 ? c : string.Compare(a, b, StringComparison.CurrentCulture);
                }
                if (na != nb) return na ? -1 : 1;   // 数で始まる値を先に
                return string.Compare(a, b, StringComparison.CurrentCulture);
            }

            private static bool TryNumber(string s, out double v)
            {
                v = 0;
                var m = NumRe.Match(s);
                if (!m.Success || (m.Groups[1].Length == 0 && m.Groups[2].Length == 0)) return false;
                return double.TryParse(m.Value.Replace(",", "").Trim(),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out v);
            }
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
                txtCompareTitle.Text = _rows.Count == 0
                    ? (_state.ScopeIds == null ? Loc.S("DupInst.NoWarnings") : Loc.S("DupInst.NoWarningsInScope"))
                    : Loc.S("DupInst.Compare.Hint");
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
                Binding = new System.Windows.Data.Binding(nameof(CompareRow.Name)),
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
                    Binding = new System.Windows.Data.Binding($"Values[{i}]"),
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

            DeleteResult result;
            var progress = new DeleteProgressWindow { Owner = this };
            IsEnabled = false;
            try
            {
                progress.Show();
                result = new DuplicateDeleter().Delete(_doc, selected, progress);
            }
            catch (Exception ex)
            {
                result = null;
                progress.Finish();
                IsEnabled = true;
                TaskDialog.Show(Loc.S("Common.Error"), string.Format(Loc.S("DupInst.Result.Error"), ex.Message));
                this.BringToFrontDeferred();
                return;
            }
            finally
            {
                progress.Finish();
                IsEnabled = true;
            }

            if (result.Cancelled)
            {
                TaskDialog.Show(Loc.S("DupInst.Result.Title"), Loc.S("DupInst.Result.Cancelled"));
                this.BringToFrontDeferred();
                return;
            }

            ShowDeleteResult(result);

            // 削除できた要素のチェックだけ外す。削除できなかった要素はチェックを残し、すぐ見つけられるようにする
            foreach (var r in _rows)
                r.IsSelected = r.IsSelected && !result.DeletedIds.Contains(r.IdValue);
            CaptureChecks();
            LoadRows();
            RefreshView();
            this.BringToFrontDeferred();
        }

        /// <summary>削除の結果を表示する。削除できなかった要素は原因ごとにまとめ、要素IDを添える。</summary>
        private void ShowDeleteResult(DeleteResult result)
        {
            int ok = result.DeletedIds.Count, ng = result.Failed.Count;
            var td = new TaskDialog(Loc.S("DupInst.Result.Title"))
            {
                MainInstruction = string.Format(Loc.S("DupInst.Result.Msg"), ok, ng)
            };

            var sb = new System.Text.StringBuilder();
            if (ng > 0)
            {
                sb.AppendLine(Loc.S("DupInst.Result.FailedHeader"));
                const int maxGroups = 8, maxIds = 10;
                var groups = result.Failed.GroupBy(f => f.Value).OrderByDescending(g => g.Count()).ToList();
                foreach (var g in groups.Take(maxGroups))
                {
                    sb.AppendLine();
                    sb.AppendLine(string.Format(Loc.S("DupInst.Result.ReasonLine"), g.Key, g.Count()));
                    var ids = g.Select(f => f.Key.IdValue.ToString()).ToList();
                    string idText = string.Join(", ", ids.Take(maxIds));
                    if (ids.Count > maxIds) idText += string.Format(Loc.S("DupInst.Result.More"), ids.Count - maxIds);
                    sb.AppendLine("    ID: " + idText);
                }
                if (groups.Count > maxGroups)
                    sb.AppendLine().AppendLine(string.Format(Loc.S("DupInst.Result.MoreReasons"), groups.Count - maxGroups));
                sb.AppendLine().Append(Loc.S("DupInst.Result.FailedHint"));
            }
            td.MainContent = sb.ToString().Trim();

            if (result.Warnings.Count > 0)
            {
                string w = Loc.S("Common.RevitWarnings") + "\n"
                         + string.Join("\n", result.Warnings.Take(15).Select(x => "・" + x));
                if (result.Warnings.Count > 15) w += "\n…";
                td.ExpandedContent = w;
            }
            td.Show();
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
    }
}

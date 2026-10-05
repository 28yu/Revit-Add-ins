using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Microsoft.Win32;
using Tools28.Commands.ExcelExportImport.Services;
using Tools28.Localization;

namespace Tools28.Commands.ExcelExportImport.Views
{
    /// <summary>
    /// EXCELインポートダイアログ
    /// </summary>
    public partial class ImportDialog : Window
    {
        private readonly Document _doc;
        private string _selectedFilePath;
        private List<ImportPreviewRow> _previewRows;

        // 変更プレビューに表示する行（変更あり かつ 書き込み可能）と、その表示用ビュー（並べ替え・絞り込み）
        private List<ImportPreviewRow> _changedRows = new List<ImportPreviewRow>();
        private System.Windows.Data.ListCollectionView _view;

        // ===== Excel風 列フィルター/並べ替え（パラメータ整理と同じ操作感）=====
        // 列キー -> 表示を許可する値の集合（キーが無い列は絞り込みなし）
        private readonly Dictionary<string, HashSet<string>> _columnFilters = new Dictionary<string, HashSet<string>>();
        private readonly Dictionary<string, Func<ImportPreviewRow, string>> _colAccessors = new Dictionary<string, Func<ImportPreviewRow, string>>();
        private readonly Dictionary<string, string> _colSortPaths = new Dictionary<string, string>();
        private readonly Dictionary<string, Button> _filterButtons = new Dictionary<string, Button>();
        private Popup _activePopup;

        // 一括選択中はチェック変更ごとのサマリー更新を止める（行数が多いと重いため）
        private bool _suppressSummary;

        // 読み込み結果のサマリー（再計算用）
        private int _totalCount;
        private int _readOnlyChangeCount;
        private int _missingChangeCount;

        /// <summary>インポートが実行されたかどうか</summary>
        public bool ImportExecuted { get; private set; }

        /// <summary>インポート結果</summary>
        public ImportResult ImportResultData { get; private set; }

        public ImportDialog(Document doc)
        {
            InitializeComponent();
            ApplyLocalization();
            _doc = doc;

            // 起動時に開いているExcelファイルを自動検出
            AutoDetectOpenFiles();
        }

        private void ApplyLocalization()
        {
            this.Title = Loc.S("Import.Title");
            grpExcelFile.Header = Loc.S("Import.ExcelFile");
            btnOpenFiles.Content = Loc.S("Import.OpenFile");
            btnBrowse.Content = Loc.S("Import.Browse");
            grpPreview.Header = Loc.S("Import.Preview");
            colSelect.Header = Loc.S("Import.ColSelect");
            RegisterColumn("ElementId", colElementId, Loc.S("Import.ColElementId"), "ElementId", r => r.ElementId.ToString());
            RegisterColumn("Category", colCategory, Loc.S("Import.ColCategory"), "CategoryName", r => r.CategoryName ?? "");
            RegisterColumn("Parameter", colParameter, Loc.S("Import.ColParameter"), "ParameterName", r => r.ParameterName ?? "");
            RegisterColumn("Current", colCurrentValue, Loc.S("Import.ColCurrentValue"), "CurrentValue", r => r.CurrentValue ?? "");
            RegisterColumn("New", colNewValue, Loc.S("Import.ColNewValue"), "NewValue", r => r.NewValue ?? "");
            btnSelectAllVisible.Content = Loc.S("Import.SelectAllVisible");
            btnDeselectAllVisible.Content = Loc.S("Import.DeselectAllVisible");
            btnClearFilters.Content = Loc.S("Import.ClearFilters");
            txtFilterHint.Text = Loc.S("Import.FilterHint");
            ImportButton.Content = Loc.S("Import.Execute");
            btnCancel.Content = Loc.S("Common.Cancel");
        }

        private void AutoDetectOpenFiles()
        {
            try
            {
                var openFiles = ExcelProcessHelper.GetOpenExcelFiles();
                if (openFiles.Count == 1)
                {
                    // 起動時の自動検出なので、解決できなくてもメッセージは出さない
                    SelectOpenWorkbook(openFiles[0], showError: false);
                }
                else if (openFiles.Count > 1)
                {
                    ShowOpenFileSelection(openFiles);
                }
            }
            catch
            {
                // 自動検出の失敗は無視
            }
        }

        private void OpenFilesButton_Click(object sender, RoutedEventArgs e)
        {
            var openFiles = ExcelProcessHelper.GetOpenExcelFiles();
            if (openFiles.Count == 0)
            {
                MessageBox.Show(Loc.S("Import.NoOpenFile"),
                    Loc.S("Common.Confirm"), MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            else if (openFiles.Count == 1)
            {
                SelectOpenWorkbook(openFiles[0], showError: true);
            }
            else
            {
                ShowOpenFileSelection(openFiles);
            }
        }

        private void ShowOpenFileSelection(List<string> openFiles)
        {
            var selectWindow = new Window
            {
                Title = Loc.S("Import.SelectExcelFile"),
                Width = 500,
                Height = 300,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                Background = System.Windows.Media.Brushes.WhiteSmoke
            };

            var grid = new System.Windows.Controls.Grid { Margin = new Thickness(10) };
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Auto) });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Auto) });

            var label = new TextBlock
            {
                Text = "インポートするExcelファイルを選択してください:",
                FontSize = 12,
                Margin = new Thickness(0, 0, 0, 8)
            };
            System.Windows.Controls.Grid.SetRow(label, 0);
            grid.Children.Add(label);

            var listBox = new ListBox { FontSize = 11 };
            foreach (var file in openFiles)
            {
                listBox.Items.Add(new ListBoxItem
                {
                    Content = CloudExcelPathResolver.GetDisplayFileName(file),
                    Tag = file,
                    ToolTip = file
                });
            }
            if (listBox.Items.Count > 0)
                listBox.SelectedIndex = 0;
            System.Windows.Controls.Grid.SetRow(listBox, 1);
            grid.Children.Add(listBox);

            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 8, 0, 0)
            };

            string selectedPath = null;
            var okBtn = new Button
            {
                Content = "OK",
                Width = 80,
                Height = 28,
                Margin = new Thickness(4, 0, 0, 0),
                IsDefault = true
            };
            okBtn.Click += (s, args) =>
            {
                var selected = listBox.SelectedItem as ListBoxItem;
                if (selected != null)
                {
                    selectedPath = selected.Tag as string;
                    selectWindow.DialogResult = true;
                }
            };

            var cancelBtn = new Button
            {
                Content = Loc.S("Common.Cancel"),
                Width = 80,
                Height = 28,
                Margin = new Thickness(4, 0, 0, 0),
                IsCancel = true
            };

            btnPanel.Children.Add(okBtn);
            btnPanel.Children.Add(cancelBtn);
            System.Windows.Controls.Grid.SetRow(btnPanel, 2);
            grid.Children.Add(btnPanel);

            selectWindow.Content = grid;

            // ダブルクリックで選択
            listBox.MouseDoubleClick += (s, args) =>
            {
                var selected = listBox.SelectedItem as ListBoxItem;
                if (selected != null)
                {
                    selectedPath = selected.Tag as string;
                    selectWindow.DialogResult = true;
                }
            };

            selectWindow.Owner = this;
            if (selectWindow.ShowDialog() == true && selectedPath != null)
            {
                SelectOpenWorkbook(selectedPath, showError: true);
            }
        }

        /// <summary>
        /// Excel で開いているブックを読み込み対象にする。
        /// クラウド上のブック（URL）は、「参照」で選んだときと同じく同期フォルダ内の実ファイルに置き換える
        /// （見つからなければ開いている内容の複製を使う）。
        /// </summary>
        private void SelectOpenWorkbook(string fullName, bool showError)
        {
            DiagLog.Write($"[ExcelImport] 開いているブックを選択: {fullName}");

            if (!CloudExcelPathResolver.IsCloudPath(fullName))
            {
                LoadPreview(fullName, showError);
                return;
            }

            // クラウド上のブック: まず同期フォルダの実ファイル（「参照」と同じファイル）で読む。
            // 見つからない／読めない場合は、Excel で開いている内容の複製で読み直す。
            string path = CloudExcelPathResolver.Resolve(fullName);
            if (path != null && LoadPreview(path, showError: false))
                return;

            string copy = ExcelProcessHelper.SaveOpenWorkbookCopy(fullName);
            if (copy != null && !string.Equals(copy, path, StringComparison.OrdinalIgnoreCase))
            {
                DiagLog.Write($"[ExcelImport] 複製で読み直し: {copy}");
                if (LoadPreview(copy, showError))
                    return;
            }
            else if (path != null && showError)
            {
                // 複製もできない → 最初の読み込みエラーを表示する
                LoadPreview(path, showError: true);
                return;
            }

            if (copy == null && showError)
            {
                MessageBox.Show(string.Format(Loc.S("Import.CloudFileNotResolved"), fullName),
                    Loc.S("Common.Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excelファイル (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx"
            };

            if (dialog.ShowDialog(this) == true)
            {
                _selectedFilePath = dialog.FileName;
                FilePathTextBox.Text = _selectedFilePath;
                LoadPreview();
            }
        }

        private void LoadPreview()
        {
            LoadPreview(_selectedFilePath, showError: true);
        }

        /// <summary>
        /// 指定ファイルを読み込んでプレビューを表示する。成功した場合だけ選択ファイルとして確定する。
        /// </summary>
        /// <returns>読み込めた場合 true</returns>
        private bool LoadPreview(string path, bool showError)
        {
            try
            {
                // プレビューを生成
                _previewRows = ExcelImportService.GeneratePreview(_doc, path);
                _selectedFilePath = path;
                FilePathTextBox.Text = path;

                // 書き込み可能な変更のみ表示（読み取り専用パラメータは除外）
                foreach (var old in _changedRows)
                    old.PropertyChanged -= Row_PropertyChanged;

                _changedRows = _previewRows
                    .Where(r => r.HasChange && !r.IsReadOnly)
                    .OrderBy(r => r.CategoryName)
                    .ThenBy(r => r.ElementId)
                    .ToList();

                foreach (var row in _changedRows)
                    row.PropertyChanged += Row_PropertyChanged;

                // 別ファイルを読み直したときは、前のファイルの絞り込みを持ち越さない
                CloseActivePopup();
                _columnFilters.Clear();
                foreach (var key in _filterButtons.Keys.ToList())
                    UpdateFilterIndicator(key);

                _view = new System.Windows.Data.ListCollectionView(_changedRows) { Filter = RowFilter };
                PreviewDataGrid.ItemsSource = _view;

                _totalCount = _previewRows.Count;
                _readOnlyChangeCount = _previewRows.Count(r => r.HasChange && r.IsReadOnly && !r.ParamMissing);
                _missingChangeCount = _previewRows.Count(r => r.HasChange && r.ParamMissing);
                UpdateSummary();
                return true;
            }
            catch (Exception ex)
            {
                // 原因調査用に、読もうとした場所と例外の詳細を残す
                DiagLog.Write($"[ExcelImport] 読み込み失敗: {path}\n{ex}");
                if (showError)
                {
                    MessageBox.Show(string.Format(Loc.S("Import.ReadFailed"), ex.Message),
                        Loc.S("Common.Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                }
                ImportButton.IsEnabled = false;
                return false;
            }
        }

        /// <summary>
        /// インポート対象（＝表示中 かつ「取込」にチェックあり）の行。
        /// 絞り込みで隠れている行やチェックを外した行は取り込まない。
        /// </summary>
        private List<ImportPreviewRow> GetImportTargets()
        {
            if (_view == null) return new List<ImportPreviewRow>();
            return _view.Cast<ImportPreviewRow>().Where(r => r.IsSelected).ToList();
        }

        private void UpdateSummary()
        {
            if (_view == null) return;

            int changed = _changedRows.Count;
            int shown = _view.Count;
            int targets = GetImportTargets().Count;

            string summary = string.Format(Loc.S("Import.Summary"), _totalCount, changed, shown, targets);
            if (_readOnlyChangeCount > 0)
                summary += string.Format(Loc.S("Import.SummaryReadOnly"), _readOnlyChangeCount);
            if (_missingChangeCount > 0)
                summary += string.Format(Loc.S("Import.SummaryMissing"), _missingChangeCount);
            SummaryText.Text = summary;

            ImportButton.IsEnabled = targets > 0;
        }

        private void Row_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (_suppressSummary) return;
            if (e.PropertyName == nameof(ImportPreviewRow.IsSelected))
                UpdateSummary();
        }

        private void SelectAllVisible_Click(object sender, RoutedEventArgs e) => SetVisibleSelected(true);

        private void DeselectAllVisible_Click(object sender, RoutedEventArgs e) => SetVisibleSelected(false);

        /// <summary>表示中（絞り込み後）の行の「取込」チェックを一括で切り替える</summary>
        private void SetVisibleSelected(bool value)
        {
            if (_view == null) return;
            _suppressSummary = true;
            try
            {
                foreach (ImportPreviewRow r in _view)
                    r.IsSelected = value;
            }
            finally
            {
                _suppressSummary = false;
            }
            UpdateSummary();
        }

        private void ClearFilters_Click(object sender, RoutedEventArgs e)
        {
            _columnFilters.Clear();
            foreach (var key in _filterButtons.Keys.ToList())
                UpdateFilterIndicator(key);
            RefreshView();
        }

        private bool RowFilter(object o)
        {
            if (!(o is ImportPreviewRow r)) return false;
            foreach (var kv in _columnFilters)
            {
                if (_colAccessors.TryGetValue(kv.Key, out var acc) && !kv.Value.Contains(acc(r) ?? ""))
                    return false;
            }
            return true;
        }

        private void RefreshView()
        {
            _view?.Refresh();
            UpdateSummary();
        }

        // ===================== Excel風 列メニュー（並べ替え/フィルター） =====================

        /// <summary>列見出しを、並べ替え/絞り込みメニュー（▾）付きにする。</summary>
        private void RegisterColumn(string key, DataGridTextColumn col, string title,
                                    string sortPath, Func<ImportPreviewRow, string> accessor)
        {
            _colSortPaths[key] = sortPath;
            _colAccessors[key] = accessor;
            col.CanUserSort = false;   // 既定のヘッダークリックソートは使わず、▾ メニューに統一
            col.Header = BuildFilterHeader(key, title);
        }

        private FrameworkElement BuildFilterHeader(string key, string title)
        {
            var dock = new DockPanel { LastChildFill = true, HorizontalAlignment = HorizontalAlignment.Stretch };

            var btn = new Button
            {
                Content = "▾",
                Width = 20,
                Padding = new Thickness(0),
                Margin = new Thickness(4, 0, 0, 0),
                FontSize = 10,
                Focusable = false,
                Cursor = Cursors.Hand,
                Background = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground = System.Windows.Media.Brushes.Gray,
                Tag = key,
                ToolTip = Loc.S("Import.FilterHint"),
                VerticalAlignment = System.Windows.VerticalAlignment.Center
            };
            btn.Click += (s, e) => ShowColumnMenu(key, btn);
            DockPanel.SetDock(btn, Dock.Right);
            dock.Children.Add(btn);
            _filterButtons[key] = btn;

            dock.Children.Add(new TextBlock
            {
                Text = title,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                FontWeight = FontWeights.SemiBold
            });
            return dock;
        }

        /// <summary>
        /// 列メニュー（昇順/降順・値のチェックリストで絞り込み）を表示する。
        /// 候補値は「他の列の絞り込みを適用した行」から作る（Excel のオートフィルターと同じ）。
        /// </summary>
        private void ShowColumnMenu(string key, UIElement anchor)
        {
            CloseActivePopup();
            if (!_colAccessors.TryGetValue(key, out var accessor)) return;

            var candidates = _changedRows.Where(r => MatchesOtherFilters(r, key));
            var distinct = candidates.Select(r => accessor(r) ?? "").Distinct().ToList();
            if (key == "ElementId")
                distinct = distinct.OrderBy(v => long.TryParse(v, out var n) ? n : long.MaxValue).ToList();
            else
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

            // --- 値リスト（長い値は … で省略、全文はツールチップ）---
            var listPanel = new StackPanel();
            var checks = new List<CheckBox>();
            foreach (var val in distinct)
            {
                string disp = string.IsNullOrEmpty(val) ? Loc.S("ParamCleanup.Filter.Blank") : val;
                var cb = new CheckBox
                {
                    Tag = val,
                    IsChecked = allowed == null || allowed.Contains(val),
                    Margin = new Thickness(0, 1, 0, 1),
                    Content = new TextBlock
                    {
                        Text = disp,
                        TextTrimming = TextTrimming.CharacterEllipsis,
                        MaxWidth = 340,
                        ToolTip = disp
                    }
                };
                checks.Add(cb);
                listPanel.Children.Add(cb);
            }

            // --- 全選択 / 選択解除（検索で表示中の項目のみ対象）---
            var selRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 4) };
            var btnAll = new Button
            {
                Content = Loc.S("ParamCleanup.Filter.CheckAll"),
                Height = 24, Padding = new Thickness(8, 0, 8, 0), Margin = new Thickness(0, 0, 4, 0), Cursor = Cursors.Hand
            };
            btnAll.Click += (s, ev) =>
            {
                foreach (var c in checks)
                    if (c.Visibility == System.Windows.Visibility.Visible) c.IsChecked = true;
            };
            var btnNone = new Button
            {
                Content = Loc.S("ParamCleanup.Filter.UncheckAll"),
                Height = 24, Padding = new Thickness(8, 0, 8, 0), Cursor = Cursors.Hand
            };
            btnNone.Click += (s, ev) =>
            {
                foreach (var c in checks)
                    if (c.Visibility == System.Windows.Visibility.Visible) c.IsChecked = false;
            };
            selRow.Children.Add(btnAll);
            selRow.Children.Add(btnNone);
            root.Children.Add(selRow);

            search.TextChanged += (s, ev) =>
            {
                string q = search.Text.Trim();
                foreach (var c in checks)
                {
                    string t = (c.Tag as string) ?? "";
                    c.Visibility = (q.Length == 0 || t.IndexOf(q, StringComparison.CurrentCultureIgnoreCase) >= 0)
                        ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                }
            };

            root.Children.Add(new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                MaxHeight = 240,
                Content = listPanel
            });

            // --- クリア / OK / キャンセル ---
            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 6, 0, 0)
            };
            var btnClear = MenuButton(Loc.S("ParamCleanup.Filter.Clear"));
            btnClear.Width = 64; btnClear.Margin = new Thickness(0, 0, 4, 0);
            btnClear.HorizontalContentAlignment = HorizontalAlignment.Center;
            btnClear.Click += (s, ev) =>
            {
                _columnFilters.Remove(key);
                UpdateFilterIndicator(key);
                RefreshView();
                CloseActivePopup();
            };
            var btnOk = MenuButton(Loc.S("Common.OK"));
            btnOk.Width = 64; btnOk.Margin = new Thickness(0, 0, 4, 0);
            btnOk.HorizontalContentAlignment = HorizontalAlignment.Center;
            btnOk.Click += (s, ev) =>
            {
                var sel = new HashSet<string>(checks.Where(c => c.IsChecked == true).Select(c => (string)c.Tag));
                if (sel.Count >= distinct.Count) _columnFilters.Remove(key);  // 全選択＝絞り込みなし
                else _columnFilters[key] = sel;
                UpdateFilterIndicator(key);
                RefreshView();
                CloseActivePopup();
            };
            var btnCancel = MenuButton(Loc.S("Common.Cancel"));
            btnCancel.Width = 64;
            btnCancel.HorizontalContentAlignment = HorizontalAlignment.Center;
            btnCancel.Click += (s, ev) => CloseActivePopup();
            buttons.Children.Add(btnClear);
            buttons.Children.Add(btnOk);
            buttons.Children.Add(btnCancel);
            root.Children.Add(buttons);

            var border = new Border
            {
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0xCC, 0xCC, 0xCC)),
                BorderThickness = new Thickness(1),
                Child = root,
                MinWidth = 230,
                MaxWidth = 420
            };
            // 配置元ボタン（FontSize=10）からフォントを継承して小さくならないよう明示する
            System.Windows.Documents.TextElement.SetFontSize(border, 12d);

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

        /// <summary>指定列以外の絞り込み条件をすべて満たすか（列メニューの候補値の算出用）</summary>
        private bool MatchesOtherFilters(ImportPreviewRow r, string exceptKey)
        {
            foreach (var kv in _columnFilters)
            {
                if (kv.Key == exceptKey) continue;
                if (_colAccessors.TryGetValue(kv.Key, out var acc) && !kv.Value.Contains(acc(r) ?? ""))
                    return false;
            }
            return true;
        }

        private static Button MenuButton(string text)
        {
            return new Button
            {
                Content = text,
                Height = 26,
                Margin = new Thickness(0, 1, 0, 1),
                Padding = new Thickness(8, 0, 8, 0),
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Background = System.Windows.Media.Brushes.White,
                Cursor = Cursors.Hand
            };
        }

        private void SortByColumn(string key, ListSortDirection dir)
        {
            if (_view == null || !_colSortPaths.TryGetValue(key, out var path)) return;
            _view.SortDescriptions.Clear();
            _view.SortDescriptions.Add(new SortDescription(path, dir));
        }

        private void UpdateFilterIndicator(string key)
        {
            if (!_filterButtons.TryGetValue(key, out var btn)) return;
            bool active = _columnFilters.ContainsKey(key);
            btn.Foreground = active ? System.Windows.Media.Brushes.RoyalBlue : System.Windows.Media.Brushes.Gray;
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

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            ImportExecuted = true;
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        /// <summary>選択されたファイルパスを取得</summary>
        public string SelectedFilePath => _selectedFilePath;

        /// <summary>
        /// インポートに使うプレビュー行（書き込み・色付け用）。
        /// 変更プレビューで対象外にした行（絞り込みで非表示・「取込」のチェックなし）は除く。
        /// 読み取り専用の行は件数集計（スキップ数）のため残す。
        /// </summary>
        public List<ImportPreviewRow> PreviewRows
        {
            get
            {
                if (_previewRows == null) return null;
                var targets = new HashSet<ImportPreviewRow>(GetImportTargets());
                return _previewRows
                    .Where(r => !(r.HasChange && !r.IsReadOnly) || targets.Contains(r))
                    .ToList();
            }
        }
    }
}

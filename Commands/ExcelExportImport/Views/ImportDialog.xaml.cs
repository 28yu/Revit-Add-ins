using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
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
            PreviewDataGrid.Columns[0].Header = Loc.S("Import.ColElementId");
            PreviewDataGrid.Columns[1].Header = Loc.S("Import.ColCategory");
            PreviewDataGrid.Columns[2].Header = Loc.S("Import.ColParameter");
            PreviewDataGrid.Columns[3].Header = Loc.S("Import.ColCurrentValue");
            PreviewDataGrid.Columns[4].Header = Loc.S("Import.ColNewValue");
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
                var changedRows = _previewRows
                    .Where(r => r.HasChange && !r.IsReadOnly)
                    .OrderBy(r => r.CategoryName)
                    .ThenBy(r => r.ElementId)
                    .ToList();

                PreviewDataGrid.ItemsSource = changedRows;

                // サマリーを表示
                int writableChangeCount = changedRows.Count;
                int readOnlyChangeCount = _previewRows.Count(r => r.HasChange && r.IsReadOnly);
                int totalCount = _previewRows.Count;

                string summary = $"全{totalCount}件中  変更あり: {writableChangeCount}件";
                if (readOnlyChangeCount > 0)
                    summary += $"  読み取り専用で変更不可: {readOnlyChangeCount}件（長さ等の計算値）";
                SummaryText.Text = summary;

                ImportButton.IsEnabled = writableChangeCount > 0;
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

        /// <summary>プレビュー行を取得（色付け用）</summary>
        public List<ImportPreviewRow> PreviewRows => _previewRows;
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Microsoft.Win32;
using Tools28.Commands.ExcelExportImport.Services;

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
            _doc = doc;

            // 起動時に開いているExcelファイルを自動検出
            AutoDetectOpenFiles();
        }

        private void AutoDetectOpenFiles()
        {
            try
            {
                var openFiles = ExcelProcessHelper.GetOpenExcelFiles();
                if (openFiles.Count == 1)
                {
                    _selectedFilePath = openFiles[0];
                    FilePathTextBox.Text = _selectedFilePath;
                    LoadPreview();
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
                MessageBox.Show("開いているExcelファイルが見つかりません。",
                    "確認", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            else if (openFiles.Count == 1)
            {
                _selectedFilePath = openFiles[0];
                FilePathTextBox.Text = _selectedFilePath;
                LoadPreview();
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
                Title = "Excelファイルを選択",
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
                    Content = Path.GetFileName(file),
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
                Content = "キャンセル",
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

            if (selectWindow.ShowDialog() == true && selectedPath != null)
            {
                _selectedFilePath = selectedPath;
                FilePathTextBox.Text = _selectedFilePath;
                LoadPreview();
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excelファイル (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                _selectedFilePath = dialog.FileName;
                FilePathTextBox.Text = _selectedFilePath;
                LoadPreview();
            }
        }

        private void LoadPreview()
        {
            try
            {
                // シート情報を表示
                var sheetNames = ExcelImportService.GetSheetNames(_selectedFilePath);
                SheetInfoText.Text = $"シート数: {sheetNames.Count}  ({string.Join(", ", sheetNames)})";

                // プレビューを生成
                _previewRows = ExcelImportService.GeneratePreview(_doc, _selectedFilePath);

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
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ファイルの読み込みに失敗しました。\n{ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                ImportButton.IsEnabled = false;
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

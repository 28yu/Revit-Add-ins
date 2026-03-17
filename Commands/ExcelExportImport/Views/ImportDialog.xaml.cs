using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
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

                // 変更のあるものを先頭に表示
                var sorted = _previewRows
                    .OrderByDescending(r => r.HasChange)
                    .ThenBy(r => r.CategoryName)
                    .ThenBy(r => r.ElementId)
                    .ToList();

                PreviewDataGrid.ItemsSource = sorted;

                // サマリーを表示
                int changeCount = _previewRows.Count(r => r.HasChange);
                int readOnlyCount = _previewRows.Count(r => r.IsReadOnly);
                int totalCount = _previewRows.Count;
                SummaryText.Text = $"合計: {totalCount}件  変更あり: {changeCount}件  読み取り専用: {readOnlyCount}件";

                ImportButton.IsEnabled = changeCount > 0;
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
            int changeCount = _previewRows?.Count(r => r.HasChange) ?? 0;
            var confirmResult = MessageBox.Show(
                $"{changeCount}件の値を更新します。よろしいですか？",
                "インポート確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirmResult != MessageBoxResult.Yes)
                return;

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
    }
}

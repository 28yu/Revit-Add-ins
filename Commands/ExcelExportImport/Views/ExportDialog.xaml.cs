using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Autodesk.Revit.DB;
using Microsoft.Win32;
using Tools28.Commands.ExcelExportImport.Models;
using Tools28.Commands.ExcelExportImport.Services;

namespace Tools28.Commands.ExcelExportImport.Views
{
    /// <summary>
    /// EXCELエクスポートダイアログ
    /// </summary>
    public partial class ExportDialog : Window
    {
        private readonly Document _doc;

        // 全カテゴリ一覧
        private List<CategoryInfo> _allCategories;
        // 全パラメータ一覧（カテゴリ選択に応じて更新）
        private List<ParameterInfo> _allParameters = new List<ParameterInfo>();
        // 出力パラメータリスト
        private List<ParameterInfo> _outputParameters = new List<ParameterInfo>();

        /// <summary>エクスポート対象カテゴリ</summary>
        public List<CategoryInfo> SelectedCategories { get; private set; }

        /// <summary>エクスポート対象パラメータ（順序付き）</summary>
        public List<ParameterInfo> OutputParameters => _outputParameters;

        /// <summary>カテゴリ毎にシートを分けるか</summary>
        public bool SplitByCategory { get; private set; } = true;

        public ExportDialog(Document doc)
        {
            InitializeComponent();
            _doc = doc;

            // カテゴリ一覧を取得・表示
            _allCategories = RevitCategoryHelper.GetCategoriesWithElements(doc);
            CategoryListBox.ItemsSource = _allCategories;
        }

        #region カテゴリ選択

        private void CategoryListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // リストボックス選択変更時は特に処理不要（チェックボックスで制御）
        }

        private void CategoryCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            UpdateParameterList();
        }

        private void CategorySearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterCategoryList();
        }

        private void CategorySearchButton_Click(object sender, RoutedEventArgs e)
        {
            FilterCategoryList();
        }

        private void FilterCategoryList()
        {
            string filter = CategorySearchBox.Text.Trim();
            if (string.IsNullOrEmpty(filter))
            {
                CategoryListBox.ItemsSource = _allCategories;
            }
            else
            {
                CategoryListBox.ItemsSource = _allCategories
                    .Where(c => c.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
            }
        }

        #endregion

        #region パラメータ一覧

        private void UpdateParameterList()
        {
            _allParameters.Clear();

            var checkedCategories = _allCategories.Where(c => c.IsChecked).ToList();

            foreach (var cat in checkedCategories)
            {
                var parameters = ParameterService.GetParametersForCategory(
                    _doc, cat.BuiltInCategory, cat.Name);
                _allParameters.AddRange(parameters);
            }

            // 重複を除去し、出力リストに既にあるものは除外
            var outputDisplayNames = new HashSet<string>(_outputParameters.Select(p => p.DisplayName + "|" + p.CategoryName));
            var filteredParams = _allParameters
                .GroupBy(p => p.DisplayName + "|" + p.CategoryName)
                .Select(g => g.First())
                .Where(p => !outputDisplayNames.Contains(p.DisplayName + "|" + p.CategoryName))
                .ToList();

            FilterParameterList(filteredParams);
        }

        private void ParameterSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterParameterList(null);
        }

        private void ParameterSearchButton_Click(object sender, RoutedEventArgs e)
        {
            FilterParameterList(null);
        }

        private void FilterParameterList(List<ParameterInfo> source)
        {
            if (source == null)
            {
                // 現在のソースから再フィルタ
                var outputDisplayNames = new HashSet<string>(_outputParameters.Select(p => p.DisplayName + "|" + p.CategoryName));
                source = _allParameters
                    .GroupBy(p => p.DisplayName + "|" + p.CategoryName)
                    .Select(g => g.First())
                    .Where(p => !outputDisplayNames.Contains(p.DisplayName + "|" + p.CategoryName))
                    .ToList();
            }

            string filter = ParameterSearchBox.Text.Trim();
            List<ParameterInfo> filtered;
            if (string.IsNullOrEmpty(filter))
            {
                filtered = source;
            }
            else
            {
                filtered = source
                    .Where(p => p.DisplayName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
            }

            // カテゴリ別にグループ化して表示
            var view = new ListCollectionView(filtered);
            view.GroupDescriptions.Add(new PropertyGroupDescription("CategoryName"));
            ParameterListBox.ItemsSource = view;
        }

        #endregion

        #region 出力リスト操作

        private void AddToOutputButton_Click(object sender, RoutedEventArgs e)
        {
            // ListCollectionView からパラメータを列挙
            var checkedParams = new List<ParameterInfo>();
            if (ParameterListBox.ItemsSource is ListCollectionView view)
            {
                foreach (var item in view)
                {
                    if (item is ParameterInfo p && p.IsChecked)
                        checkedParams.Add(p);
                }
            }
            else if (ParameterListBox.ItemsSource is IEnumerable<ParameterInfo> list)
            {
                checkedParams = list.Where(p => p.IsChecked).ToList();
            }

            if (checkedParams.Count == 0) return;

            foreach (var param in checkedParams)
            {
                param.IsChecked = false;
                _outputParameters.Add(param);
            }

            RefreshOutputList();
            UpdateParameterList();
        }

        private void RemoveFromOutputButton_Click(object sender, RoutedEventArgs e)
        {
            var selected = OutputListBox.SelectedItem as ParameterInfo;
            if (selected == null) return;

            _outputParameters.Remove(selected);
            RefreshOutputList();
            UpdateParameterList();
        }

        private void MoveUpButton_Click(object sender, RoutedEventArgs e)
        {
            int index = OutputListBox.SelectedIndex;
            if (index <= 0) return;

            var item = _outputParameters[index];
            _outputParameters.RemoveAt(index);
            _outputParameters.Insert(index - 1, item);
            RefreshOutputList();
            OutputListBox.SelectedIndex = index - 1;
        }

        private void MoveDownButton_Click(object sender, RoutedEventArgs e)
        {
            int index = OutputListBox.SelectedIndex;
            if (index < 0 || index >= _outputParameters.Count - 1) return;

            var item = _outputParameters[index];
            _outputParameters.RemoveAt(index);
            _outputParameters.Insert(index + 1, item);
            RefreshOutputList();
            OutputListBox.SelectedIndex = index + 1;
        }

        private void RefreshOutputList()
        {
            string filter = OutputSearchBox.Text.Trim();
            List<ParameterInfo> source;
            if (string.IsNullOrEmpty(filter))
            {
                source = _outputParameters;
            }
            else
            {
                source = _outputParameters
                    .Where(p => p.DisplayName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
            }

            var view = new ListCollectionView(source);
            view.GroupDescriptions.Add(new PropertyGroupDescription("CategoryName"));
            OutputListBox.ItemsSource = view;
        }

        private void OutputSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshOutputList();
        }

        private void OutputSearchButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshOutputList();
        }

        #endregion

        #region 設定保存/読込

        private void SaveSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "JSON設定ファイル (*.json)|*.json",
                DefaultExt = ".json",
                FileName = "ExcelExportSettings"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var checkedCategories = _allCategories.Where(c => c.IsChecked).ToList();
                    var settings = SettingsService.CreateFromSelection(checkedCategories, _outputParameters);
                    SettingsService.SaveSettings(dialog.FileName, settings);
                    MessageBox.Show("設定を保存しました。", "設定保存", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"設定の保存に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void LoadSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "JSON設定ファイル (*.json)|*.json",
                DefaultExt = ".json"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var settings = SettingsService.LoadSettings(dialog.FileName);
                    ApplySettings(settings);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"設定の読み込みに失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ResetSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            // 全カテゴリのチェックを外す
            foreach (var cat in _allCategories)
                cat.IsChecked = false;

            _outputParameters.Clear();
            _allParameters.Clear();

            CategoryListBox.ItemsSource = null;
            CategoryListBox.ItemsSource = _allCategories;
            ParameterListBox.ItemsSource = null;
            RefreshOutputList();
        }

        private void ApplySettings(ExportSettings settings)
        {
            // カテゴリ選択を復元
            foreach (var cat in _allCategories)
            {
                cat.IsChecked = settings.SelectedCategories.Contains(cat.Name);
            }

            CategoryListBox.ItemsSource = null;
            CategoryListBox.ItemsSource = _allCategories;

            // パラメータ一覧を更新
            UpdateParameterList();

            // 出力パラメータを復元
            _outputParameters.Clear();
            foreach (var entry in settings.OutputParameters)
            {
                var match = _allParameters.FirstOrDefault(p =>
                    p.RawName == entry.RawName &&
                    p.IsTypeParameter == entry.IsTypeParameter &&
                    p.CategoryName == entry.CategoryName);

                if (match != null)
                {
                    _outputParameters.Add(match);
                }
            }

            RefreshOutputList();
            UpdateParameterList();
        }

        #endregion

        #region OK/キャンセル

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (_outputParameters.Count == 0)
            {
                MessageBox.Show("出力するパラメータを選択してください。", "確認",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SelectedCategories = _allCategories.Where(c => c.IsChecked).ToList();
            if (SelectedCategories.Count == 0)
            {
                MessageBox.Show("カテゴリを選択してください。", "確認",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SplitByCategory = SplitByCategoryCheckBox.IsChecked == true;
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        #endregion
    }
}

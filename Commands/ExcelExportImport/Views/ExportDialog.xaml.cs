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
using Tools28.Localization;

namespace Tools28.Commands.ExcelExportImport.Views
{
    /// <summary>
    /// EXCELエクスポートダイアログ
    /// </summary>
    public partial class ExportDialog : Window
    {
        private readonly Document _doc;
        private readonly ExportScope _scope;
        private readonly View _activeView;
        private readonly ICollection<ElementId> _selectionIds;

        // 全カテゴリ一覧
        private List<CategoryInfo> _allCategories;
        // 全パラメータ一覧（カテゴリ選択に応じて更新）
        private List<ParameterInfo> _allParameters = new List<ParameterInfo>();
        // 出力パラメータリスト
        private List<ParameterInfo> _outputParameters = new List<ParameterInfo>();

        // 全選択/選択解除でカテゴリのチェックを一括変更する際、
        // CheckBox の Checked/Unchecked からの UpdateParameterList を抑制するフラグ
        private bool _suppressCategoryUpdate;

        /// <summary>エクスポート対象カテゴリ</summary>
        public List<CategoryInfo> SelectedCategories { get; private set; }

        /// <summary>エクスポート対象パラメータ（順序付き）</summary>
        public List<ParameterInfo> OutputParameters => _outputParameters;

        /// <summary>カテゴリ毎にシートを分けるか</summary>
        public bool SplitByCategory { get; private set; } = true;

        public ExportDialog(Document doc)
            : this(doc, ExportScope.EntireProject, null, null)
        {
        }

        public ExportDialog(
            Document doc,
            ExportScope scope,
            View activeView,
            ICollection<ElementId> selectionIds)
        {
            InitializeComponent();
            ApplyLocalization();
            _doc = doc;
            _scope = scope;
            _activeView = activeView;
            _selectionIds = selectionIds;

            // 選択スコープに応じてカテゴリ一覧を取得・表示
            _allCategories = RevitCategoryHelper.GetCategoriesWithElements(
                doc, _scope, _activeView, _selectionIds);
            CategoryListBox.ItemsSource = _allCategories;
        }

        private void ApplyLocalization()
        {
            this.Title = Loc.S("Export.Title");
            grpCategory.Header = Loc.S("Export.Category");
            btnSearchCat.Content = Loc.S("Common.Search");
            grpParameter.Header = Loc.S("Export.Parameter");
            btnSearchParam.Content = Loc.S("Common.Search");
            btnCatSelectAll.Content = Loc.S("Common.SelectAll");
            btnCatSelectNone.Content = Loc.S("Common.SelectNone");
            btnParamSelectAll.Content = Loc.S("Common.SelectAll");
            btnParamSelectNone.Content = Loc.S("Common.SelectNone");
            ParamPrefixLegend.Text = Loc.S("Export.ParamPrefixLegend");
            btnAddToOutput.ToolTip = Loc.S("Export.AddToOutput");
            btnRemoveFromOutput.ToolTip = Loc.S("Export.RemoveFromOutput");
            grpOutput.Header = Loc.S("Export.Output");
            btnSearchOutput.Content = Loc.S("Common.Search");
            btnMoveUp.ToolTip = Loc.S("Export.MoveUp");
            btnMoveDown.ToolTip = Loc.S("Export.MoveDown");
            SplitByCategoryCheckBox.Content = Loc.S("Export.SeparateSheets");
            btnResetSettings.Content = Loc.S("Export.RestoreSettings");
            btnLoadSettings.Content = Loc.S("Export.LoadSettings");
            btnSaveSettings.Content = Loc.S("Export.SaveSettings");
            btnOK.Content = Loc.S("Common.OK");
            btnCancel.Content = Loc.S("Common.Cancel");
        }

        #region カテゴリ選択

        private void CategoryListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // リストボックス選択変更時は特に処理不要（チェックボックスで制御）
        }

        private void CategoryCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            // 全選択/選択解除による一括変更中は個別更新を抑制（最後に1回だけ更新する）
            if (_suppressCategoryUpdate) return;
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

        /// <summary>検索テキストで絞り込んだ、現在表示中のカテゴリ一覧を返す</summary>
        private List<CategoryInfo> GetVisibleCategories()
        {
            string filter = CategorySearchBox.Text.Trim();
            if (string.IsNullOrEmpty(filter))
                return _allCategories;

            return _allCategories
                .Where(c => c.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                         || c.DisplayLabel.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();
        }

        private void FilterCategoryList()
        {
            CategoryListBox.ItemsSource = GetVisibleCategories();
        }

        private void CategorySelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisibleCategoriesChecked(true);
        }

        private void CategorySelectNoneButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisibleCategoriesChecked(false);
        }

        /// <summary>現在表示中（検索絞り込み後）のカテゴリを一括でチェック/解除する</summary>
        private void SetVisibleCategoriesChecked(bool value)
        {
            _suppressCategoryUpdate = true;
            foreach (var cat in GetVisibleCategories())
                cat.IsChecked = value;
            _suppressCategoryUpdate = false;

            // 一括変更後にパラメータ一覧を1回だけ更新
            UpdateParameterList();
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
                    _doc, cat.BuiltInCategory, cat.Name, _scope, _activeView, _selectionIds);
                _allParameters.AddRange(parameters);
            }

            FilterParameterList(null);
        }

        /// <summary>
        /// 重複を除去し、出力リストに既にあるものを除いたパラメータ一覧（絞り込み前）を返す。
        /// </summary>
        private List<ParameterInfo> GetBaseParameters()
        {
            var outputDisplayNames = new HashSet<string>(_outputParameters.Select(p => p.DisplayName + "|" + p.CategoryName));
            return _allParameters
                .GroupBy(p => p.DisplayName + "|" + p.CategoryName)
                .Select(g => g.First())
                .Where(p => !outputDisplayNames.Contains(p.DisplayName + "|" + p.CategoryName))
                .ToList();
        }

        /// <summary>
        /// テキスト検索を適用した表示対象を返す。
        /// </summary>
        private List<ParameterInfo> ApplyParameterFilters(List<ParameterInfo> source)
        {
            string filter = ParameterSearchBox.Text.Trim();
            if (string.IsNullOrEmpty(filter))
                return source;

            return source
                .Where(p => p.DisplayName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();
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
                source = GetBaseParameters();

            var filtered = ApplyParameterFilters(source);

            // カテゴリ別にグループ化して表示
            var view = new ListCollectionView(filtered);
            view.GroupDescriptions.Add(new PropertyGroupDescription("CategoryName"));
            ParameterListBox.ItemsSource = view;
        }

        private void ParameterSelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisibleParametersChecked(true);
        }

        private void ParameterSelectNoneButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisibleParametersChecked(false);
        }

        /// <summary>現在表示中（検索絞り込み後）のパラメータを一括でチェック/解除する</summary>
        private void SetVisibleParametersChecked(bool value)
        {
            // 表示中の対象のみを一括変更（INotifyPropertyChanged で CheckBox に即時反映）
            foreach (var p in ApplyParameterFilters(GetBaseParameters()))
                p.IsChecked = value;
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

            if (dialog.ShowDialog(this) == true)
            {
                try
                {
                    var checkedCategories = _allCategories.Where(c => c.IsChecked).ToList();
                    var settings = SettingsService.CreateFromSelection(checkedCategories, _outputParameters);
                    SettingsService.SaveSettings(dialog.FileName, settings);
                    MessageBox.Show(Loc.S("Export.SettingsSaved"), Loc.S("Export.SettingsSaveTitle"), MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format(Loc.S("Export.SettingsSaveFailed"), ex.Message), Loc.S("Common.Error"), MessageBoxButton.OK, MessageBoxImage.Error);
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

            if (dialog.ShowDialog(this) == true)
            {
                try
                {
                    var settings = SettingsService.LoadSettings(dialog.FileName);
                    ApplySettings(settings);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format(Loc.S("Export.SettingsLoadFailed"), ex.Message), Loc.S("Common.Error"), MessageBoxButton.OK, MessageBoxImage.Error);
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
            // カテゴリ選択を復元（Contains の O(n) を避けるため HashSet 化）
            var selectedSet = new HashSet<string>(settings.SelectedCategories ?? new List<string>());
            foreach (var cat in _allCategories)
            {
                cat.IsChecked = selectedSet.Contains(cat.Name);
            }

            CategoryListBox.ItemsSource = null;
            CategoryListBox.ItemsSource = _allCategories;

            // パラメータ一覧を更新（Revitからのパラメータ取得はここで1回だけ実行）
            UpdateParameterList();

            // 出力パラメータの照合用に辞書を構築（線形探索の繰り返しを回避）
            var paramLookup = new Dictionary<string, ParameterInfo>();
            foreach (var p in _allParameters)
            {
                string key = p.RawName + "|" + p.IsTypeParameter + "|" + p.CategoryName;
                if (!paramLookup.ContainsKey(key))
                    paramLookup[key] = p;
            }

            // 出力パラメータを復元
            _outputParameters.Clear();
            foreach (var entry in settings.OutputParameters)
            {
                string key = entry.RawName + "|" + entry.IsTypeParameter + "|" + entry.CategoryName;
                if (paramLookup.TryGetValue(key, out var match))
                {
                    _outputParameters.Add(match);
                }
            }

            RefreshOutputList();

            // 出力に移したパラメータを中央リストから除外して再表示。
            // ここで UpdateParameterList() を再呼び出しすると Revit へのパラメータ取得が
            // もう一度走って重いため、取得済みの _allParameters から表示だけ更新する。
            FilterParameterList(null);
        }

        #endregion

        #region OK/キャンセル

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (_outputParameters.Count == 0)
            {
                MessageBox.Show(Loc.S("Export.SelectParams"), Loc.S("Common.Confirm"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SelectedCategories = _allCategories.Where(c => c.IsChecked).ToList();
            if (SelectedCategories.Count == 0)
            {
                MessageBox.Show(Loc.S("Export.SelectCategory"), Loc.S("Common.Confirm"),
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

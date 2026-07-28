using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        // 出力パラメータリスト（エクスポート順の情報源）
        private List<ParameterInfo> _outputParameters = new List<ParameterInfo>();

        // 出力欄の表示用コレクション。ItemsSource を作り直さず要素の Move で
        // 並べ替えることで、上下移動時にスクロール位置を維持する。
        private readonly ObservableCollection<ParameterInfo> _outputDisplay = new ObservableCollection<ParameterInfo>();
        private ListCollectionView _outputView;

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
            btnClearOutput.Content = Loc.S("Export.ClearOutput");
            btnClearOutput.ToolTip = Loc.S("Export.ClearOutput.Desc");
            btnMoveUp.ToolTip = Loc.S("Export.MoveUp");
            btnMoveDown.ToolTip = Loc.S("Export.MoveDown");
            SplitByCategoryLabel.Text = Loc.S("Export.SeparateSheets");
            // 設定ボタンはコンパクト表示。説明はホバー時のツールチップで出す。
            btnLoadSettings.Content = Loc.S("Export.LoadSettings");
            btnLoadSettings.ToolTip = Loc.S("Export.LoadSettings.Desc");
            btnLoadSettingsFromExcel.Content = Loc.S("Export.LoadSettingsFromExcel");
            btnLoadSettingsFromExcel.ToolTip = Loc.S("Export.LoadSettingsFromExcel.Desc");
            btnSaveSettings.Content = Loc.S("Export.SaveSettings");
            btnSaveSettings.ToolTip = Loc.S("Export.SaveSettings.Desc");
            btnOK.Content = Loc.S("Export.RunExport");
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
            var selected = OutputListBox.SelectedItems.Cast<ParameterInfo>().ToList();
            if (selected.Count == 0) return;

            foreach (var p in selected)
                _outputParameters.Remove(p);
            RefreshOutputList();
            UpdateParameterList();
        }

        private void ClearOutputButton_Click(object sender, RoutedEventArgs e)
        {
            if (_outputParameters.Count == 0) return;

            // 並べ替え等の作業を失うため、誤操作防止に確認する
            var result = MessageBox.Show(
                Loc.S("Export.ClearOutputConfirm"), Loc.S("Common.Confirm"),
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes) return;

            _outputParameters.Clear();
            RefreshOutputList();

            // クリアした分を選択中カテゴリのパラメータ欄へ戻す
            UpdateParameterList();
        }

        private void MoveUpButton_Click(object sender, RoutedEventArgs e)
        {
            MoveSelectedOutput(-1);
        }

        private void MoveDownButton_Click(object sender, RoutedEventArgs e)
        {
            MoveSelectedOutput(+1);
        }

        /// <summary>
        /// 出力リストで選択中のパラメータ（複数可）を上(-1)/下(+1)に移動する。
        /// 表示はカテゴリでグループ化されているため、カテゴリ単位で「選択ブロック」を
        /// 1つ分ずらす。選択した要素そのものを基準にするため、表示位置と内部リストの
        /// ずれや、複数選択・飛び選択でも常に選択したものが動く。
        /// </summary>
        private void MoveSelectedOutput(int direction)
        {
            var selected = OutputListBox.SelectedItems.Cast<ParameterInfo>().ToList();
            if (selected.Count == 0) return;

            var selectedSet = new HashSet<ParameterInfo>(selected);

            bool anyMoved = false;
            // 選択に含まれるカテゴリごとに、そのカテゴリ内の並びだけを対象に移動する
            foreach (var cat in selected.Select(p => p.CategoryName).Distinct().ToList())
            {
                // このカテゴリのパラメータが _outputParameters 上で占める位置（順序保持）
                var positions = new List<int>();
                for (int i = 0; i < _outputParameters.Count; i++)
                    if (_outputParameters[i].CategoryName == cat)
                        positions.Add(i);

                var seq = positions.Select(i => _outputParameters[i]).ToList();
                if (ReorderBlock(seq, selectedSet, direction))
                {
                    anyMoved = true;
                    // 並べ替え結果を同じ位置（他カテゴリの位置は動かさない）へ書き戻す
                    for (int k = 0; k < positions.Count; k++)
                        _outputParameters[positions[k]] = seq[k];
                }
            }

            if (!anyMoved) return;

            // 表示コレクションを _outputParameters の並びへ同期（可能なら Move でスクロール維持）
            SyncDisplayToParameters();

            // 選択を復元し、移動方向側の端が見えるようにする
            OutputListBox.SelectedItems.Clear();
            foreach (var p in selected)
                OutputListBox.SelectedItems.Add(p);
            OutputListBox.ScrollIntoView(direction < 0 ? selected[0] : selected[selected.Count - 1]);
        }

        /// <summary>
        /// カテゴリ内の並び seq に対し、選択要素を1つ分だけ上(-1)/下(+1)へ移動する。
        /// 隣が選択要素なら詰めず（ブロックを保つ）、端では止まる。移動が起きたら true。
        /// </summary>
        private static bool ReorderBlock(List<ParameterInfo> seq, HashSet<ParameterInfo> selectedSet, int direction)
        {
            bool moved = false;
            if (direction < 0)
            {
                for (int i = 1; i < seq.Count; i++)
                {
                    if (selectedSet.Contains(seq[i]) && !selectedSet.Contains(seq[i - 1]))
                    {
                        var t = seq[i]; seq[i] = seq[i - 1]; seq[i - 1] = t;
                        moved = true;
                    }
                }
            }
            else
            {
                for (int i = seq.Count - 2; i >= 0; i--)
                {
                    if (selectedSet.Contains(seq[i]) && !selectedSet.Contains(seq[i + 1]))
                    {
                        var t = seq[i]; seq[i] = seq[i + 1]; seq[i + 1] = t;
                        moved = true;
                    }
                }
            }
            return moved;
        }

        /// <summary>
        /// 表示コレクション(_outputDisplay)を _outputParameters の並びへ合わせる。
        /// 検索フィルタが無い場合は要素を Move で並べ替えてスクロール位置を維持する。
        /// フィルタ中は表示が部分集合のため素直に作り直す。
        /// </summary>
        private void SyncDisplayToParameters()
        {
            if (!string.IsNullOrEmpty(OutputSearchBox.Text.Trim()))
            {
                RefreshOutputList();
                return;
            }

            // 同一要素集合の並べ替えなので、先頭から順に位置合わせ（各ズレを Move で解消）
            for (int i = 0; i < _outputParameters.Count && i < _outputDisplay.Count; i++)
            {
                if (ReferenceEquals(_outputDisplay[i], _outputParameters[i])) continue;

                int j = -1;
                for (int k = i + 1; k < _outputDisplay.Count; k++)
                {
                    if (ReferenceEquals(_outputDisplay[k], _outputParameters[i])) { j = k; break; }
                }
                if (j > i)
                    _outputDisplay.Move(j, i);
                else
                {
                    // 想定外（フィルタ等で不整合）→ 安全側に作り直す
                    RefreshOutputList();
                    return;
                }
            }
        }

        private void RefreshOutputList()
        {
            // 表示用の ListCollectionView（グループ化）を一度だけ生成して使い回す。
            // 以降は _outputDisplay の中身を差し替えるだけにして、並べ替え時の
            // ItemsSource 再設定（＝スクロールリセット）を避ける。
            if (_outputView == null)
            {
                _outputView = new ListCollectionView(_outputDisplay);
                _outputView.GroupDescriptions.Add(new PropertyGroupDescription("CategoryName"));
                OutputListBox.ItemsSource = _outputView;
            }

            string filter = OutputSearchBox.Text.Trim();
            IEnumerable<ParameterInfo> source = string.IsNullOrEmpty(filter)
                ? _outputParameters
                : _outputParameters.Where(p => p.DisplayName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);

            _outputDisplay.Clear();
            foreach (var p in source)
                _outputDisplay.Add(p);
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

        /// <summary>
        /// このアドインで書き出し済みの Excel ファイルから出力設定を読み込む。
        /// 設定JSONを保存し忘れた場合や、他ユーザーが書き出した Excel の並びを
        /// そのまま再利用したい場合に使う。
        /// </summary>
        private void LoadSettingsFromExcelButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excelファイル (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx"
            };

            if (dialog.ShowDialog(this) != true) return;

            try
            {
                var settings = ExportSettingsExcelReader.ReadFromExcel(dialog.FileName);

                if (settings.OutputParameters.Count == 0)
                {
                    // 出力設定として解釈できる列が無い（このアドイン以外で作った Excel 等）
                    MessageBox.Show(Loc.S("Export.ExcelSettingsNoParams"),
                        Loc.S("Common.Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                ApplySettings(settings);

                // 現在の対象カテゴリに存在せず復元できなかったものがあれば知らせる
                int restored = _outputParameters.Count;
                int total = settings.OutputParameters.Count;
                if (restored < total)
                {
                    MessageBox.Show(
                        string.Format(Loc.S("Export.ExcelSettingsPartial"), restored, total),
                        Loc.S("Common.Confirm"), MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(Loc.S("Export.SettingsLoadFailed"), ex.Message),
                    Loc.S("Common.Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

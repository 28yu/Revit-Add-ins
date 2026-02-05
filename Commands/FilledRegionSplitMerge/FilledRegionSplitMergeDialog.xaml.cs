using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FilledRegionSplitMerge
{
    public partial class FilledRegionSplitMergeDialog : Window
    {
        private FilledRegionHelper.SelectionAnalysis _analysis;

        public enum OperationType
        {
            Split,
            Merge
        }

        public OperationType SelectedOperation { get; private set; }
        public FilledRegionType SelectedPattern { get; private set; }

        public FilledRegionSplitMergeDialog(FilledRegionHelper.SelectionAnalysis analysis)
        {
            InitializeComponent();
            _analysis = analysis;
            InitializeDialog();
        }

        private void InitializeDialog()
        {
            // ラジオボタンの有効/無効設定
            RadioSplit.IsEnabled = _analysis.CanSplit;
            RadioMerge.IsEnabled = _analysis.CanMerge;

            // デフォルト選択
            if (_analysis.CanSplit && !_analysis.CanMerge)
            {
                RadioSplit.IsChecked = true;
            }
            else if (!_analysis.CanSplit && _analysis.CanMerge)
            {
                RadioMerge.IsChecked = true;
            }
            else if (_analysis.CanSplit && _analysis.CanMerge)
            {
                RadioSplit.IsChecked = true;
            }

            // パターンリストを設定
            ComboPattern.ItemsSource = _analysis.AvailableTypes;
            if (_analysis.DefaultType != null)
            {
                ComboPattern.SelectedItem = _analysis.DefaultType;
            }
            else if (_analysis.AvailableTypes.Count > 0)
            {
                ComboPattern.SelectedIndex = 0;
            }

            // 情報表示を更新
            UpdateInfo();
        }

        private void RadioOperation_Checked(object sender, RoutedEventArgs e)
        {
            if (RadioMerge == null || GroupPattern == null)
                return;

            // 統合選択時のみパターン選択を有効化
            GroupPattern.IsEnabled = RadioMerge.IsChecked == true;

            UpdateInfo();
        }

        private void UpdateInfo()
        {
            if (TextInfo == null || _analysis == null)
                return;

            if (RadioSplit?.IsChecked == true)
            {
                var multiAreaRegions = _analysis.FilledRegions
                    .Where(fr => fr.GetBoundaries().Count > 1)
                    .ToList();

                int totalAreas = multiAreaRegions.Sum(fr => fr.GetBoundaries().Count);

                TextInfo.Text = $"選択: {_analysis.FilledRegions.Count}個の領域（分割可能: {multiAreaRegions.Count}個、合計{totalAreas}エリア）";
            }
            else if (RadioMerge?.IsChecked == true)
            {
                int totalAreas = _analysis.FilledRegions.Sum(fr => fr.GetBoundaries().Count);
                TextInfo.Text = $"選択: {_analysis.FilledRegions.Count}個の領域（合計{totalAreas}エリア）→ 1個の領域に統合";
            }
            else
            {
                TextInfo.Text = $"選択: {_analysis.FilledRegions.Count}個の領域";
            }
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            // 操作が選択されているか確認
            if (RadioSplit.IsChecked != true && RadioMerge.IsChecked != true)
            {
                MessageBox.Show("操作（分割または統合）を選択してください。",
                    "入力エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // 統合時はパターンが選択されているか確認
            if (RadioMerge.IsChecked == true)
            {
                if (ComboPattern.SelectedItem == null)
                {
                    MessageBox.Show("塗り潰しパターンを選択してください。",
                        "入力エラー",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                SelectedPattern = ComboPattern.SelectedItem as FilledRegionType;
            }

            // 選択された操作を設定
            SelectedOperation = RadioSplit.IsChecked == true
                ? OperationType.Split
                : OperationType.Merge;

            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}

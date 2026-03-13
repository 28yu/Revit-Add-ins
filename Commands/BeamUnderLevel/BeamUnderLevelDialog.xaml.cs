using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamUnderLevel
{
    public partial class BeamUnderLevelDialog : Window
    {
        private readonly BeamUnderLevelDialogData _data;
        private int _currentStep = 1;
        private const int TotalSteps = 3;

        // ファミリ毎のパラメータ選択コンボボックス
        private readonly Dictionary<string, ComboBox> _familyParamComboBoxes =
            new Dictionary<string, ComboBox>();

        // ファミリ毎のカスタム入力テキストボックス
        private readonly Dictionary<string, TextBox> _familyCustomTextBoxes =
            new Dictionary<string, TextBox>();

        // ファミリ毎の「カスタム入力」ラジオボタン
        private readonly Dictionary<string, RadioButton> _familyCustomRadios =
            new Dictionary<string, RadioButton>();

        // 結果プロパティ
        public Level SelectedLowerLevel { get; private set; }
        public Dictionary<string, string> FamilyParamSelection { get; private set; }
        public bool OverwriteExistingFilters { get; private set; }

        public BeamUnderLevelDialog(BeamUnderLevelDialogData data)
        {
            InitializeComponent();
            _data = data;
            InitializeStep1();
        }

        #region ステップ1: レベル設定

        private void InitializeStep1()
        {
            ViewNameText.Text = _data.ViewName;
            BeamCountText.Text = $"{_data.BeamCount} 本";

            double refElevMm = BeamCalculator.FeetToMm(_data.RefLevel.Elevation);
            RefLevelText.Text = $"{_data.RefLevel.Name} ({refElevMm:+0;-0}mm)";

            // 上位レベルドロップダウン
            var levelItems = _data.LowerLevels
                .Select(l => new LevelItem(l))
                .ToList();
            LowerLevelComboBox.ItemsSource = levelItems;

            // デフォルト選択
            var defaultItem = levelItems.FirstOrDefault(li =>
                li.Level.Id == _data.DefaultLowerLevel.Id);
            if (defaultItem != null)
                LowerLevelComboBox.SelectedItem = defaultItem;
            else if (levelItems.Count > 0)
                LowerLevelComboBox.SelectedIndex = 0;

            UpdateFloorHeight();
        }

        private void LowerLevelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateFloorHeight();
        }

        private void UpdateFloorHeight()
        {
            var selectedItem = LowerLevelComboBox.SelectedItem as LevelItem;
            if (selectedItem != null)
            {
                double floorHeightFeet = selectedItem.Level.Elevation - _data.RefLevel.Elevation;
                double floorHeightMm = BeamCalculator.FeetToMm(floorHeightFeet);
                FloorHeightText.Text = $"{floorHeightMm:0}mm";
                SelectedLowerLevel = selectedItem.Level;
            }
        }

        #endregion

        #region ステップ2: パラメータ選択

        private void InitializeStep2()
        {
            FamilyParamPanel.Children.Clear();
            _familyParamComboBoxes.Clear();
            _familyCustomTextBoxes.Clear();
            _familyCustomRadios.Clear();

            foreach (var entry in _data.BeamsByFamily)
            {
                string familyName = entry.Key;
                int beamCount = entry.Value.Count;
                var candidates = _data.ParamCandidates.ContainsKey(familyName)
                    ? _data.ParamCandidates[familyName]
                    : new List<ParamCandidate>();

                // ファミリグループ
                var groupBox = new Border
                {
                    Background = System.Windows.Media.Brushes.White,
                    BorderBrush = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0xD3, 0xD3, 0xD3)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(12, 10, 12, 10),
                    Margin = new Thickness(0, 0, 0, 10)
                };

                var stackPanel = new StackPanel();

                // ファミリ名ヘッダー
                var header = new TextBlock
                {
                    Text = $"{familyName} ({beamCount}梁)",
                    FontSize = 13,
                    FontWeight = FontWeights.Bold,
                    Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0x33, 0x33, 0x33)),
                    Margin = new Thickness(0, 0, 0, 8)
                };
                stackPanel.Children.Add(header);

                if (candidates.Count > 0)
                {
                    // パラメータ候補をラジオボタンで表示
                    string groupName = $"family_{familyName.GetHashCode()}";
                    bool firstSelected = false;

                    foreach (var candidate in candidates.OrderByDescending(c => c.DetectedCount))
                    {
                        var radio = new RadioButton
                        {
                            Content = $"{candidate.ParamName}  ({candidate.DetectedCount}梁で検出)",
                            GroupName = groupName,
                            Tag = candidate.ParamName,
                            FontSize = 11,
                            Margin = new Thickness(10, 3, 0, 3),
                            IsChecked = !firstSelected
                        };

                        if (!firstSelected)
                        {
                            firstSelected = true;
                        }

                        stackPanel.Children.Add(radio);
                    }

                    // カスタム入力オプション
                    var customPanel = new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        Margin = new Thickness(10, 3, 0, 3)
                    };

                    var customRadio = new RadioButton
                    {
                        Content = "カスタム: ",
                        GroupName = groupName,
                        FontSize = 11,
                        VerticalAlignment = VerticalAlignment.Center
                    };

                    var customTextBox = new TextBox
                    {
                        Width = 150,
                        Height = 22,
                        FontSize = 11,
                        IsEnabled = false,
                        VerticalContentAlignment = VerticalAlignment.Center,
                        Padding = new Thickness(4, 0, 4, 0)
                    };

                    customRadio.Checked += (s, e) => { customTextBox.IsEnabled = true; };
                    customRadio.Unchecked += (s, e) => { customTextBox.IsEnabled = false; };

                    _familyCustomRadios[familyName] = customRadio;
                    _familyCustomTextBoxes[familyName] = customTextBox;

                    customPanel.Children.Add(customRadio);
                    customPanel.Children.Add(customTextBox);
                    stackPanel.Children.Add(customPanel);
                }
                else
                {
                    // 候補なし
                    var noParamText = new TextBlock
                    {
                        Text = "梁高さパラメータが見つかりませんでした。カスタム入力してください:",
                        FontSize = 11,
                        Foreground = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(0xCC, 0x00, 0x00)),
                        Margin = new Thickness(10, 0, 0, 5)
                    };
                    stackPanel.Children.Add(noParamText);

                    var customTextBox = new TextBox
                    {
                        Width = 200,
                        Height = 22,
                        FontSize = 11,
                        Margin = new Thickness(10, 0, 0, 0),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalContentAlignment = VerticalAlignment.Center,
                        Padding = new Thickness(4, 0, 4, 0)
                    };
                    _familyCustomTextBoxes[familyName] = customTextBox;
                    stackPanel.Children.Add(customTextBox);
                }

                groupBox.Child = stackPanel;
                FamilyParamPanel.Children.Add(groupBox);
            }
        }

        /// <summary>
        /// ステップ2からパラメータ選択結果を収集
        /// </summary>
        private Dictionary<string, string> CollectParamSelections()
        {
            var result = new Dictionary<string, string>();

            foreach (var entry in _data.BeamsByFamily)
            {
                string familyName = entry.Key;

                // カスタム入力がアクティブか確認
                if (_familyCustomRadios.ContainsKey(familyName) &&
                    _familyCustomRadios[familyName].IsChecked == true)
                {
                    string customValue = _familyCustomTextBoxes[familyName].Text?.Trim();
                    if (!string.IsNullOrEmpty(customValue))
                        result[familyName] = customValue;
                    continue;
                }

                // カスタムテキストボックスのみ（候補なしの場合）
                if (!_familyCustomRadios.ContainsKey(familyName) &&
                    _familyCustomTextBoxes.ContainsKey(familyName))
                {
                    string customValue = _familyCustomTextBoxes[familyName].Text?.Trim();
                    if (!string.IsNullOrEmpty(customValue))
                        result[familyName] = customValue;
                    continue;
                }

                // ラジオボタンから選択を取得
                foreach (var child in FamilyParamPanel.Children)
                {
                    var border = child as Border;
                    if (border == null) continue;
                    var panel = border.Child as StackPanel;
                    if (panel == null) continue;

                    // このファミリのパネルか確認
                    var headerText = panel.Children.OfType<TextBlock>().FirstOrDefault();
                    if (headerText == null || !headerText.Text.StartsWith(familyName))
                        continue;

                    // 選択されたラジオボタンを探す
                    foreach (var panelChild in panel.Children)
                    {
                        var radio = panelChild as RadioButton;
                        if (radio != null && radio.IsChecked == true && radio.Tag is string paramName)
                        {
                            result[familyName] = paramName;
                            break;
                        }
                    }
                    break;
                }
            }

            return result;
        }

        #endregion

        #region ステップ3: 確認

        private void InitializeStep3()
        {
            FamilyParamSelection = CollectParamSelections();

            string summary = $"ビュー: {_data.ViewName}\n" +
                $"対象梁数: {_data.BeamCount} 本\n" +
                $"ファミリ数: {_data.BeamsByFamily.Count}\n\n";

            summary += $"参照レベル: {_data.RefLevel.Name}\n";
            summary += $"上位レベル: {SelectedLowerLevel.Name}\n";
            double floorHeightMm = BeamCalculator.FeetToMm(
                SelectedLowerLevel.Elevation - _data.RefLevel.Elevation);
            summary += $"階高: {floorHeightMm:0}mm\n\n";

            summary += "ファミリ別パラメータ:\n";
            foreach (var entry in FamilyParamSelection)
            {
                int count = _data.BeamsByFamily.ContainsKey(entry.Key)
                    ? _data.BeamsByFamily[entry.Key].Count : 0;
                summary += $"  {entry.Key} ({count}梁) → {entry.Value}\n";
            }

            // パラメータ未選択のファミリ
            var unselected = _data.BeamsByFamily.Keys
                .Where(f => !FamilyParamSelection.ContainsKey(f))
                .ToList();
            if (unselected.Count > 0)
            {
                summary += "\n未選択のファミリ (処理スキップ):\n";
                foreach (var f in unselected)
                {
                    int count = _data.BeamsByFamily[f].Count;
                    summary += $"  {f} ({count}梁)\n";
                }
            }

            SummaryText.Text = summary;
        }

        #endregion

        #region ナビゲーション

        private void ShowStep(int step)
        {
            _currentStep = step;

            Step1Panel.Visibility = step == 1 ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            Step2Panel.Visibility = step == 2 ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            Step3Panel.Visibility = step == 3 ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

            BackButton.IsEnabled = step > 1;
            NextButton.Content = step == TotalSteps ? "実行" : "次へ";

            string[] stepNames = { "", "レベル設定", "梁高さパラメータ選択", "処理確認" };
            StepIndicator.Text = $"ステップ {step} / {TotalSteps}  {stepNames[step]}";
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentStep == 1)
            {
                // バリデーション
                if (SelectedLowerLevel == null)
                {
                    MessageBox.Show("上位レベルを選択してください。", "入力エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                InitializeStep2();
                ShowStep(2);
            }
            else if (_currentStep == 2)
            {
                // パラメータ選択の収集と検証
                var selections = CollectParamSelections();
                if (selections.Count == 0)
                {
                    MessageBox.Show(
                        "少なくとも1つのファミリで梁高さパラメータを選択してください。",
                        "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                InitializeStep3();
                ShowStep(3);
            }
            else if (_currentStep == 3)
            {
                // 実行
                FamilyParamSelection = CollectParamSelections();
                OverwriteExistingFilters = OverwriteCheckBox.IsChecked == true;
                DialogResult = true;
                Close();
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentStep > 1)
            {
                ShowStep(_currentStep - 1);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }

        #endregion
    }
}

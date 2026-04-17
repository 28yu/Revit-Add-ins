using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.BeamUnderLevel
{
    public partial class BeamUnderLevelDialog : Window
    {
        private readonly BeamUnderLevelDialogData _data;
        private int _currentStep = 1;
        private const int TotalSteps = 4;

            // ファミリ毎の「その他」ラジオボタン（梁高さ）
        private readonly Dictionary<string, RadioButton> _familyCustomRadios =
            new Dictionary<string, RadioButton>();
        // ファミリ毎の「その他」ComboBox（梁高さ）
        private readonly Dictionary<string, ComboBox> _familyCustomComboBoxes =
            new Dictionary<string, ComboBox>();

        // ファミリ毎の天端レベルパラメータ選択用
        private readonly Dictionary<string, RadioButton> _familyTopLevelCustomRadios =
            new Dictionary<string, RadioButton>();
        private readonly Dictionary<string, ComboBox> _familyTopLevelCustomComboBoxes =
            new Dictionary<string, ComboBox>();

        // 結果プロパティ
        public Level SelectedLowerLevel { get; private set; }
        public Dictionary<string, string> FamilyParamSelection { get; private set; }
        public Dictionary<string, string> FamilyTopLevelParamSelection { get; private set; }
        public bool OverwriteExistingFilters { get; private set; }
        public ElementId SelectedTextNoteTypeId { get; private set; }
        public ElementId SelectedBeamLabelTypeId { get; private set; }

        public BeamUnderLevelDialog(BeamUnderLevelDialogData data)
        {
            InitializeComponent();
            ApplyLocalization();
            _data = data;
            InitializeStep1();
        }

        private void ApplyLocalization()
        {
            this.Title = Loc.S("BeamUnder.Title");
            StepIndicator.Text = Loc.S("BeamUnder.Step1");
            txtViewInfo.Text = Loc.S("BeamUnder.ViewInfo");
            txtViewName.Text = Loc.S("BeamUnder.ViewName");
            txtBeamCount.Text = Loc.S("BeamUnder.BeamCount");
            txtTextTypeSetting.Text = Loc.S("BeamUnder.TextTypeSetting");
            txtLegendTextType.Text = Loc.S("BeamUnder.LegendTextType");
            txtLabelTextType.Text = Loc.S("BeamUnder.LabelTextType");
            txtTextTypeHint.Text = Loc.S("BeamUnder.TextTypeHint");
            txtLevelSetting.Text = Loc.S("BeamUnder.LevelSetting");
            txtRefLevel.Text = Loc.S("BeamUnder.RefLevel");
            txtUpperLevel.Text = Loc.S("BeamUnder.UpperLevel");
            txtFloorHeight.Text = Loc.S("BeamUnder.FloorHeight");
            txtLevelHint.Text = Loc.S("BeamUnder.LevelHint");
            txtHeightParam.Text = Loc.S("BeamUnder.HeightParam");
            txtHeightParamHint.Text = Loc.S("BeamUnder.HeightParamHint");
            txtTopParam.Text = Loc.S("BeamUnder.TopParam");
            txtTopParamHint.Text = Loc.S("BeamUnder.TopParamHint");
            txtSummary.Text = Loc.S("BeamUnder.Summary");
            txtProcessContent.Text = Loc.S("BeamUnder.ProcessContent");
            chkSharedParam.Content = Loc.S("BeamUnder.SharedParam");
            chkCalcValue.Content = Loc.S("BeamUnder.CalcValue");
            chkCreateFilter.Content = Loc.S("BeamUnder.CreateFilter");
            chkCreateLegend.Content = Loc.S("BeamUnder.CreateLegend");
            txtExistingFilter.Text = Loc.S("BeamUnder.ExistingFilter");
            OverwriteCheckBox.Content = Loc.S("BeamUnder.OverwriteFilter");
            txtOverwriteHint.Text = Loc.S("BeamUnder.OverwriteHint");
            BackButton.Content = Loc.S("Common.Back");
            NextButton.Content = Loc.S("Common.Next");
            CancelButton.Content = Loc.S("Common.Cancel");
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

            // 文字タイプドロップダウン
            if (_data.TextNoteTypes != null && _data.TextNoteTypes.Count > 0)
            {
                TextNoteTypeComboBox.ItemsSource = _data.TextNoteTypes;
                TextNoteTypeComboBox.SelectedIndex = 0;

                BeamLabelTypeComboBox.ItemsSource = _data.TextNoteTypes;
                BeamLabelTypeComboBox.SelectedIndex = 0;
            }

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
            _familyCustomComboBoxes.Clear();
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

                    // その他パラメータ（ComboBox）
                    var additionalParams = _data.AdditionalHeightParams != null &&
                        _data.AdditionalHeightParams.ContainsKey(familyName)
                        ? _data.AdditionalHeightParams[familyName]
                        : new List<string>();

                    if (additionalParams.Count > 0)
                    {
                        var customPanel = new StackPanel
                        {
                            Orientation = Orientation.Horizontal,
                            Margin = new Thickness(10, 3, 0, 3)
                        };

                        var customRadio = new RadioButton
                        {
                            Content = "その他: ",
                            GroupName = groupName,
                            FontSize = 11,
                            VerticalAlignment = VerticalAlignment.Center
                        };

                        var customComboBox = new ComboBox
                        {
                            Width = 250,
                            Height = 24,
                            FontSize = 11,
                            IsEnabled = true,
                            VerticalContentAlignment = VerticalAlignment.Center
                        };

                        foreach (string paramName in additionalParams)
                        {
                            customComboBox.Items.Add(paramName);
                        }
                        if (customComboBox.Items.Count > 0)
                            customComboBox.SelectedIndex = 0;

                        // ComboBox選択時にラジオボタンを自動チェック
                        customComboBox.DropDownOpened += (s, e) => { customRadio.IsChecked = true; };
                        customComboBox.SelectionChanged += (s, e) =>
                        {
                            if (customComboBox.IsDropDownOpen || customComboBox.IsFocused)
                                customRadio.IsChecked = true;
                        };

                        _familyCustomRadios[familyName] = customRadio;
                        _familyCustomComboBoxes[familyName] = customComboBox;

                        customPanel.Children.Add(customRadio);
                        customPanel.Children.Add(customComboBox);
                        stackPanel.Children.Add(customPanel);
                    }
                }
                else
                {
                    // 候補なし
                    var noParamText = new TextBlock
                    {
                        Text = "梁高さパラメータが見つかりませんでした。",
                        FontSize = 11,
                        Foreground = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(0xCC, 0x00, 0x00)),
                        Margin = new Thickness(10, 0, 0, 5)
                    };
                    stackPanel.Children.Add(noParamText);
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

                // 「その他」ComboBoxがアクティブか確認
                if (_familyCustomRadios.ContainsKey(familyName) &&
                    _familyCustomRadios[familyName].IsChecked == true)
                {
                    if (_familyCustomComboBoxes.ContainsKey(familyName))
                    {
                        string selected = _familyCustomComboBoxes[familyName].SelectedItem as string;
                        if (!string.IsNullOrEmpty(selected))
                            result[familyName] = selected;
                    }
                    continue;
                }

                // ラジオボタンから選択を取得
                foreach (var child in FamilyParamPanel.Children)
                {
                    var border = child as Border;
                    if (border == null) continue;
                    var panel = border.Child as StackPanel;
                    if (panel == null) continue;

                    var headerText = panel.Children.OfType<TextBlock>().FirstOrDefault();
                    if (headerText == null || !headerText.Text.StartsWith(familyName))
                        continue;

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

        #region ステップ3: 梁天端レベルパラメータ選択

        private void InitializeStep3()
        {
            FamilyTopLevelParamPanel.Children.Clear();
            _familyTopLevelCustomRadios.Clear();
            _familyTopLevelCustomComboBoxes.Clear();

            foreach (var entry in _data.BeamsByFamily)
            {
                string familyName = entry.Key;
                int beamCount = entry.Value.Count;
                var candidates = _data.TopLevelParamCandidates.ContainsKey(familyName)
                    ? _data.TopLevelParamCandidates[familyName]
                    : new List<ParamCandidate>();
                var additionalParams = _data.AdditionalLevelParams != null &&
                    _data.AdditionalLevelParams.ContainsKey(familyName)
                    ? _data.AdditionalLevelParams[familyName]
                    : new List<string>();

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

                string groupName = $"toplevel_{familyName.GetHashCode()}";
                bool firstSelected = false;

                // 主要候補をラジオボタンで表示
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
                        firstSelected = true;

                    stackPanel.Children.Add(radio);
                }

                // その他のパラメータ（ComboBox）
                if (additionalParams.Count > 0)
                {
                    var customPanel = new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        Margin = new Thickness(10, 3, 0, 3)
                    };

                    var customRadio = new RadioButton
                    {
                        Content = "その他: ",
                        GroupName = groupName,
                        FontSize = 11,
                        VerticalAlignment = VerticalAlignment.Center,
                        IsChecked = !firstSelected
                    };

                    if (!firstSelected)
                        firstSelected = true;

                    var customComboBox = new ComboBox
                    {
                        Width = 250,
                        Height = 24,
                        FontSize = 11,
                        IsEnabled = true,
                        VerticalContentAlignment = VerticalAlignment.Center
                    };

                    foreach (string paramName in additionalParams)
                    {
                        customComboBox.Items.Add(paramName);
                    }
                    if (customComboBox.Items.Count > 0)
                        customComboBox.SelectedIndex = 0;

                    // ComboBox選択時にラジオボタンを自動チェック
                    customComboBox.DropDownOpened += (s, e) => { customRadio.IsChecked = true; };
                    customComboBox.SelectionChanged += (s, e) =>
                    {
                        if (customComboBox.IsDropDownOpen || customComboBox.IsFocused)
                            customRadio.IsChecked = true;
                    };

                    _familyTopLevelCustomRadios[familyName] = customRadio;
                    _familyTopLevelCustomComboBoxes[familyName] = customComboBox;

                    customPanel.Children.Add(customRadio);
                    customPanel.Children.Add(customComboBox);
                    stackPanel.Children.Add(customPanel);
                }
                else if (candidates.Count == 0)
                {
                    var noParamText = new TextBlock
                    {
                        Text = "天端レベル関連のパラメータが見つかりませんでした。",
                        FontSize = 11,
                        Foreground = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(0xCC, 0x00, 0x00)),
                        Margin = new Thickness(10, 0, 0, 5)
                    };
                    stackPanel.Children.Add(noParamText);
                }

                groupBox.Child = stackPanel;
                FamilyTopLevelParamPanel.Children.Add(groupBox);
            }
        }

        /// <summary>
        /// ステップ3から天端レベルパラメータ選択結果を収集
        /// </summary>
        private Dictionary<string, string> CollectTopLevelParamSelections()
        {
            var result = new Dictionary<string, string>();

            foreach (var entry in _data.BeamsByFamily)
            {
                string familyName = entry.Key;

                // 「その他」ComboBoxがアクティブか確認
                if (_familyTopLevelCustomRadios.ContainsKey(familyName) &&
                    _familyTopLevelCustomRadios[familyName].IsChecked == true)
                {
                    if (_familyTopLevelCustomComboBoxes.ContainsKey(familyName))
                    {
                        string selected = _familyTopLevelCustomComboBoxes[familyName].SelectedItem as string;
                        if (!string.IsNullOrEmpty(selected))
                            result[familyName] = selected;
                    }
                    continue;
                }

                // ラジオボタンから選択を取得
                foreach (var child in FamilyTopLevelParamPanel.Children)
                {
                    var border = child as Border;
                    if (border == null) continue;
                    var panel = border.Child as StackPanel;
                    if (panel == null) continue;

                    var headerText = panel.Children.OfType<TextBlock>().FirstOrDefault();
                    if (headerText == null || !headerText.Text.StartsWith(familyName))
                        continue;

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

        #region ステップ4: 確認

        private void InitializeStep4()
        {
            FamilyParamSelection = CollectParamSelections();
            FamilyTopLevelParamSelection = CollectTopLevelParamSelections();

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
                string topParam = FamilyTopLevelParamSelection.ContainsKey(entry.Key)
                    ? FamilyTopLevelParamSelection[entry.Key] : "未選択";
                summary += $"  {entry.Key} ({count}梁)\n";
                summary += $"    梁高さ: {entry.Value}\n";
                summary += $"    天端レベル: {topParam}\n";
            }

            // パラメータ未選択のファミリ
            var unselected = _data.BeamsByFamily.Keys
                .Where(f => !FamilyParamSelection.ContainsKey(f) ||
                            !FamilyTopLevelParamSelection.ContainsKey(f))
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
            Step4Panel.Visibility = step == 4 ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

            BackButton.IsEnabled = step > 1;
            NextButton.Content = step == TotalSteps ? Loc.S("Common.Execute") : Loc.S("Common.Next");

            string[] stepNames = { "", Loc.S("BeamUnder.StepName1"), Loc.S("BeamUnder.StepName2"), Loc.S("BeamUnder.StepName3"), Loc.S("BeamUnder.StepName4") };
            StepIndicator.Text = Loc.S("Common.StepIndicator", step, TotalSteps, stepNames[step]);
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
                // 天端レベルパラメータ選択の収集と検証
                var selections = CollectTopLevelParamSelections();
                if (selections.Count == 0)
                {
                    MessageBox.Show(
                        "少なくとも1つのファミリで天端レベルパラメータを選択してください。",
                        "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                InitializeStep4();
                ShowStep(4);
            }
            else if (_currentStep == 4)
            {
                // 実行
                FamilyParamSelection = CollectParamSelections();
                FamilyTopLevelParamSelection = CollectTopLevelParamSelections();
                OverwriteExistingFilters = OverwriteCheckBox.IsChecked == true;
                var selectedTextType = TextNoteTypeComboBox.SelectedItem as TextNoteTypeItem;
                SelectedTextNoteTypeId = selectedTextType?.Id;
                var selectedBeamLabelType = BeamLabelTypeComboBox.SelectedItem as TextNoteTypeItem;
                SelectedBeamLabelTypeId = selectedBeamLabelType?.Id;
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FireProtection
{
    public partial class FireProtectionDialog : Window
    {
        private readonly FireProtectionDialogData _data;
        private int _currentStep = 1;
        private const int TotalSteps = 4;

        private readonly List<FireProtectionTypeEntry> _typeEntries =
            new List<FireProtectionTypeEntry>();
        private readonly Dictionary<string, TextBox> _perTypeOffsetBoxes =
            new Dictionary<string, TextBox>();
        private readonly Dictionary<string, ComboBox> _beamAssignmentCombos =
            new Dictionary<string, ComboBox>();

        public FireProtectionDialog(FireProtectionDialogData data)
        {
            InitializeComponent();
            _data = data;
            InitializeStep1();
        }

        public FireProtectionResult GetResult()
        {
            var result = new FireProtectionResult
            {
                Types = new List<FireProtectionTypeEntry>(_typeEntries),
                UseCommonOffset = CommonOffsetRadio.IsChecked == true,
                CommonOffsetMm = ParseDouble(CommonOffsetInput.Text, 50),
                BeamTypeAssignments = CollectBeamAssignments(),
                OverwriteExisting = OverwriteCheckBox.IsChecked == true
            };

            var selectedLineStyle = LineStyleComboBox.SelectedItem as LineStyleItem;
            result.LineStyleId = selectedLineStyle?.Id;

            var selectedFillPattern = FillPatternComboBox.SelectedItem as FillPatternItem;
            result.FillPatternId = selectedFillPattern?.Id;

            var selectedTextType = TextNoteTypeComboBox.SelectedItem as FpTextNoteTypeItem;
            result.TextNoteTypeId = selectedTextType?.Id;

            if (!result.UseCommonOffset)
            {
                foreach (var entry in _typeEntries)
                {
                    if (_perTypeOffsetBoxes.ContainsKey(entry.Name))
                    {
                        entry.OffsetMm = ParseDouble(
                            _perTypeOffsetBoxes[entry.Name].Text, 50);
                    }
                }
            }

            return result;
        }

        #region Step 1: 耐火被覆種類の定義

        private void InitializeStep1()
        {
            ViewNameText.Text = _data.ViewName;
            BeamCountText.Text = $"{_data.BeamCount} 本";
        }

        private void AddType_Click(object sender, RoutedEventArgs e)
        {
            AddTypeFromInput();
        }

        private void TypeNameInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddTypeFromInput();
                e.Handled = true;
            }
        }

        private void AddTypeFromInput()
        {
            string name = TypeNameInput.Text.Trim();
            if (string.IsNullOrEmpty(name))
                return;

            if (_typeEntries.Any(t => t.Name == name))
            {
                MessageBox.Show($"「{name}」は既に追加されています。",
                    "重複", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _typeEntries.Add(new FireProtectionTypeEntry
            {
                Name = name,
                OffsetMm = ParseDouble(CommonOffsetInput.Text, 50)
            });

            TypeNameInput.Text = "";
            TypeNameInput.Focus();
            RefreshTypeListUI();
        }

        private void RefreshTypeListUI()
        {
            TypeListPanel.Children.Clear();
            _perTypeOffsetBoxes.Clear();
            bool showPerTypeOffset = PerTypeOffsetRadio.IsChecked == true;

            foreach (var entry in _typeEntries)
            {
                var panel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 2, 0, 2)
                };

                var nameBlock = new TextBlock
                {
                    Text = entry.Name,
                    FontSize = 12,
                    Width = 200,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(8, 0, 0, 0)
                };
                panel.Children.Add(nameBlock);

                if (showPerTypeOffset)
                {
                    var offsetLabel = new TextBlock
                    {
                        Text = "オフセット:",
                        FontSize = 11,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(10, 0, 4, 0)
                    };
                    panel.Children.Add(offsetLabel);

                    var offsetBox = new TextBox
                    {
                        Width = 60,
                        Height = 24,
                        FontSize = 11,
                        Padding = new Thickness(4, 2, 4, 2),
                        Text = entry.OffsetMm.ToString("0"),
                        VerticalContentAlignment = VerticalAlignment.Center
                    };
                    panel.Children.Add(offsetBox);
                    _perTypeOffsetBoxes[entry.Name] = offsetBox;

                    var mmLabel = new TextBlock
                    {
                        Text = "mm",
                        FontSize = 11,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(2, 0, 0, 0)
                    };
                    panel.Children.Add(mmLabel);
                }

                string entryName = entry.Name;
                var removeButton = new Button
                {
                    Content = "×",
                    Width = 24,
                    Height = 24,
                    FontSize = 12,
                    Margin = new Thickness(8, 0, 0, 0),
                    Cursor = Cursors.Hand,
                    VerticalAlignment = VerticalAlignment.Center,
                    Background = Brushes.White,
                    BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xAD, 0xAD, 0xAD)),
                    BorderThickness = new Thickness(1)
                };
                removeButton.Click += (s, ev) =>
                {
                    _typeEntries.RemoveAll(t => t.Name == entryName);
                    RefreshTypeListUI();
                };
                panel.Children.Add(removeButton);

                TypeListPanel.Children.Add(panel);
            }
        }

        private void OffsetMode_Changed(object sender, RoutedEventArgs e)
        {
            if (CommonOffsetInput != null)
            {
                CommonOffsetInput.IsEnabled = CommonOffsetRadio.IsChecked == true;
            }
            RefreshTypeListUI();
        }

        #endregion

        #region Step 2: 梁の分類

        private void InitializeStep2()
        {
            BeamAssignmentPanel.Children.Clear();
            _beamAssignmentCombos.Clear();

            var options = new List<string> { "除外" };
            options.AddRange(_typeEntries.Select(t => t.Name));

            foreach (var beamType in _data.BeamTypes)
            {
                var border = new Border
                {
                    Background = Brushes.White,
                    BorderBrush = new SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0xD3, 0xD3, 0xD3)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(12, 8, 12, 8),
                    Margin = new Thickness(0, 0, 0, 6)
                };

                var grid = new Grid();
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(50, GridUnitType.Pixel) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180, GridUnitType.Pixel) });

                var nameBlock = new TextBlock
                {
                    Text = beamType.DisplayName,
                    FontSize = 12,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = beamType.DisplayName
                };
                Grid.SetColumn(nameBlock, 0);
                grid.Children.Add(nameBlock);

                var countBlock = new TextBlock
                {
                    Text = $"{beamType.Count}本",
                    FontSize = 11,
                    Foreground = new SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0x66, 0x66, 0x66)),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center
                };
                Grid.SetColumn(countBlock, 1);
                grid.Children.Add(countBlock);

                var combo = new ComboBox
                {
                    Height = 26,
                    FontSize = 11,
                    VerticalContentAlignment = VerticalAlignment.Center
                };
                foreach (var opt in options)
                    combo.Items.Add(opt);

                combo.SelectedIndex = options.Count > 1 ? 1 : 0;

                Grid.SetColumn(combo, 2);
                grid.Children.Add(combo);

                _beamAssignmentCombos[beamType.DisplayName] = combo;

                border.Child = grid;
                BeamAssignmentPanel.Children.Add(border);
            }
        }

        private Dictionary<string, string> CollectBeamAssignments()
        {
            var assignments = new Dictionary<string, string>();
            foreach (var kvp in _beamAssignmentCombos)
            {
                string selected = kvp.Value.SelectedItem as string;
                if (!string.IsNullOrEmpty(selected))
                    assignments[kvp.Key] = selected;
            }
            return assignments;
        }

        #endregion

        #region Step 3: 表示設定

        private void InitializeStep3()
        {
            if (LineStyleComboBox.Items.Count == 0)
            {
                LineStyleComboBox.ItemsSource = _data.LineStyles;
                if (_data.LineStyles.Count > 0)
                {
                    var thinLines = _data.LineStyles.FirstOrDefault(
                        ls => ls.Name.Contains("Thin") || ls.Name.Contains("細線"));
                    LineStyleComboBox.SelectedItem = thinLines ?? _data.LineStyles[0];
                }
            }

            if (FillPatternComboBox.Items.Count == 0)
            {
                FillPatternComboBox.ItemsSource = _data.FillPatterns;
                if (_data.FillPatterns.Count > 0)
                {
                    var solidFill = _data.FillPatterns.FirstOrDefault(
                        fp => fp.Name.Contains("Solid") || fp.Name.Contains("ベタ塗り"));
                    FillPatternComboBox.SelectedItem = solidFill ?? _data.FillPatterns[0];
                }
            }

            if (TextNoteTypeComboBox.Items.Count == 0)
            {
                TextNoteTypeComboBox.ItemsSource = _data.TextNoteTypes;
                if (_data.TextNoteTypes.Count > 0)
                    TextNoteTypeComboBox.SelectedIndex = 0;
            }

            RefreshColorPreview();
        }

        private void RefreshColorPreview()
        {
            ColorPreviewPanel.Children.Clear();
            var colors = FilledRegionCreator.GenerateColors(_typeEntries.Count);

            for (int i = 0; i < _typeEntries.Count; i++)
            {
                var panel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 3, 0, 3)
                };

                var colorRect = new System.Windows.Shapes.Rectangle
                {
                    Width = 30,
                    Height = 18,
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(
                        colors[i].Red, colors[i].Green, colors[i].Blue)),
                    Stroke = new SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0x99, 0x99, 0x99)),
                    StrokeThickness = 1,
                    Margin = new Thickness(0, 0, 10, 0)
                };
                panel.Children.Add(colorRect);

                var nameBlock = new TextBlock
                {
                    Text = _typeEntries[i].Name,
                    FontSize = 12,
                    VerticalAlignment = VerticalAlignment.Center
                };
                panel.Children.Add(nameBlock);

                ColorPreviewPanel.Children.Add(panel);
            }
        }

        #endregion

        #region Step 4: 確認

        private void InitializeStep4()
        {
            var assignments = CollectBeamAssignments();

            string summary = $"ビュー: {_data.ViewName}\n" +
                $"対象梁数: {_data.BeamCount} 本\n\n";

            summary += $"耐火被覆種類: {_typeEntries.Count}\n";
            foreach (var entry in _typeEntries)
            {
                double offset = CommonOffsetRadio.IsChecked == true
                    ? ParseDouble(CommonOffsetInput.Text, 50)
                    : entry.OffsetMm;
                summary += $"  - {entry.Name}（オフセット: {offset}mm）\n";
            }

            summary += "\n梁の分類:\n";
            foreach (var bt in _data.BeamTypes)
            {
                string assigned;
                if (!assignments.TryGetValue(bt.DisplayName, out assigned))
                    assigned = "未設定";
                summary += $"  {bt.DisplayName} ({bt.Count}本) → {assigned}\n";
            }

            var selectedLine = LineStyleComboBox.SelectedItem as LineStyleItem;
            var selectedPattern = FillPatternComboBox.SelectedItem as FillPatternItem;
            summary += $"\n境界線: {(selectedLine != null ? selectedLine.Name : "なし")}\n";
            summary += $"塗りパターン: {(selectedPattern != null ? selectedPattern.Name : "なし")}";

            SummaryText.Text = summary;
        }

        #endregion

        #region Navigation

        private void ShowStep(int step)
        {
            _currentStep = step;

            Step1Panel.Visibility = step == 1 ? Visibility.Visible : Visibility.Collapsed;
            Step2Panel.Visibility = step == 2 ? Visibility.Visible : Visibility.Collapsed;
            Step3Panel.Visibility = step == 3 ? Visibility.Visible : Visibility.Collapsed;
            Step4Panel.Visibility = step == 4 ? Visibility.Visible : Visibility.Collapsed;

            BackButton.IsEnabled = step > 1;
            NextButton.Content = step == TotalSteps ? "実行" : "次へ";

            string[] stepNames = { "",
                "耐火被覆種類の定義",
                "梁の分類",
                "表示設定",
                "処理確認" };
            StepIndicator.Text = $"ステップ {step} / {TotalSteps}  {stepNames[step]}";
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentStep == 1)
            {
                if (_typeEntries.Count == 0)
                {
                    MessageBox.Show("耐火被覆の種類を1つ以上追加してください。",
                        "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SavePerTypeOffsets();
                InitializeStep2();
                ShowStep(2);
            }
            else if (_currentStep == 2)
            {
                var assignments = CollectBeamAssignments();
                bool hasAssignment = assignments.Any(a => a.Value != "除外");
                if (!hasAssignment)
                {
                    MessageBox.Show("少なくとも1つの梁タイプを耐火被覆種類に割り当ててください。",
                        "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                InitializeStep3();
                ShowStep(3);
            }
            else if (_currentStep == 3)
            {
                InitializeStep4();
                ShowStep(4);
            }
            else if (_currentStep == 4)
            {
                DialogResult = true;
                Close();
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentStep == 2)
            {
                ShowStep(1);
            }
            else if (_currentStep > 1)
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

        #region Helpers

        private void SavePerTypeOffsets()
        {
            if (PerTypeOffsetRadio.IsChecked == true)
            {
                foreach (var entry in _typeEntries)
                {
                    if (_perTypeOffsetBoxes.ContainsKey(entry.Name))
                    {
                        entry.OffsetMm = ParseDouble(
                            _perTypeOffsetBoxes[entry.Name].Text, 50);
                    }
                }
            }
        }

        private static double ParseDouble(string text, double defaultValue)
        {
            double val;
            if (double.TryParse(text, out val))
                return val;
            return defaultValue;
        }

        #endregion
    }
}

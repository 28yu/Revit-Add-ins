using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WpfRectangle = System.Windows.Shapes.Rectangle;
using Autodesk.Revit.DB;
using WinForms = System.Windows.Forms;
using WpfGrid = System.Windows.Controls.Grid;

namespace Tools28.Commands.FireProtection
{
    public partial class FireProtectionDialog : Window
    {
        private readonly FireProtectionDialogData _data;
        private int _currentStep = 1;
        private const int TotalSteps = 3;

        private List<FireProtectionTypeEntry> _typeEntries = new List<FireProtectionTypeEntry>();
        private readonly Dictionary<string, TextBox> _perTypeOffsetBoxes =
            new Dictionary<string, TextBox>();

        public FireProtectionDialog(FireProtectionDialogData data)
        {
            InitializeComponent();
            _data = data;
            InitializeStep1();
        }

        public FireProtectionResult GetResult()
        {
            SavePerTypeOffsets();

            var result = new FireProtectionResult
            {
                IncludeBeams = IncludeBeamsCheck.IsChecked == true,
                IncludeColumns = IncludeColumnsCheck.IsChecked == true,
                Types = new List<FireProtectionTypeEntry>(_typeEntries),
                UseCommonOffset = CommonOffsetRadio.IsChecked == true,
                CommonOffsetMm = ParseDouble(CommonOffsetInput.Text, 50),
                OverwriteExisting = OverwriteCheckBox.IsChecked == true
            };

            var selectedParam = ParameterComboBox.SelectedItem as FireProtectionParameterInfo;
            result.SelectedParameterName = selectedParam?.ParameterName;

            var selectedLine = LineStyleComboBox.SelectedItem as LineStyleItem;
            result.LineStyleId = selectedLine?.Id;

            var selectedPattern = FillPatternComboBox.SelectedItem as FillPatternItem;
            result.FillPatternId = selectedPattern?.Id;

            var selectedText = TextNoteTypeComboBox.SelectedItem as FpTextNoteTypeItem;
            result.TextNoteTypeId = selectedText?.Id;

            return result;
        }

        #region Step 1: 基本設定

        private void InitializeStep1()
        {
            IncludeBeamsCheck.IsChecked = _data.HasBeams;
            IncludeBeamsCheck.IsEnabled = _data.HasBeams;
            IncludeColumnsCheck.IsChecked = _data.HasColumns;
            IncludeColumnsCheck.IsEnabled = _data.HasColumns;

            UpdateCategoryCount();
            UpdateParameterList();
        }

        private void Category_Changed(object sender, RoutedEventArgs e)
        {
            if (ParameterComboBox == null) return;
            UpdateCategoryCount();
            UpdateParameterList();
        }

        private void UpdateCategoryCount()
        {
            int count = 0;
            if (IncludeBeamsCheck.IsChecked == true) count += _data.BeamCount;
            if (IncludeColumnsCheck.IsChecked == true) count += _data.ColumnCount;
            CategoryCountText.Text = $"※ 対象要素数: {count}";
        }

        private void UpdateParameterList()
        {
            var allParams = new Dictionary<string, FireProtectionParameterInfo>();

            if (IncludeBeamsCheck.IsChecked == true && _data.BeamParameters != null)
            {
                foreach (var p in _data.BeamParameters)
                {
                    if (!allParams.ContainsKey(p.ParameterName))
                        allParams[p.ParameterName] = new FireProtectionParameterInfo
                        {
                            ParameterName = p.ParameterName,
                            DetectedCount = p.DetectedCount,
                            UniqueValues = new List<string>(p.UniqueValues)
                        };
                    else
                    {
                        allParams[p.ParameterName].DetectedCount += p.DetectedCount;
                        foreach (var v in p.UniqueValues)
                        {
                            if (!allParams[p.ParameterName].UniqueValues.Contains(v))
                                allParams[p.ParameterName].UniqueValues.Add(v);
                        }
                    }
                }
            }

            if (IncludeColumnsCheck.IsChecked == true && _data.ColumnParameters != null)
            {
                foreach (var p in _data.ColumnParameters)
                {
                    if (!allParams.ContainsKey(p.ParameterName))
                        allParams[p.ParameterName] = new FireProtectionParameterInfo
                        {
                            ParameterName = p.ParameterName,
                            DetectedCount = p.DetectedCount,
                            UniqueValues = new List<string>(p.UniqueValues)
                        };
                    else
                    {
                        allParams[p.ParameterName].DetectedCount += p.DetectedCount;
                        foreach (var v in p.UniqueValues)
                        {
                            if (!allParams[p.ParameterName].UniqueValues.Contains(v))
                                allParams[p.ParameterName].UniqueValues.Add(v);
                        }
                    }
                }
            }

            var paramList = allParams.Values
                .OrderByDescending(p => p.DetectedCount)
                .ToList();

            ParameterComboBox.ItemsSource = paramList;
            ParameterComboBox.DisplayMemberPath = null;

            if (paramList.Count > 0)
            {
                // DisplayMemberPath を使わず ToString で表示
                foreach (var p in paramList)
                    p.ParameterName = $"{p.ParameterName}（{p.DetectedCount}件検出）";

                ParameterComboBox.DisplayMemberPath = "ParameterName";
                ParameterComboBox.SelectedIndex = 0;
                ParameterNoteText.Text = $"※ {paramList.Count} 個のパラメータを検出しました";
            }
            else
            {
                ParameterNoteText.Text = "※「耐火被覆」を含むパラメータが見つかりません";
                DetectedTypesPanel.Children.Clear();
                _typeEntries.Clear();
            }
        }

        private void Parameter_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (ParameterComboBox == null) return;
            var selected = ParameterComboBox.SelectedItem as FireProtectionParameterInfo;
            if (selected == null) return;

            var colors = FilledRegionCreator.GenerateColors(selected.UniqueValues.Count);
            _typeEntries = new List<FireProtectionTypeEntry>();

            for (int i = 0; i < selected.UniqueValues.Count; i++)
            {
                _typeEntries.Add(new FireProtectionTypeEntry
                {
                    Name = selected.UniqueValues[i],
                    OffsetMm = ParseDouble(CommonOffsetInput.Text, 50),
                    ColorR = colors[i].Red,
                    ColorG = colors[i].Green,
                    ColorB = colors[i].Blue
                });
            }

            RefreshDetectedTypesUI();
            RefreshPerTypeOffsetUI();
        }

        private void RefreshDetectedTypesUI()
        {
            DetectedTypesPanel.Children.Clear();

            if (_typeEntries.Count == 0)
            {
                DetectedTypesPanel.Children.Add(new TextBlock
                {
                    Text = "（パラメータに値が設定された要素がありません）",
                    FontSize = 11,
                    Foreground = new SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0x99, 0x99, 0x99)),
                    Margin = new Thickness(4, 4, 0, 0)
                });
                return;
            }

            for (int i = 0; i < _typeEntries.Count; i++)
            {
                var entry = _typeEntries[i];
                int idx = i;

                var panel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 2, 0, 2)
                };

                var colorRect = new WpfRectangle
                {
                    Width = 22,
                    Height = 16,
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(
                        entry.ColorR, entry.ColorG, entry.ColorB)),
                    Stroke = new SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0x66, 0x66, 0x66)),
                    StrokeThickness = 1,
                    Margin = new Thickness(4, 0, 8, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    Cursor = Cursors.Hand
                };
                colorRect.MouseLeftButtonDown += (s, ev) =>
                {
                    ShowColorPicker(idx);
                };
                panel.Children.Add(colorRect);

                panel.Children.Add(new TextBlock
                {
                    Text = entry.Name,
                    FontSize = 12,
                    VerticalAlignment = VerticalAlignment.Center
                });

                DetectedTypesPanel.Children.Add(panel);
            }
        }

        private void OffsetMode_Changed(object sender, RoutedEventArgs e)
        {
            if (CommonOffsetInput == null || PerTypeOffsetPanel == null) return;
            CommonOffsetInput.IsEnabled = CommonOffsetRadio.IsChecked == true;
            RefreshPerTypeOffsetUI();
        }

        private void RefreshPerTypeOffsetUI()
        {
            if (PerTypeOffsetPanel == null) return;
            PerTypeOffsetPanel.Children.Clear();
            _perTypeOffsetBoxes.Clear();

            if (PerTypeOffsetRadio.IsChecked != true) return;

            foreach (var entry in _typeEntries)
            {
                var panel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(20, 2, 0, 2)
                };

                panel.Children.Add(new TextBlock
                {
                    Text = entry.Name,
                    FontSize = 11,
                    Width = 180,
                    VerticalAlignment = VerticalAlignment.Center
                });

                var offsetBox = new TextBox
                {
                    Width = 60, Height = 24, FontSize = 11,
                    Padding = new Thickness(4, 2, 4, 2),
                    Text = entry.OffsetMm.ToString("0"),
                    VerticalContentAlignment = VerticalAlignment.Center
                };
                panel.Children.Add(offsetBox);
                _perTypeOffsetBoxes[entry.Name] = offsetBox;

                panel.Children.Add(new TextBlock
                {
                    Text = " mm", FontSize = 11,
                    VerticalAlignment = VerticalAlignment.Center
                });

                PerTypeOffsetPanel.Children.Add(panel);
            }
        }

        #endregion

        #region Step 2: 表示設定

        private void InitializeStep2()
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
                        fp => fp.Name.Contains("Solid") || fp.Name.Contains("ベタ"));
                    FillPatternComboBox.SelectedItem = solidFill ?? _data.FillPatterns[0];
                }
            }

            if (TextNoteTypeComboBox.Items.Count == 0)
            {
                TextNoteTypeComboBox.ItemsSource = _data.TextNoteTypes;
                if (_data.TextNoteTypes.Count > 0)
                    TextNoteTypeComboBox.SelectedIndex = 0;
            }
        }

        private void ShowColorPicker(int typeIndex)
        {
            if (typeIndex < 0 || typeIndex >= _typeEntries.Count) return;

            var entry = _typeEntries[typeIndex];
            var colorDialog = new WinForms.ColorDialog
            {
                FullOpen = true,
                Color = System.Drawing.Color.FromArgb(
                    entry.ColorR, entry.ColorG, entry.ColorB)
            };

            if (colorDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                entry.ColorR = colorDialog.Color.R;
                entry.ColorG = colorDialog.Color.G;
                entry.ColorB = colorDialog.Color.B;
                RefreshDetectedTypesUI();
            }
        }

        #endregion

        #region Step 3: 確認

        private void InitializeStep3()
        {
            SavePerTypeOffsets();

            string catText = "";
            if (IncludeBeamsCheck.IsChecked == true) catText += "梁";
            if (IncludeColumnsCheck.IsChecked == true)
                catText += (catText.Length > 0 ? "・" : "") + "柱";

            string summary = $"ビュー: {_data.ViewName}\n" +
                $"対象カテゴリ: {catText}\n\n";

            var selectedParam = ParameterComboBox.SelectedItem as FireProtectionParameterInfo;
            summary += $"耐火被覆パラメータ: {selectedParam?.ParameterName ?? "なし"}\n";
            summary += $"検出された種類: {_typeEntries.Count}\n";

            foreach (var entry in _typeEntries)
            {
                double offset = CommonOffsetRadio.IsChecked == true
                    ? ParseDouble(CommonOffsetInput.Text, 50)
                    : entry.OffsetMm;
                summary += $"  - {entry.Name}（オフセット: {offset}mm）\n";
            }

            var selectedLine = LineStyleComboBox.SelectedItem as LineStyleItem;
            var selectedPattern = FillPatternComboBox.SelectedItem as FillPatternItem;
            summary += $"\n境界線: {selectedLine?.Name ?? "なし"}\n";
            summary += $"塗りパターン: {selectedPattern?.Name ?? "なし"}";

            SummaryText.Text = summary;
        }

        #endregion

        #region Navigation

        private void ShowStep(int step)
        {
            _currentStep = step;

            Step1Panel.Visibility = step == 1
                ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            Step2Panel.Visibility = step == 2
                ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            Step3Panel.Visibility = step == 3
                ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

            BackButton.IsEnabled = step > 1;
            NextButton.Content = step == TotalSteps ? "実行" : "次へ";

            string[] stepNames = { "", "基本設定", "オフセット・表示設定", "処理確認" };
            StepIndicator.Text = $"ステップ {step} / {TotalSteps}  {stepNames[step]}";
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentStep == 1)
            {
                if (IncludeBeamsCheck.IsChecked != true &&
                    IncludeColumnsCheck.IsChecked != true)
                {
                    MessageBox.Show("対象カテゴリを1つ以上選択してください。",
                        "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (_typeEntries.Count == 0)
                {
                    MessageBox.Show("耐火被覆パラメータに値が設定された要素がありません。\n" +
                        "パラメータを確認してください。",
                        "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                InitializeStep2();
                ShowStep(2);
            }
            else if (_currentStep == 2)
            {
                InitializeStep3();
                ShowStep(3);
            }
            else if (_currentStep == 3)
            {
                DialogResult = true;
                Close();
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentStep > 1)
                ShowStep(_currentStep - 1);
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
            return double.TryParse(text, out val) ? val : defaultValue;
        }

        #endregion
    }
}

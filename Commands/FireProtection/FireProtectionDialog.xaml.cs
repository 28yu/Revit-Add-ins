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
                OverwriteExisting = OverwriteCheckBox.IsChecked == true,
                ColumnA_mm = ParseDouble(ColumnAInput.Text, 400),
                ColumnB_mm = ParseDouble(ColumnBInput.Text, 150)
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
            // 断面ビュー: 柱デフォルトON、それ以外: OFF
            IncludeColumnsCheck.IsChecked = _data.IsSectionView && _data.HasColumns;
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

            var beamColors = FilledRegionCreator.GenerateBeamColors(selected.UniqueValues.Count);
            var colColors = FilledRegionCreator.GenerateColumnColors(selected.UniqueValues.Count);
            _typeEntries = new List<FireProtectionTypeEntry>();

            for (int i = 0; i < selected.UniqueValues.Count; i++)
            {
                _typeEntries.Add(new FireProtectionTypeEntry
                {
                    Name = selected.UniqueValues[i],
                    OffsetMm = ParseDouble(CommonOffsetInput.Text, 50),
                    ColorR = beamColors[i].Red,
                    ColorG = beamColors[i].Green,
                    ColorB = beamColors[i].Blue,
                    ColColorR = colColors[i].Red,
                    ColColorG = colColors[i].Green,
                    ColColorB = colColors[i].Blue,
                    ColBgColorR = 255,
                    ColBgColorG = 255,
                    ColBgColorB = 255
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

                // 1行: 種類名 + 前景 + 背景 を横並び
                var rowPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, i > 0 ? 6 : 0, 0, 0)
                };

                // 種類名
                rowPanel.Children.Add(new TextBlock
                {
                    Text = entry.Name,
                    FontSize = 12,
                    FontWeight = FontWeights.Bold,
                    Width = 140,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = entry.Name,
                    Margin = new Thickness(4, 0, 8, 0)
                });

                // 前景
                int fgIdx = idx;
                var fgCheck = new CheckBox
                {
                    Content = "前景",
                    IsChecked = entry.ForegroundVisible,
                    FontSize = 11,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 4, 0)
                };
                fgCheck.Checked += (s, ev) => { _typeEntries[fgIdx].ForegroundVisible = true; };
                fgCheck.Unchecked += (s, ev) => { _typeEntries[fgIdx].ForegroundVisible = false; };
                rowPanel.Children.Add(fgCheck);

                var fgRect = new WpfRectangle
                {
                    Width = 32, Height = 20,
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(
                        entry.ColorR, entry.ColorG, entry.ColorB)),
                    Stroke = new SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0x66, 0x66, 0x66)),
                    StrokeThickness = 1,
                    Margin = new Thickness(0, 0, 16, 0),
                    Cursor = Cursors.Hand
                };
                fgRect.MouseLeftButtonDown += (s, ev) => { ShowColorPicker(fgIdx); };
                rowPanel.Children.Add(fgRect);

                // 背景
                int bgIdx = idx;
                var bgCheck = new CheckBox
                {
                    Content = "背景",
                    IsChecked = entry.BackgroundVisible,
                    FontSize = 11,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 4, 0)
                };
                bgCheck.Checked += (s, ev) => { _typeEntries[bgIdx].BackgroundVisible = true; };
                bgCheck.Unchecked += (s, ev) => { _typeEntries[bgIdx].BackgroundVisible = false; };
                rowPanel.Children.Add(bgCheck);

                var bgRect = new WpfRectangle
                {
                    Width = 32, Height = 20,
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(
                        entry.BgColorR, entry.BgColorG, entry.BgColorB)),
                    Stroke = new SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0x66, 0x66, 0x66)),
                    StrokeThickness = 1,
                    Cursor = Cursors.Hand
                };
                bgRect.MouseLeftButtonDown += (s, ev) => { ShowBgColorPicker(bgIdx); };
                rowPanel.Children.Add(bgRect);

                // 区切り + 柱色（平面/天伏ビューのみ）
                if (!_data.IsSectionView)
                {
                    int colIdx = idx;

                    // 区切り線
                    rowPanel.Children.Add(new WpfRectangle
                    {
                        Width = 1, Height = 18,
                        Fill = new SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(0xCC, 0xCC, 0xCC)),
                        Margin = new Thickness(12, 0, 12, 0),
                        VerticalAlignment = VerticalAlignment.Center
                    });

                    rowPanel.Children.Add(new TextBlock
                    {
                        Text = "柱:",
                        FontSize = 11, FontWeight = FontWeights.Bold,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(0, 0, 6, 0)
                    });

                    // 柱前景
                    var colFgCheck = new CheckBox
                    {
                        Content = "前景",
                        IsChecked = entry.ColForegroundVisible,
                        FontSize = 10,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(0, 0, 3, 0)
                    };
                    colFgCheck.Checked += (s, ev) => { _typeEntries[colIdx].ColForegroundVisible = true; };
                    colFgCheck.Unchecked += (s, ev) => { _typeEntries[colIdx].ColForegroundVisible = false; };
                    rowPanel.Children.Add(colFgCheck);

                    var colFgRect = new WpfRectangle
                    {
                        Width = 28, Height = 18,
                        Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(
                            entry.ColColorR, entry.ColColorG, entry.ColColorB)),
                        Stroke = new SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(0x66, 0x66, 0x66)),
                        StrokeThickness = 1,
                        Margin = new Thickness(0, 0, 8, 0),
                        Cursor = Cursors.Hand
                    };
                    colFgRect.MouseLeftButtonDown += (s, ev) => { ShowColColorPicker(colIdx); };
                    rowPanel.Children.Add(colFgRect);

                    // 柱背景
                    var colBgCheck = new CheckBox
                    {
                        Content = "背景",
                        IsChecked = entry.ColBackgroundVisible,
                        FontSize = 10,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(0, 0, 3, 0)
                    };
                    colBgCheck.Checked += (s, ev) => { _typeEntries[colIdx].ColBackgroundVisible = true; };
                    colBgCheck.Unchecked += (s, ev) => { _typeEntries[colIdx].ColBackgroundVisible = false; };
                    rowPanel.Children.Add(colBgCheck);

                    var colBgRect = new WpfRectangle
                    {
                        Width = 28, Height = 18,
                        Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(
                            entry.ColBgColorR, entry.ColBgColorG, entry.ColBgColorB)),
                        Stroke = new SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(0x66, 0x66, 0x66)),
                        StrokeThickness = 1,
                        Cursor = Cursors.Hand
                    };
                    colBgRect.MouseLeftButtonDown += (s, ev) => { ShowColBgColorPicker(colIdx); };
                    rowPanel.Children.Add(colBgRect);
                }

                DetectedTypesPanel.Children.Add(rowPanel);
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
            // 断面ビューでは柱設定を非表示
            ColumnSettingsPanel.Visibility = _data.IsSectionView
                ? System.Windows.Visibility.Collapsed
                : System.Windows.Visibility.Visible;
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
            var dlg = new WinForms.ColorDialog { FullOpen = true,
                Color = System.Drawing.Color.FromArgb(entry.ColorR, entry.ColorG, entry.ColorB) };
            if (dlg.ShowDialog() == WinForms.DialogResult.OK)
            { entry.ColorR = dlg.Color.R; entry.ColorG = dlg.Color.G; entry.ColorB = dlg.Color.B; RefreshDetectedTypesUI(); }
        }

        private void ShowBgColorPicker(int typeIndex)
        {
            if (typeIndex < 0 || typeIndex >= _typeEntries.Count) return;
            var entry = _typeEntries[typeIndex];
            var dlg = new WinForms.ColorDialog { FullOpen = true,
                Color = System.Drawing.Color.FromArgb(entry.BgColorR, entry.BgColorG, entry.BgColorB) };
            if (dlg.ShowDialog() == WinForms.DialogResult.OK)
            { entry.BgColorR = dlg.Color.R; entry.BgColorG = dlg.Color.G; entry.BgColorB = dlg.Color.B; RefreshDetectedTypesUI(); }
        }

        private void ShowColColorPicker(int typeIndex)
        {
            if (typeIndex < 0 || typeIndex >= _typeEntries.Count) return;
            var entry = _typeEntries[typeIndex];
            var dlg = new WinForms.ColorDialog { FullOpen = true,
                Color = System.Drawing.Color.FromArgb(entry.ColColorR, entry.ColColorG, entry.ColColorB) };
            if (dlg.ShowDialog() == WinForms.DialogResult.OK)
            { entry.ColColorR = dlg.Color.R; entry.ColColorG = dlg.Color.G; entry.ColColorB = dlg.Color.B; RefreshDetectedTypesUI(); }
        }

        private void ShowColBgColorPicker(int typeIndex)
        {
            if (typeIndex < 0 || typeIndex >= _typeEntries.Count) return;
            var entry = _typeEntries[typeIndex];
            var dlg = new WinForms.ColorDialog { FullOpen = true,
                Color = System.Drawing.Color.FromArgb(entry.ColBgColorR, entry.ColBgColorG, entry.ColBgColorB) };
            if (dlg.ShowDialog() == WinForms.DialogResult.OK)
            { entry.ColBgColorR = dlg.Color.R; entry.ColBgColorG = dlg.Color.G; entry.ColBgColorB = dlg.Color.B; RefreshDetectedTypesUI(); }
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

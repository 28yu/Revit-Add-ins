using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
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

        private void ComboPattern_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (CanvasPreview == null || ComboPattern.SelectedItem == null)
                return;

            var selectedType = ComboPattern.SelectedItem as FilledRegionType;
            if (selectedType == null)
                return;

            UpdatePatternPreview(selectedType);
        }

        private void UpdatePatternPreview(FilledRegionType patternType)
        {
            if (BorderPreview == null || CanvasPreview == null)
                return;

            // Canvasをクリア
            CanvasPreview.Children.Clear();

            try
            {
                // FillPatternElementを取得
                var patternId = patternType.ForegroundPatternId;
                if (patternId == null || patternId == ElementId.InvalidElementId)
                {
                    patternId = patternType.BackgroundPatternId;
                }

                if (patternId == null || patternId == ElementId.InvalidElementId)
                {
                    ShowPatternName(patternType.Name);
                    return;
                }

                var doc = patternType.Document;
                var fillPatternElement = doc.GetElement(patternId) as FillPatternElement;
                if (fillPatternElement == null)
                {
                    ShowPatternName(patternType.Name);
                    return;
                }

                var fillPattern = fillPatternElement.GetFillPattern();
                if (fillPattern == null)
                {
                    ShowPatternName(patternType.Name);
                    return;
                }

                // パターンを描画
                DrawFillPattern(fillPattern, patternType.Name);
            }
            catch
            {
                // エラー時はパターン名を表示
                ShowPatternName(patternType.Name);
            }
        }

        private void ShowPatternName(string patternName)
        {
            var textBlock = new System.Windows.Controls.TextBlock
            {
                Text = patternName,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 12,
                Foreground = System.Windows.Media.Brushes.Gray
            };

            CanvasPreview.Children.Add(textBlock);
            System.Windows.Controls.Canvas.SetLeft(textBlock, 10);
            System.Windows.Controls.Canvas.SetTop(textBlock, 30);
        }

        private void DrawFillPattern(FillPattern fillPattern, string patternName)
        {
            var width = BorderPreview.ActualWidth > 0 ? BorderPreview.ActualWidth : 420;
            var height = BorderPreview.ActualHeight > 0 ? BorderPreview.ActualHeight : 80;

            // 実線パターンの場合
            if (fillPattern.IsSolidFill)
            {
                var rect = new System.Windows.Shapes.Rectangle
                {
                    Width = width - 10,
                    Height = height - 10,
                    Fill = System.Windows.Media.Brushes.Black,
                    Stroke = System.Windows.Media.Brushes.Gray,
                    StrokeThickness = 1
                };
                CanvasPreview.Children.Add(rect);
                System.Windows.Controls.Canvas.SetLeft(rect, 5);
                System.Windows.Controls.Canvas.SetTop(rect, 5);
                return;
            }

            // ハッチングパターンの場合
            var gridSegments = fillPattern.GetFillGrids();
            if (gridSegments == null || gridSegments.Count == 0)
            {
                ShowPatternName(patternName);
                return;
            }

            // 背景を白で塗りつぶし
            var background = new System.Windows.Shapes.Rectangle
            {
                Width = width,
                Height = height,
                Fill = System.Windows.Media.Brushes.White
            };
            CanvasPreview.Children.Add(background);

            // 各グリッドセグメントを描画
            foreach (var grid in gridSegments)
            {
                DrawGridSegment(grid, width, height);
            }
        }

        private void DrawGridSegment(FillGrid grid, double canvasWidth, double canvasHeight)
        {
            // 角度を取得（ラジアンから度に変換）
            double angleRad = grid.Angle;
            double angleDeg = angleRad * 180.0 / Math.PI;

            // オフセットを取得（Revit単位からピクセルに変換、スケール調整）
            double offset = grid.Offset * 100; // スケール調整

            // 間隔を取得（デフォルト値を使用）
            double spacing = 15; // デフォルト間隔

            // ハッチング線を描画
            int lineCount = (int)(canvasHeight / spacing) + 10;
            for (int i = -5; i < lineCount; i++)
            {
                double y = i * spacing + offset;

                var line = new System.Windows.Shapes.Line
                {
                    X1 = -canvasWidth,
                    Y1 = y,
                    X2 = canvasWidth * 2,
                    Y2 = y,
                    Stroke = System.Windows.Media.Brushes.Black,
                    StrokeThickness = 0.5
                };

                // 回転変換を適用
                var rotateTransform = new RotateTransform(angleDeg, canvasWidth / 2, canvasHeight / 2);
                line.RenderTransform = rotateTransform;

                CanvasPreview.Children.Add(line);
            }
        }
    }
}

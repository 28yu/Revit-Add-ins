using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using Tools28.Commands.RoomTagCreator.Model;
using Tools28.Localization;

namespace Tools28.Commands.RoomTagCreator
{
    public partial class RoomTagUI : Window
    {
        private readonly Document _doc;
        private ObservableCollection<RoomInfo> _rooms;
        private List<RoomTagTypeInfo> _tagTypes;
        private int _rowCount = 2;

        // 結果プロパティ
        public string NewViewName { get; private set; }
        public string SelectedViewFamilyTypeName { get; private set; }
        public List<RoomInfo> SelectedRooms { get; private set; }
        public ElementId SelectedTagTypeId { get; private set; } = ElementId.InvalidElementId;
        public LayoutSettings Layout { get; private set; }

        public RoomTagUI(Document doc, List<RoomInfo> rooms, string sourceViewName)
        {
            InitializeComponent();

            _doc = doc;

            // 部屋リストをObservableCollectionに変換
            _rooms = new ObservableCollection<RoomInfo>(rooms);
            _rooms.CollectionChanged += (s, e) => UpdatePreview();
            RoomListBox.ItemsSource = _rooms;

            // ビュー名の自動生成
            ViewNameTextBox.Text = RoomTagService.GenerateViewName(sourceViewName);

            // ドロップダウンを初期化
            LoadTagTypes();

            // 初期表示
            RowCountText.Text = _rowCount.ToString();
            UpdatePreview();
        }

        private void LoadTagTypes()
        {
            _tagTypes = RoomTagService.GetRoomTagTypes(_doc);
            TagTypeComboBox.ItemsSource = _tagTypes.Select(t => t.DisplayName).ToList();
            if (_tagTypes.Count > 0)
                TagTypeComboBox.SelectedIndex = 0;
        }

        private void UpdatePreview()
        {
            if (PreviewCanvas == null || _rooms == null) return;

            PreviewCanvas.Children.Clear();

            int roomCount = _rooms.Count;
            if (roomCount == 0) return;

            int rows = _rowCount;
            int cols = (int)Math.Ceiling((double)roomCount / rows);

            double canvasW = PreviewCanvas.ActualWidth > 0 ? PreviewCanvas.ActualWidth : 620;
            double canvasH = PreviewCanvas.ActualHeight > 0 ? PreviewCanvas.ActualHeight : 120;

            // タグサイズとマージンの計算
            double margin = 4;
            double cellW = (canvasW - margin) / cols - margin;
            double cellH = (canvasH - margin) / rows - margin;

            // セルの最大・最小サイズ制限
            cellW = Math.Max(20, Math.Min(cellW, 120));
            cellH = Math.Max(14, Math.Min(cellH, 24));

            var tagBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xCC, 0xE5, 0xFF));
            var tagBorder = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x00, 0x66, 0xCC));
            var textBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x33, 0x33, 0x33));

            for (int i = 0; i < roomCount; i++)
            {
                // 横並び: 左→右、指定行数で折り返し
                int col = i / rows;
                int row = i % rows;

                double x = margin + col * (cellW + margin);
                double y = margin + row * (cellH + margin);

                var rect = new System.Windows.Shapes.Rectangle
                {
                    Width = cellW,
                    Height = cellH,
                    Fill = tagBrush,
                    Stroke = tagBorder,
                    StrokeThickness = 1,
                    RadiusX = 2,
                    RadiusY = 2
                };
                Canvas.SetLeft(rect, x);
                Canvas.SetTop(rect, y);
                PreviewCanvas.Children.Add(rect);

                // 部屋名テキスト
                string name = _rooms[i].Name;
                if (name.Length > 6) name = name.Substring(0, 5) + "..";
                var text = new TextBlock
                {
                    Text = name,
                    FontSize = 9,
                    Foreground = textBrush,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    MaxWidth = cellW - 4
                };
                Canvas.SetLeft(text, x + 3);
                Canvas.SetTop(text, y + (cellH - 14) / 2);
                PreviewCanvas.Children.Add(text);
            }
        }

        // --- イベントハンドラ ---

        private void RowCountUp_Click(object sender, RoutedEventArgs e)
        {
            if (_rowCount < _rooms.Count)
            {
                _rowCount++;
                RowCountText.Text = _rowCount.ToString();
                UpdatePreview();
            }
        }

        private void RowCountDown_Click(object sender, RoutedEventArgs e)
        {
            if (_rowCount > 1)
            {
                _rowCount--;
                RowCountText.Text = _rowCount.ToString();
                UpdatePreview();
            }
        }

        private void MoveUpButton_Click(object sender, RoutedEventArgs e)
        {
            int index = RoomListBox.SelectedIndex;
            if (index > 0)
            {
                var item = _rooms[index];
                _rooms.RemoveAt(index);
                _rooms.Insert(index - 1, item);
                RoomListBox.SelectedIndex = index - 1;
            }
        }

        private void MoveDownButton_Click(object sender, RoutedEventArgs e)
        {
            int index = RoomListBox.SelectedIndex;
            if (index >= 0 && index < _rooms.Count - 1)
            {
                var item = _rooms[index];
                _rooms.RemoveAt(index);
                _rooms.Insert(index + 1, item);
                RoomListBox.SelectedIndex = index + 1;
            }
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = RoomListBox.SelectedItems.Cast<RoomInfo>().ToList();
            foreach (var item in selectedItems)
            {
                _rooms.Remove(item);
            }
        }

        private void SpacingTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !Regex.IsMatch(e.Text, "^[0-9.]+$");
        }

        private void ExecuteButton_Click(object sender, RoutedEventArgs e)
        {
            // バリデーション
            if (string.IsNullOrWhiteSpace(ViewNameTextBox.Text))
            {
                MessageBox.Show(Loc.S("RoomTag.EnterViewName"), Loc.S("Common.Warning"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_rooms.Count == 0)
            {
                MessageBox.Show(Loc.S("RoomTag.SelectRooms"), Loc.S("Common.Warning"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (TagTypeComboBox.SelectedIndex < 0 || _tagTypes.Count == 0)
            {
                MessageBox.Show(Loc.S("RoomTag.SelectTagType"), Loc.S("Common.Warning"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(SpacingTextBox.Text, out double spacing) || spacing < 0)
            {
                MessageBox.Show(Loc.S("RoomTag.InvalidSpacing"), Loc.S("Common.InputError"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 結果を設定
            NewViewName = ViewNameTextBox.Text.Trim();
            SelectedRooms = _rooms.ToList();

            SelectedViewFamilyTypeName = FloorPlanRadio.IsChecked == true ? "仕上表_平面図" : "仕上表_天井伏図";
            SelectedTagTypeId = _tagTypes[TagTypeComboBox.SelectedIndex].Id;

            Layout = new LayoutSettings
            {
                Count = _rowCount,
                SpacingMm = spacing
            };

            DialogResult = true;
            Close();
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
    }
}

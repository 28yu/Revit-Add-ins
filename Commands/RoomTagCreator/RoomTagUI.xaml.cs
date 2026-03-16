using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Tools28.Commands.RoomTagCreator.Model;

namespace Tools28.Commands.RoomTagCreator
{
    public partial class RoomTagUI : Window
    {
        private readonly Document _doc;
        private ObservableCollection<RoomInfo> _rooms;
        private List<RoomTagTypeInfo> _tagTypes;
        private List<View> _viewTemplates;

        // 結果プロパティ
        public string NewViewName { get; private set; }
        public string SelectedViewFamilyTypeName { get; private set; }
        public List<RoomInfo> SelectedRooms { get; private set; }
        public ElementId SelectedTagTypeId { get; private set; } = ElementId.InvalidElementId;
        public LayoutSettings Layout { get; private set; }
        public ElementId SelectedViewTemplateId { get; private set; } = ElementId.InvalidElementId;

        public RoomTagUI(Document doc, List<RoomInfo> rooms, string sourceViewName)
        {
            InitializeComponent();

            _doc = doc;

            // 部屋リストをObservableCollectionに変換
            _rooms = new ObservableCollection<RoomInfo>(rooms);
            RoomListBox.ItemsSource = _rooms;

            // ビュー名の自動生成
            ViewNameTextBox.Text = RoomTagService.GenerateViewName(sourceViewName);

            // ドロップダウンを初期化
            LoadTagTypes();
            LoadViewTemplates();

            // 初期状態のラベル設定
            UpdateCountLabel();
        }

        private void LoadTagTypes()
        {
            _tagTypes = RoomTagService.GetRoomTagTypes(_doc);
            TagTypeComboBox.ItemsSource = _tagTypes.Select(t => t.DisplayName).ToList();
            if (_tagTypes.Count > 0)
                TagTypeComboBox.SelectedIndex = 0;
        }

        private void LoadViewTemplates()
        {
            _viewTemplates = RoomTagService.GetViewTemplates(_doc);
            var templateNames = new List<string> { "(なし)" };
            templateNames.AddRange(_viewTemplates.Select(v => v.Name));
            ViewTemplateComboBox.ItemsSource = templateNames;
            ViewTemplateComboBox.SelectedIndex = 0;
        }

        private void UpdateCountLabel()
        {
            if (CountLabel == null) return;

            if (HorizontalRadio != null && HorizontalRadio.IsChecked == true)
            {
                CountLabel.Text = "列数:";
            }
            else
            {
                CountLabel.Text = "行数:";
            }
        }

        // --- イベントハンドラ ---

        private void LayoutDirection_Changed(object sender, RoutedEventArgs e)
        {
            UpdateCountLabel();
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

        private void CountTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !Regex.IsMatch(e.Text, "^[0-9]+$");
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
                MessageBox.Show("ビュー名を入力してください。", "警告",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_rooms.Count == 0)
            {
                MessageBox.Show("部屋が選択されていません。", "警告",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (TagTypeComboBox.SelectedIndex < 0 || _tagTypes.Count == 0)
            {
                MessageBox.Show("タグファミリタイプを選択してください。", "警告",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(CountTextBox.Text, out int count) || count < 1)
            {
                MessageBox.Show("行数/列数には1以上の整数を入力してください。", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(SpacingTextBox.Text, out double spacing) || spacing < 0)
            {
                MessageBox.Show("タグ間隔には0以上の数値を入力してください。", "入力エラー",
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
                IsHorizontal = HorizontalRadio.IsChecked == true,
                Count = count,
                SpacingMm = spacing
            };

            // ビューテンプレート（0=なし、1以降=テンプレート）
            int templateIndex = ViewTemplateComboBox.SelectedIndex;
            if (templateIndex > 0)
                SelectedViewTemplateId = _viewTemplates[templateIndex - 1].Id;
            else
                SelectedViewTemplateId = ElementId.InvalidElementId;

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

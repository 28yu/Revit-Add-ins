using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;

namespace Tools28.Commands.SheetCreation
{
    /// <summary>
    /// 図枠表示用クラス
    /// </summary>
    public class TitleBlockItem
    {
        public FamilySymbol Symbol { get; set; }
        public string DisplayName { get; set; }

        public TitleBlockItem(FamilySymbol symbol)
        {
            Symbol = symbol;
            DisplayName = $"{symbol.FamilyName} : {symbol.Name}";
        }
    }

    /// <summary>
    /// シート一括作成ダイアログ
    /// </summary>
    public partial class SheetCreationDialog : Window
    {
        private readonly Document _doc;
        private List<TitleBlockItem> _allTitleBlocks;
        private static ElementId _lastUsedTitleBlockId = null;

        public TitleBlockItem SelectedTitleBlock { get; private set; }
        public int SheetCount { get; private set; }
        public string Prefix { get; private set; }

        public SheetCreationDialog(Document doc)
        {
            InitializeComponent();

            _doc = doc;
            SheetCount = 5;
            Prefix = "";

            LoadTitleBlocks();
            RestoreLastSettings();
        }

        private void LoadTitleBlocks()
        {
            var symbols = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .OrderBy(s => s.FamilyName)
                .ThenBy(s => s.Name)
                .ToList();

            if (symbols.Count == 0)
            {
                MessageBox.Show(
                    "プロジェクトに図枠ファミリがロードされていません。",
                    "警告",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                DialogResult = false;
                Close();
                return;
            }

            _allTitleBlocks = symbols.Select(s => new TitleBlockItem(s)).ToList();
            TitleBlockComboBox.ItemsSource = _allTitleBlocks;

            if (_allTitleBlocks.Count > 0)
            {
                TitleBlockComboBox.SelectedIndex = 0;
            }
        }

        private void RestoreLastSettings()
        {
            if (_lastUsedTitleBlockId != null && _allTitleBlocks != null)
            {
                var lastUsed = _allTitleBlocks.FirstOrDefault(tb => tb.Symbol.Id == _lastUsedTitleBlockId);
                if (lastUsed != null)
                {
                    TitleBlockComboBox.SelectedItem = lastUsed;
                }
            }
        }

        private void TitleBlockComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectedTitleBlock = TitleBlockComboBox.SelectedItem as TitleBlockItem;
        }

        private void SheetCountTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SheetCountTextBox.Text))
            {
                SheetCount = 0;
                return;
            }

            if (int.TryParse(SheetCountTextBox.Text, out int count))
            {
                if (count < 1)
                {
                    SheetCountTextBox.Text = "1";
                    SheetCountTextBox.SelectionStart = 1;
                    count = 1;
                }
                else if (count > 100)
                {
                    SheetCountTextBox.Text = "100";
                    SheetCountTextBox.SelectionStart = 3;
                    count = 100;
                }
                SheetCount = count;
            }
        }

        private void SheetCountTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !Regex.IsMatch(e.Text, "^[0-9]+$");
        }

        private void PrefixTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Prefix = PrefixTextBox.Text ?? "";
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedTitleBlock == null)
            {
                MessageBox.Show("図枠を選択してください。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (SheetCount < 1 || SheetCount > 100)
            {
                MessageBox.Show("作成枚数は1～100の範囲で入力してください。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _lastUsedTitleBlockId = SelectedTitleBlock.Symbol.Id;
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
            else if (e.Key == Key.Enter)
            {
                CreateButton_Click(null, null);
            }
        }
    }
}
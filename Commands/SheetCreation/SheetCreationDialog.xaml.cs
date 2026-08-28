using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Tools28.Localization;

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
    /// <summary>リスト入力モードで指定された1シート分の情報。</summary>
    public class SheetListEntry
    {
        /// <summary>シート番号（必須）</summary>
        public string Number { get; set; }

        /// <summary>シート名（空なら既定名を使う）</summary>
        public string Name { get; set; }
    }

    public partial class SheetCreationDialog : Window
    {
        private readonly Document _doc;
        private List<TitleBlockItem> _allTitleBlocks;
        private static ElementId _lastUsedTitleBlockId = null;

        public TitleBlockItem SelectedTitleBlock { get; private set; }
        public int SheetCount { get; private set; }
        public string Prefix { get; private set; }

        /// <summary>リスト入力モードか（false なら従来の「枚数＋図面No」モード）。</summary>
        public bool UseListMode { get; private set; }

        /// <summary>リスト入力モードで指定されたシート（番号・名前）。連番モードでは空。</summary>
        public List<SheetListEntry> SheetList { get; private set; } = new List<SheetListEntry>();

        public SheetCreationDialog(Document doc)
        {
            InitializeComponent();
            ApplyLocalization();

            _doc = doc;
            SheetCount = 5;
            Prefix = "";

            LoadTitleBlocks();
            RestoreLastSettings();
        }

        private void ApplyLocalization()
        {
            this.Title = Loc.S("Sheet.Title");
            txtSelectTitleBlock.Text = Loc.S("Sheet.SelectTitleBlock");
            txtCount.Text = Loc.S("Sheet.Count");
            txtCountHint.Text = Loc.S("Sheet.CountHint");
            txtDrawingNo.Text = Loc.S("Sheet.DrawingNo");
            txtDrawingNoHint.Text = Loc.S("Sheet.DrawingNoHint");
            txtMode.Text = Loc.S("Sheet.Mode");
            rbModeCount.Content = Loc.S("Sheet.Mode.Count");
            rbModeList.Content = Loc.S("Sheet.Mode.List");
            txtList.Text = Loc.S("Sheet.List");
            txtListHint.Text = Loc.S("Sheet.ListHint");
            CreateButton.Content = Loc.S("Common.Create");
            btnCancel.Content = Loc.S("Common.Cancel");
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
                    Loc.S("Sheet.NoTitleBlock"),
                    Loc.S("Common.Warning"),
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
                MessageBox.Show(Loc.S("Sheet.SelectTitleBlockWarn"), Loc.S("Common.Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (UseListMode)
            {
                SheetList = ParseSheetList(SheetListTextBox?.Text);
                if (SheetList.Count == 0)
                {
                    MessageBox.Show(Loc.S("Sheet.ListEmpty"), Loc.S("Common.Warning"),
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }
            else if (SheetCount < 1 || SheetCount > 100)
            {
                MessageBox.Show(Loc.S("Sheet.CountRange"), Loc.S("Common.Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _lastUsedTitleBlockId = SelectedTitleBlock.Symbol.Id;
            DialogResult = true;
            Close();
        }

        /// <summary>作成方法のラジオボタン切替。選んだ方の入力欄だけを表示する。</summary>
        private void Mode_Changed(object sender, RoutedEventArgs e)
        {
            // InitializeComponent 中にも Checked が飛ぶため、生成前は何もしない
            if (pnlCountMode == null || pnlListMode == null) return;

            UseListMode = rbModeList.IsChecked == true;

            // ⚠ Window 派生クラスは Visibility というインスタンスプロパティを継承するため、
            //   列挙型を単純名で参照すると CS0176 になる。必ず完全修飾する。
            pnlCountMode.Visibility = UseListMode
                ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
            pnlListMode.Visibility = UseListMode
                ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
        }

        private void SheetListTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            SheetList = ParseSheetList(SheetListTextBox?.Text);
        }

        /// <summary>
        /// 「シート番号[Tab]シート名」形式のテキストを解析する。
        /// - 空行は無視する
        /// - タブが無い行は「番号のみ」とみなし、名前は空にする
        /// - 3列目以降は無視する（Excel から余分な列ごと貼り付けても壊れない）
        /// - 入力内での番号重複は先勝ちで除く（既存シートとの重複は作成時に判定）
        /// </summary>
        internal static List<SheetListEntry> ParseSheetList(string text)
        {
            var list = new List<SheetListEntry>();
            if (string.IsNullOrWhiteSpace(text)) return list;

            var seen = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);

            foreach (var rawLine in text.Split('\n'))
            {
                string line = rawLine.TrimEnd('\r');
                if (string.IsNullOrWhiteSpace(line)) continue;

                string[] cols = line.Split('\t');
                string number = cols[0].Trim();
                if (number.Length == 0) continue;

                string name = cols.Length > 1 ? cols[1].Trim() : "";

                if (!seen.Add(number)) continue;   // 入力内の重複
                list.Add(new SheetListEntry { Number = number, Name = name });
            }

            return list;
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
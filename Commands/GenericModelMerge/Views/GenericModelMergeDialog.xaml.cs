using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Tools28.Commands.GenericModelMerge.Models;
using Tools28.Localization;

namespace Tools28.Commands.GenericModelMerge.Views
{
    /// <summary>
    /// 一般モデル化の実行条件を決めるダイアログ。
    /// 出力形式・まとめ方の各選択肢には、専門用語を知らなくても選べるよう
    /// 必ず説明文（*.Hint）を添える。
    /// </summary>
    public partial class GenericModelMergeDialog : Window
    {
        private readonly List<MergeCategoryRow> _categories;
        private readonly string _defaultDirectory;

        /// <summary>
        /// 既定名に使う言語コード。作成する要素／ファミリの名前は**モデルに保存される**ため、
        /// アドインの言語設定ではなく Revit 本体の UI 言語に合わせる（CLAUDE.md「文字列の3分類」B）。
        /// </summary>
        private readonly string _modelLang;

        /// <summary>OK で閉じたときの実行条件。</summary>
        internal MergeOptions Options { get; private set; }

        /// <summary>OK で閉じたときの対象要素 Id。</summary>
        internal List<ElementId> TargetElementIds { get; private set; } = new List<ElementId>();

        internal GenericModelMergeDialog(
            List<MergeCategoryRow> categories,
            List<MaterialRow> materials,
            string defaultDirectory,
            string modelLang)
        {
            _categories = categories ?? new List<MergeCategoryRow>();
            _defaultDirectory = defaultDirectory;
            _modelLang = modelLang;

            InitializeComponent();
            ApplyLocalization();

            lstCategories.ItemsSource = _categories;
            foreach (var row in _categories)
                row.PropertyChanged += Category_PropertyChanged;

            cmbMaterial.ItemsSource = materials;
            if (materials != null && materials.Count > 0)
                cmbMaterial.SelectedIndex = 0;

            txtName.Text = $"{Loc.S("GmMerge.DefaultName", _modelLang)}_{DateTime.Now:yyyyMMdd_HHmm}";

            UpdateSummary();
        }

        private void ApplyLocalization()
        {
            Title = Loc.S("GmMerge.Title");
            txtDescription.Text = Loc.S("GmMerge.Description");

            grpCategories.Header = Loc.S("GmMerge.Categories");
            txtCategoryHint.Text = Loc.S("GmMerge.Categories.Hint");
            btnSelectAll.Content = Loc.S("GmMerge.SelectAll");
            btnDeselectAll.Content = Loc.S("GmMerge.DeselectAll");

            grpOutput.Header = Loc.S("GmMerge.Output");
            rbDirectShape.Content = Loc.S("GmMerge.Output.DirectShape");
            txtDirectShapeHint.Text = Loc.S("GmMerge.Output.DirectShape.Hint");
            rbFamily.Content = Loc.S("GmMerge.Output.Family");
            txtFamilyHint.Text = Loc.S("GmMerge.Output.Family.Hint");

            grpCombine.Header = Loc.S("GmMerge.Combine");
            rbKeepShapes.Content = Loc.S("GmMerge.Combine.Keep");
            txtKeepShapesHint.Text = Loc.S("GmMerge.Combine.Keep.Hint");
            rbUnionAll.Content = Loc.S("GmMerge.Combine.UnionAll");
            txtUnionAllHint.Text = Loc.S("GmMerge.Combine.UnionAll.Hint");
            rbUnionTouching.Content = Loc.S("GmMerge.Combine.UnionTouching");
            txtUnionTouchingHint.Text = Loc.S("GmMerge.Combine.UnionTouching.Hint");

            grpResult.Header = Loc.S("GmMerge.Detail");
            lblMaterial.Text = Loc.S("GmMerge.Material");
            txtMaterialHint.Text = Loc.S("GmMerge.Material.Hint");
            lblName.Text = Loc.S("GmMerge.Name");
            txtNameHint.Text = Loc.S("GmMerge.Name.Hint");
            chkHideSource.Content = Loc.S("GmMerge.HideSource");
            txtHideSourceHint.Text = Loc.S("GmMerge.HideSource.Hint");

            btnRun.Content = Loc.S("GmMerge.Btn.Run");
            btnCancel.Content = Loc.S("Common.Cancel");
        }

        private void Category_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(MergeCategoryRow.IsSelected))
                UpdateSummary();
        }

        private void UpdateSummary()
        {
            int categoryCount = _categories.Count(c => c.IsSelected);
            int elementCount = _categories.Where(c => c.IsSelected).Sum(c => c.ElementCount);

            txtSummary.Text = string.Format(
                Loc.S("GmMerge.Summary"), categoryCount, elementCount);
            btnRun.IsEnabled = elementCount > 0;
        }

        private void btnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var row in _categories) row.IsSelected = true;
        }

        private void btnDeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var row in _categories) row.IsSelected = false;
        }

        private void OutputKind_Changed(object sender, RoutedEventArgs e)
        {
            // 名前の説明文は出力形式で意味が変わる（要素名 / ファミリ名）
            if (txtNameHint == null) return;
            txtNameHint.Text = rbFamily.IsChecked == true
                ? Loc.S("GmMerge.Name.Hint.Family")
                : Loc.S("GmMerge.Name.Hint");
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            var selected = _categories.Where(c => c.IsSelected).ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show(this, Loc.S("GmMerge.Err.NoCategory"),
                    Loc.S("GmMerge.Title"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string name = (txtName.Text ?? "").Trim();
            if (name.Length == 0)
            {
                MessageBox.Show(this, Loc.S("GmMerge.Err.NoName"),
                    Loc.S("GmMerge.Title"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var options = new MergeOptions
            {
                OutputKind = rbFamily.IsChecked == true
                    ? MergeOutputKind.Family
                    : MergeOutputKind.DirectShape,
                CombineMode =
                    rbUnionAll.IsChecked == true ? MergeCombineMode.UnionAll :
                    rbUnionTouching.IsChecked == true ? MergeCombineMode.UnionTouching :
                    MergeCombineMode.KeepShapes,
                MaterialId = (cmbMaterial.SelectedItem as MaterialRow)?.Id ?? ElementId.InvalidElementId,
                HideSourceElements = chkHideSource.IsChecked == true,
                Name = name,
            };

            if (options.OutputKind == MergeOutputKind.Family)
            {
                string path = AskFamilyPath(name);
                if (string.IsNullOrEmpty(path)) return;   // 保存先ダイアログでキャンセル
                options.FamilyPath = path;
            }

            Options = options;
            TargetElementIds = selected.SelectMany(c => c.ElementIds).ToList();
            DialogResult = true;   // DialogResult を設定すると ShowDialog が閉じる
        }

        /// <summary>.rfa の保存先をユーザーに決めてもらう。</summary>
        private string AskFamilyPath(string suggestedName)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = Loc.S("GmMerge.SaveFamily.Title"),
                Filter = Loc.S("GmMerge.SaveFamily.Filter") + " (*.rfa)|*.rfa",
                DefaultExt = ".rfa",
                FileName = SanitizeFileName(suggestedName) + ".rfa",
                OverwritePrompt = true,
            };

            if (!string.IsNullOrEmpty(_defaultDirectory) && Directory.Exists(_defaultDirectory))
                dlg.InitialDirectory = _defaultDirectory;

            return dlg.ShowDialog(this) == true ? dlg.FileName : null;
        }

        private static string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}

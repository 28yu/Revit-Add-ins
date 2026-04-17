using System.Windows;
using Tools28.Commands.ExcelExportImport.Models;
using Tools28.Localization;

namespace Tools28.Commands.ExcelExportImport.Views
{
    /// <summary>
    /// エクスポート範囲選択ダイアログ
    /// </summary>
    public partial class ScopeSelectionDialog : Window
    {
        /// <summary>ユーザーが選択した範囲</summary>
        public ExportScope SelectedScope { get; private set; } = ExportScope.EntireProject;

        public ScopeSelectionDialog(bool hasActiveView, bool hasSelection)
        {
            InitializeComponent();
            ApplyLocalization();

            btnView.IsEnabled = hasActiveView;
            btnSelection.IsEnabled = hasSelection;
        }

        private void ApplyLocalization()
        {
            this.Title = Loc.S("Export.Scope.WindowTitle");
            TitleText.Text = Loc.S("Export.Scope.Title");
            DescriptionText.Text = Loc.S("Export.Scope.Description");
            ProjectText.Text = "\u2192  " + Loc.S("Export.Scope.Project");
            ViewText.Text = "\u2192  " + Loc.S("Export.Scope.View");
            SelectionText.Text = "\u2192  " + Loc.S("Export.Scope.Selection");
        }

        private void btnProject_Click(object sender, RoutedEventArgs e)
        {
            SelectedScope = ExportScope.EntireProject;
            DialogResult = true;
            Close();
        }

        private void btnView_Click(object sender, RoutedEventArgs e)
        {
            SelectedScope = ExportScope.ActiveView;
            DialogResult = true;
            Close();
        }

        private void btnSelection_Click(object sender, RoutedEventArgs e)
        {
            SelectedScope = ExportScope.Selection;
            DialogResult = true;
            Close();
        }
    }
}

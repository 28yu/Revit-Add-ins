using System.Collections.Generic;
using System.Windows;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Engine;
using Tools28.Commands.FormworkCalculator.Models;
using Tools28.Localization;

namespace Tools28.Commands.FormworkCalculator.Views
{
    public partial class FormworkDialog : Window
    {
        public FormworkSettings Settings { get; private set; }

        private readonly Document _doc;

        public FormworkDialog(Document doc = null, FormworkSettings defaults = null)
        {
            _doc = doc;
            InitializeComponent();
            ApplyLocalization();
            PopulateParameterCandidates();
            Load(defaults ?? new FormworkSettings());
        }

        private void ApplyLocalization()
        {
            this.Title = Loc.S("Formwork.Title");

            grpScope.Header = Loc.S("Formwork.Scope.Header");
            txtScopeProject.Text = Loc.S("Formwork.Scope.Project");
            txtScopeView.Text = Loc.S("Formwork.Scope.View");

            grpGrouping.Header = Loc.S("Formwork.Group.Header");
            txtByCategory.Text = Loc.S("Formwork.Group.Category");
            txtByZone.Text = Loc.S("Formwork.Group.Zone");
            txtByType.Text = Loc.S("Formwork.Group.Type");
            lblZoneParam.Text = Loc.S("Formwork.Group.ParamName");
            lblTypeParam.Text = Loc.S("Formwork.Group.ParamName");

            grpOutput.Header = Loc.S("Formwork.Output.Header");
            txtExportExcel.Text = Loc.S("Formwork.Output.Excel");
            txtCreateSchedule.Text = Loc.S("Formwork.Output.Schedule");
            txtCreate3DView.Text = Loc.S("Formwork.Output.View3D");
            txtCreateSheet.Text = Loc.S("Formwork.Output.Sheet");

            grpColor.Header = Loc.S("Formwork.Color.Header");
            txtColorCat.Text = Loc.S("Formwork.Color.Category");
            txtColorZone.Text = Loc.S("Formwork.Color.Zone");
            txtColorType.Text = Loc.S("Formwork.Color.Type");

            grpOption.Header = Loc.S("Formwork.Option.Header");
            txtShowDeducted.Text = Loc.S("Formwork.Option.ShowDeducted");
            txtUseGL.Text = Loc.S("Formwork.Option.UseGL");

            btnOK.Content = Loc.S("Common.Execute");
            btnCancel.Content = Loc.S("Common.Cancel");
        }

        /// <summary>
        /// 工区別・型枠種別のパラメータ候補を ComboBox に投入する。
        /// 名前にキーワード (工区/Zone/型枠/Type 等) を含むパラメータを自動検出。
        /// </summary>
        private void PopulateParameterCandidates()
        {
            if (_doc == null) return;
            try
            {
                List<string> zoneCands = ParameterCandidateScanner.FindCandidates(
                    _doc, ParameterCandidateScanner.ZoneKeywords);
                CmbZoneParam.ItemsSource = zoneCands;
            }
            catch { }
            try
            {
                List<string> typeCands = ParameterCandidateScanner.FindCandidates(
                    _doc, ParameterCandidateScanner.FormworkTypeKeywords);
                CmbTypeParam.ItemsSource = typeCands;
            }
            catch { }
        }

        private void Load(FormworkSettings s)
        {
            if (s.Scope == CalculationScope.EntireProject)
                RadioEntireProject.IsChecked = true;
            else
                RadioCurrentView.IsChecked = true;

            ChkByCategory.IsChecked = s.GroupByCategory;
            ChkByZone.IsChecked = s.GroupByZone;
            ChkByType.IsChecked = s.GroupByFormworkType;
            CmbZoneParam.Text = s.ZoneParameterName ?? string.Empty;
            CmbTypeParam.Text = s.FormworkTypeParameterName ?? string.Empty;

            ChkExportExcel.IsChecked = s.ExportToExcel;
            ChkCreateSchedule.IsChecked = s.CreateSchedule;
            ChkCreate3DView.IsChecked = s.Create3DView;
            ChkCreateSheet.IsChecked = s.CreateSheet;

            switch (s.ColorScheme)
            {
                case ColorSchemeType.ByZone: RadioColorZone.IsChecked = true; break;
                case ColorSchemeType.ByFormworkType: RadioColorType.IsChecked = true; break;
                default: RadioColorCategory.IsChecked = true; break;
            }

            ChkShowDeducted.IsChecked = s.ShowDeductedFaces;
            ChkUseGL.IsChecked = s.UseGLDeduction;
            TxtGL.Text = s.GLElevationMeters.ToString("F3");
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            var s = new FormworkSettings
            {
                Scope = RadioEntireProject.IsChecked == true
                    ? CalculationScope.EntireProject : CalculationScope.CurrentView,
                GroupByCategory = ChkByCategory.IsChecked == true,
                GroupByZone = ChkByZone.IsChecked == true,
                ZoneParameterName = CmbZoneParam.Text?.Trim() ?? string.Empty,
                GroupByFormworkType = ChkByType.IsChecked == true,
                FormworkTypeParameterName = CmbTypeParam.Text?.Trim() ?? string.Empty,
                ExportToExcel = ChkExportExcel.IsChecked == true,
                CreateSchedule = ChkCreateSchedule.IsChecked == true,
                Create3DView = ChkCreate3DView.IsChecked == true,
                CreateSheet = ChkCreateSheet.IsChecked == true,
                ShowDeductedFaces = ChkShowDeducted.IsChecked == true,
                UseGLDeduction = ChkUseGL.IsChecked == true,
            };

            if (RadioColorZone.IsChecked == true) s.ColorScheme = ColorSchemeType.ByZone;
            else if (RadioColorType.IsChecked == true) s.ColorScheme = ColorSchemeType.ByFormworkType;
            else s.ColorScheme = ColorSchemeType.ByCategory;

            if (s.GroupByZone && string.IsNullOrWhiteSpace(s.ZoneParameterName))
            {
                MessageBox.Show(Loc.S("Formwork.NeedZoneParam"),
                    Loc.S("Common.InputError"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (s.GroupByFormworkType && string.IsNullOrWhiteSpace(s.FormworkTypeParameterName))
            {
                MessageBox.Show(Loc.S("Formwork.NeedTypeParam"),
                    Loc.S("Common.InputError"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (s.UseGLDeduction)
            {
                if (!double.TryParse(TxtGL.Text?.Trim(), out double gl))
                {
                    MessageBox.Show(Loc.S("Formwork.BadGL"),
                        Loc.S("Common.InputError"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                s.GLElevationMeters = gl;
            }

            Settings = s;
            DialogResult = true;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Localization;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using WinForms = System.Windows.Forms;

namespace Tools28.Commands.FillPatternIO
{
    public partial class FillPatternIODialog : Window
    {
        private readonly Document _doc;
        private List<FillPatternItem> _all = new List<FillPatternItem>();

        public FillPatternIODialog(UIDocument uidoc)
        {
            InitializeComponent();
            _doc = uidoc.Document;
            ApplyLocalization();
            LoadPatterns();
        }

        private bool UseMm => ComboUnit.SelectedIndex != 1;

        private void ApplyLocalization()
        {
            this.Title = Loc.S("FillPatternIO.Title");
            grpFilter.Header = Loc.S("FillPatternIO.PatternType");
            RadioModel.Content = Loc.S("FillPatternIO.Model");
            RadioDrafting.Content = Loc.S("FillPatternIO.Drafting");
            btnSelectAll.Content = Loc.S("FillPatternIO.SelectAll");
            btnClearAll.Content = Loc.S("FillPatternIO.ClearAll");
            colName.Header = Loc.S("FillPatternIO.ColName");
            colPreview.Header = Loc.S("FillPatternIO.ColPreview");
            colType.Header = Loc.S("FillPatternIO.ColType");
            colGrid.Header = Loc.S("FillPatternIO.ColGrid");
            txtUnitLabel.Text = Loc.S("FillPatternIO.Unit");
            btnExport.Content = Loc.S("FillPatternIO.Export");
            btnImport.Content = Loc.S("FillPatternIO.Import");
            btnClose.Content = Loc.S("FillPatternIO.Close");
        }

        private void LoadPatterns()
        {
            _all = new FilteredElementCollector(_doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .Select(e => new FillPatternItem(e))
                .OrderBy(i => i.Target.ToString())
                .ThenBy(i => i.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (ListPatterns == null) return;

            // 表示は「モデル」「製図」のいずれか。既定は製図。
            var target = RadioModel?.IsChecked == true
                ? FillPatternTarget.Model
                : FillPatternTarget.Drafting;

            ListPatterns.ItemsSource = _all.Where(i => i.Target == target).ToList();
            UpdateCount();
        }

        private void UpdateCount()
        {
            if (txtCount == null) return;
            int total = _all.Count;
            int selected = _all.Count(i => i.IsSelected);
            txtCount.Text = string.Format(Loc.S("FillPatternIO.CountInfo"), total, selected);
        }

        private void Filter_Changed(object sender, RoutedEventArgs e) => ApplyFilter();

        private void SelectAll_Click(object sender, RoutedEventArgs e)
            => SetSelectionForVisible(true);

        private void ClearAll_Click(object sender, RoutedEventArgs e)
            => SetSelectionForVisible(false);

        private void SetSelectionForVisible(bool value)
        {
            if (ListPatterns.ItemsSource is IEnumerable<FillPatternItem> visible)
            {
                foreach (var item in visible)
                    item.IsSelected = value;
            }
            UpdateCount();
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            var selected = _all.Where(i => i.IsSelected).ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show(this, Loc.S("FillPatternIO.NoSelection"),
                    Loc.S("FillPatternIO.Title"), MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // ソリッド塗り潰しはグリッドを持たず .pat で表現できないため除外
            var exportable = selected.Where(i => !i.IsSolid && i.GridCount > 0).ToList();
            int solidSkipped = selected.Count - exportable.Count;

            if (exportable.Count == 0)
            {
                MessageBox.Show(this, Loc.S("FillPatternIO.OnlySolid"),
                    Loc.S("FillPatternIO.Title"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool useMm = UseMm;
            int written = 0;

            try
            {
                if (exportable.Count == 1)
                {
                    var sfd = new SaveFileDialog
                    {
                        Filter = Loc.S("FillPatternIO.PatFilter"),
                        DefaultExt = ".pat",
                        FileName = PatFile.SanitizeFileName(exportable[0].Name)
                    };
                    if (sfd.ShowDialog(this) != true) return;

                    var fpe = _doc.GetElement(exportable[0].Id) as FillPatternElement;
                    PatFile.ExportToFile(fpe, sfd.FileName, useMm);
                    written = 1;
                }
                else
                {
                    using (var fbd = new WinForms.FolderBrowserDialog
                    {
                        Description = Loc.S("FillPatternIO.SelectFolder")
                    })
                    {
                        if (fbd.ShowDialog() != WinForms.DialogResult.OK) return;

                        string dir = fbd.SelectedPath;
                        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var item in exportable)
                        {
                            var fpe = _doc.GetElement(item.Id) as FillPatternElement;
                            string baseName = PatFile.SanitizeFileName(item.Name);
                            string fileName = baseName;
                            int dup = 2;
                            while (!usedNames.Add(fileName))
                                fileName = $"{baseName}_{dup++}";
                            string file = Path.Combine(dir, fileName + ".pat");
                            PatFile.ExportToFile(fpe, file, useMm);
                            written++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, string.Format(Loc.S("FillPatternIO.ExportError"), ex.Message),
                    Loc.S("Common.Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string msg = string.Format(Loc.S("FillPatternIO.ExportDone"), written);
            if (solidSkipped > 0)
                msg += "\n" + string.Format(Loc.S("FillPatternIO.SolidSkipped"), solidSkipped);

            MessageBox.Show(this, msg, Loc.S("FillPatternIO.Title"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = Loc.S("FillPatternIO.PatFilter"),
                DefaultExt = ".pat",
                Multiselect = true
            };
            if (ofd.ShowDialog(this) != true) return;

            int created = 0, skipped = 0, invalid = 0;

            try
            {
                using (var trans = new Transaction(_doc, Loc.S("FillPatternIO.TransImport")))
                {
                    trans.Start();
                    foreach (var path in ofd.FileNames)
                    {
                        var patterns = PatFile.Parse(path, UseMm);
                        created += PatFile.CreatePatterns(_doc, patterns, out int s, out int inv);
                        skipped += s;
                        invalid += inv;
                    }
                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, string.Format(Loc.S("FillPatternIO.ImportError"), ex.Message),
                    Loc.S("Common.Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            LoadPatterns();

            string msg = string.Format(Loc.S("FillPatternIO.ImportDone"), created);
            if (skipped > 0)
                msg += "\n" + string.Format(Loc.S("FillPatternIO.DupSkipped"), skipped);
            if (invalid > 0)
                msg += "\n" + string.Format(Loc.S("FillPatternIO.InvalidSkipped"), invalid);

            MessageBox.Show(this, msg, Loc.S("FillPatternIO.Title"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

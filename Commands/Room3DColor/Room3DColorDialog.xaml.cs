using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WpfRectangle = System.Windows.Shapes.Rectangle;
using WinForms = System.Windows.Forms;
using RevitColor = Autodesk.Revit.DB.Color;
using Tools28.Localization;

namespace Tools28.Commands.Room3DColor
{
    public partial class Room3DColorDialog : Window
    {
        private readonly List<RoomEntry> _entries;
        private readonly List<string> _parameterNames;
        private List<RoomColorGroup> _groups = new List<RoomColorGroup>();
        private bool _initialized;

        private class BasisItem
        {
            public RoomColorBasis Basis { get; set; }
            public string Display { get; set; }
            public override string ToString() => Display;
        }

        public Room3DColorDialog(List<RoomEntry> entries, List<string> parameterNames)
        {
            InitializeComponent();
            _entries = entries ?? new List<RoomEntry>();
            _parameterNames = parameterNames ?? new List<string>();
            ApplyLocalization();
            InitializeControls();
            _initialized = true;
            RebuildGroups();
        }

        private void ApplyLocalization()
        {
            this.Title = Loc.S("Room3D.Title");
            HeaderText.Text = Loc.S("Room3D.Header");
            txtBasisLabel.Text = Loc.S("Room3D.BasisLabel");
            txtParameterLabel.Text = Loc.S("Room3D.ParameterLabel");
            txtPreviewLabel.Text = Loc.S("Room3D.PreviewLabel");
            txtColorHint.Text = Loc.S("Room3D.ColorHint");
            txtOutputLabel.Text = Loc.S("Room3D.OutputLabel");
            txtViewNameLabel.Text = Loc.S("Room3D.ViewNameLabel");
            DeleteExistingCheckBox.Content = Loc.S("Room3D.DeleteExisting");
            CreateLegendCheckBox.Content = Loc.S("Room3D.CreateLegend");
            ViewNameTextBox.Text = Loc.S("Room3D.DefaultViewName");
            OkButton.Content = Loc.S("Common.Execute");
            CancelButton.Content = Loc.S("Common.Cancel");
        }

        private void InitializeControls()
        {
            var basisItems = new List<BasisItem>
            {
                new BasisItem { Basis = RoomColorBasis.ByName, Display = Loc.S("Room3D.BasisName") },
                new BasisItem { Basis = RoomColorBasis.ByLevel, Display = Loc.S("Room3D.BasisLevel") },
                new BasisItem { Basis = RoomColorBasis.ByParameter, Display = Loc.S("Room3D.BasisParameter") },
                new BasisItem { Basis = RoomColorBasis.PerRoom, Display = Loc.S("Room3D.BasisPerRoom") },
            };
            BasisComboBox.ItemsSource = basisItems;
            BasisComboBox.SelectedIndex = 0;

            ParameterComboBox.ItemsSource = _parameterNames;
            if (_parameterNames.Count > 0)
                ParameterComboBox.SelectedIndex = 0;

            UpdateParameterPanelVisibility();
        }

        private RoomColorBasis SelectedBasis
            => (BasisComboBox.SelectedItem as BasisItem)?.Basis ?? RoomColorBasis.ByName;

        private void UpdateParameterPanelVisibility()
        {
            bool isParam = SelectedBasis == RoomColorBasis.ByParameter;
            ParameterPanel.Visibility = isParam ? Visibility.Visible : Visibility.Collapsed;
        }

        private void Basis_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!_initialized) return;
            UpdateParameterPanelVisibility();
            RebuildGroups();
        }

        private void Parameter_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!_initialized) return;
            if (SelectedBasis == RoomColorBasis.ByParameter)
                RebuildGroups();
        }

        private void RebuildGroups()
        {
            string paramName = ParameterComboBox.SelectedItem as string;
            _groups = RoomColorManager.BuildGroups(_entries, SelectedBasis, paramName);
            RefreshPreview();
        }

        private void RefreshPreview()
        {
            GroupPreviewPanel.Children.Clear();

            if (_groups.Count == 0)
            {
                GroupPreviewPanel.Children.Add(new TextBlock
                {
                    Text = Loc.S("Room3D.NoGroups"),
                    FontSize = 11,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x99, 0x99, 0x99)),
                    Margin = new Thickness(4, 4, 0, 0)
                });
                return;
            }

            for (int i = 0; i < _groups.Count; i++)
            {
                int idx = i;
                var g = _groups[i];

                var row = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, i > 0 ? 5 : 0, 0, 0)
                };

                var swatch = new WpfRectangle
                {
                    Width = 32,
                    Height = 20,
                    Fill = new SolidColorBrush(Color.FromRgb(g.Color.Red, g.Color.Green, g.Color.Blue)),
                    Stroke = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66)),
                    StrokeThickness = 1,
                    Margin = new Thickness(4, 0, 10, 0),
                    Cursor = Cursors.Hand
                };
                swatch.MouseLeftButtonDown += (s, ev) => ShowColorPicker(idx);
                row.Children.Add(swatch);

                row.Children.Add(new TextBlock
                {
                    Text = string.Format(Loc.S("Room3D.GroupRow"), g.Label, g.RoomIds.Count),
                    FontSize = 12,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = g.Label
                });

                GroupPreviewPanel.Children.Add(row);
            }
        }

        private void ShowColorPicker(int index)
        {
            if (index < 0 || index >= _groups.Count) return;
            var g = _groups[index];
            var dlg = new WinForms.ColorDialog
            {
                FullOpen = true,
                Color = System.Drawing.Color.FromArgb(g.Color.Red, g.Color.Green, g.Color.Blue)
            };
            if (dlg.ShowDialog() == WinForms.DialogResult.OK)
            {
                g.Color = new RevitColor(dlg.Color.R, dlg.Color.G, dlg.Color.B);
                RefreshPreview();
            }
        }

        public RoomColorResult GetResult()
        {
            string viewName = ViewNameTextBox.Text?.Trim();
            if (string.IsNullOrEmpty(viewName))
                viewName = Loc.S("Room3D.DefaultViewName");

            return new RoomColorResult
            {
                Basis = SelectedBasis,
                ParameterName = ParameterComboBox.SelectedItem as string,
                Groups = _groups,
                ViewName = viewName,
                DeleteExisting = DeleteExistingCheckBox.IsChecked == true,
                CreateLegend = CreateLegendCheckBox.IsChecked == true
            };
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (_groups.Count == 0)
            {
                MessageBox.Show(Loc.S("Room3D.NoGroups"),
                    Loc.S("Common.InputError"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
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

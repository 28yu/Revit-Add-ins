using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WpfRectangle = System.Windows.Shapes.Rectangle;

namespace Tools28.Commands.FireProtection
{
    public partial class ColorPickerDialog : Window
    {
        private Color _currentColor;
        private Color _selectedColor;
        private bool _updatingRgb;
        private int _nextCustomSlot;

        // 基本色 (8x6 = 48色、Windows標準と同等)
        private static readonly string[] BasicColors = new[]
        {
            "#FF8080","#FFFF80","#80FF80","#00FF80","#80FFFF","#0080FF","#FF80C0","#FF80FF",
            "#FF0000","#FFFF00","#80FF00","#00FF40","#00FFFF","#0080C0","#8080C0","#FF00FF",
            "#804040","#FF8040","#00FF00","#008080","#004080","#8080FF","#800040","#FF0080",
            "#800000","#FF8000","#008000","#008040","#0000FF","#0000A0","#800080","#8000FF",
            "#400000","#804000","#004000","#004040","#000080","#000040","#400040","#400080",
            "#000000","#808000","#808040","#808080","#408080","#C0C0C0","#400040","#FFFFFF",
        };

        // 作成した色 (16スロット)
        private static Color[] _customColors = new Color[16];
        private static bool _customColorsInitialized;

        public Color SelectedColor => _selectedColor;

        public ColorPickerDialog(Color initialColor)
        {
            InitializeComponent();
            _currentColor = initialColor;
            _selectedColor = initialColor;

            if (!_customColorsInitialized)
            {
                for (int i = 0; i < 16; i++)
                    _customColors[i] = Colors.White;
                _customColorsInitialized = true;
            }

            InitializeBasicColors();
            InitializeCustomColors();
            UpdatePreview();
            UpdateRgbInputs();
        }

        private void InitializeBasicColors()
        {
            BasicColorGrid.Children.Clear();
            foreach (var hex in BasicColors)
            {
                var color = (Color)ColorConverter.ConvertFromString(hex);
                var rect = CreateColorSwatch(color, 20, 20);
                rect.MouseLeftButtonDown += (s, e) =>
                {
                    _selectedColor = color;
                    UpdatePreview();
                    UpdateRgbInputs();
                };
                BasicColorGrid.Children.Add(rect);
            }
        }

        private void InitializeCustomColors()
        {
            CustomColorGrid.Children.Clear();
            for (int i = 0; i < 16; i++)
            {
                var color = _customColors[i];
                int slot = i;
                var rect = CreateColorSwatch(color, 20, 18);
                rect.MouseLeftButtonDown += (s, e) =>
                {
                    _selectedColor = _customColors[slot];
                    UpdatePreview();
                    UpdateRgbInputs();
                };
                CustomColorGrid.Children.Add(rect);
            }
        }

        private WpfRectangle CreateColorSwatch(Color color, double w, double h)
        {
            return new WpfRectangle
            {
                Width = w,
                Height = h,
                Fill = new SolidColorBrush(color),
                Stroke = new SolidColorBrush(Color.FromRgb(0x80, 0x80, 0x80)),
                StrokeThickness = 1,
                Margin = new Thickness(1),
                Cursor = Cursors.Hand
            };
        }

        private void UpdatePreview()
        {
            CurrentColorPreview.Fill = new SolidColorBrush(_currentColor);
            NewColorPreview.Fill = new SolidColorBrush(_selectedColor);
            ColorNameText.Text = $"RGB {_selectedColor.R}-{_selectedColor.G}-{_selectedColor.B}";
        }

        private void UpdateRgbInputs()
        {
            _updatingRgb = true;
            RedInput.Text = _selectedColor.R.ToString();
            GreenInput.Text = _selectedColor.G.ToString();
            BlueInput.Text = _selectedColor.B.ToString();
            _updatingRgb = false;
        }

        private void RGB_Changed(object sender, TextChangedEventArgs e)
        {
            if (_updatingRgb) return;

            byte r, g, b;
            if (byte.TryParse(RedInput.Text, out r) &&
                byte.TryParse(GreenInput.Text, out g) &&
                byte.TryParse(BlueInput.Text, out b))
            {
                _selectedColor = Color.FromRgb(r, g, b);
                UpdatePreview();
            }
        }

        private void AddCustomColor_Click(object sender, RoutedEventArgs e)
        {
            _customColors[_nextCustomSlot % 16] = _selectedColor;
            _nextCustomSlot++;
            InitializeCustomColors();
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
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

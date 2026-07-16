using System.ComponentModel;
using System.Windows.Media;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.FillPatternIO
{
    /// <summary>
    /// 一覧表示用の塗り潰しパターン情報（チェック状態付き）。
    /// </summary>
    public class FillPatternItem : INotifyPropertyChanged
    {
        private bool _isSelected;

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public ElementId Id { get; }
        public string Name { get; }
        public FillPatternTarget Target { get; }
        public bool IsSolid { get; }
        public int GridCount { get; }

        private readonly FillPattern _fillPattern;
        private ImageSource _preview;
        private bool _rendered;

        /// <summary>
        /// 一覧に表示するパターンのプレビュー画像。
        /// 大量パターン時のフリーズを避けるため、初回アクセス時（＝行が実体化された時）に
        /// 遅延生成する。ListView の仮想化と併せて、画面に見えている行だけ描画される。
        /// </summary>
        public ImageSource Preview
        {
            get
            {
                if (!_rendered)
                {
                    _rendered = true;
                    try { _preview = PatternPreview.Render(_fillPattern, 220, 26); }
                    catch { _preview = null; }
                }
                return _preview;
            }
        }

        /// <summary>種類の表示ラベル（ローカライズ）。</summary>
        public string TargetLabel => Target == FillPatternTarget.Model
            ? Loc.S("FillPatternIO.Model")
            : Loc.S("FillPatternIO.Drafting");

        public FillPatternItem(FillPatternElement element)
        {
            Id = element.Id;
            _fillPattern = element.GetFillPattern();
            Name = _fillPattern.Name;
            Target = _fillPattern.Target;
            IsSolid = _fillPattern.IsSolidFill;
            GridCount = _fillPattern.GetFillGrids().Count;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

using System.ComponentModel;
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

        /// <summary>種類の表示ラベル（ローカライズ）。</summary>
        public string TargetLabel => Target == FillPatternTarget.Model
            ? Loc.S("FillPatternIO.Model")
            : Loc.S("FillPatternIO.Drafting");

        public FillPatternItem(FillPatternElement element)
        {
            Id = element.Id;
            var fp = element.GetFillPattern();
            Name = fp.Name;
            Target = fp.Target;
            IsSolid = fp.IsSolidFill;
            GridCount = fp.GetFillGrids().Count;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

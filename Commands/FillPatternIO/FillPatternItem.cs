using System;
using System.ComponentModel;
using System.Windows.Media;
using System.Windows.Threading;
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

        private readonly PatternData _data;
        private ImageSource _preview;
        private bool _requested;

        /// <summary>
        /// 一覧に表示するパターンのプレビュー画像。
        /// 初回アクセス時（＝行が実体化された時）にバックグラウンドで描画を依頼し、
        /// 完了したら PropertyChanged で通知して差し替える。UI スレッドはブロックしない。
        /// </summary>
        public ImageSource Preview
        {
            get
            {
                if (!_requested)
                {
                    _requested = true;
                    var ui = Dispatcher.CurrentDispatcher; // getter は UI スレッドで呼ばれる
                    PreviewRenderQueue.Enqueue(_data, 220, 26, img =>
                    {
                        Action apply = () =>
                        {
                            _preview = img;
                            OnPropertyChanged(nameof(Preview));
                        };
                        if (ui.CheckAccess()) apply();
                        else ui.BeginInvoke(DispatcherPriority.Background, apply);
                    });
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
            var fp = element.GetFillPattern();
            Name = fp.Name;
            Target = fp.Target;
            IsSolid = fp.IsSolidFill;
            _data = PatternData.From(fp);
            GridCount = _data.Grids.Count;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

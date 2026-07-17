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
        private bool _previewRequested;
        private ImageSource _largePreview;
        private bool _largeRequested;

        /// <summary>
        /// 一覧に表示するパターンのプレビュー画像（小）。
        /// 初回アクセス時（＝行が実体化された時）にバックグラウンドで描画を依頼し、
        /// 完了したら PropertyChanged で通知して差し替える。UI スレッドはブロックしない。
        /// </summary>
        public ImageSource Preview
        {
            get
            {
                if (!_previewRequested)
                {
                    _previewRequested = true;
                    RequestRender(220, 26, img => { _preview = img; OnPropertyChanged(nameof(Preview)); });
                }
                return _preview;
            }
        }

        /// <summary>ホバー時に表示する拡大プレビュー画像。初回アクセス時に遅延生成する。</summary>
        public ImageSource LargePreview
        {
            get
            {
                if (!_largeRequested)
                {
                    _largeRequested = true;
                    RequestRender(440, 140, img => { _largePreview = img; OnPropertyChanged(nameof(LargePreview)); });
                }
                return _largePreview;
            }
        }

        private void RequestRender(int w, int h, Action<ImageSource> assign)
        {
            var ui = Dispatcher.CurrentDispatcher; // getter は UI スレッドで呼ばれる
            PreviewRenderQueue.Enqueue(_data, w, h, img =>
            {
                Action apply = () => assign(img);
                if (ui.CheckAccess()) apply();
                else ui.BeginInvoke(DispatcherPriority.Background, apply);
            });
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

using System.ComponentModel;
using Tools28.Localization;

namespace Tools28.Commands.DwgLayerTransfer.Models
{
    /// <summary>INotifyPropertyChanged の共通実装（ダイアログ用ビューモデルの基底）。</summary>
    public abstract class NotifyBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnChanged(string name)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>
    /// 移行元 / 移行先の DWG 一覧の 1 行。
    /// 移行元では「そのビューにある DWG」だけが並ぶ。
    /// </summary>
    public sealed class DwgItem : NotifyBase
    {
        public DwgDefinition Dwg { get; set; }

        public string Name => Dwg?.Name ?? "";

        public string KindLabel => Dwg?.KindLabel ?? "";

        public int LayerCount => Dwg?.Layers.Count ?? 0;

        private int _settingCount = -1;
        /// <summary>
        /// 選択中のビューでこの DWG に付いている設定の件数（移行元のみ）。
        /// -1 は「対象外（移行先の一覧）」を表す。
        /// </summary>
        public int SettingCount
        {
            get => _settingCount;
            set { if (_settingCount != value) { _settingCount = value; OnChanged(nameof(SettingCount)); OnChanged(nameof(DetailText)); } }
        }

        private int _matchedLayerCount = -1;
        /// <summary>移行元の DWG と名前が一致するレイヤ数（移行先のみ）。-1 は未算出。</summary>
        public int MatchedLayerCount
        {
            get => _matchedLayerCount;
            set { if (_matchedLayerCount != value) { _matchedLayerCount = value; OnChanged(nameof(MatchedLayerCount)); OnChanged(nameof(DetailText)); } }
        }

        /// <summary>一覧の右側に出す補足情報</summary>
        public string DetailText
        {
            get
            {
                if (SettingCount >= 0)
                    return string.Format(Loc.S("DwgVg.Dwg.SourceDetail"), LayerCount, SettingCount);
                if (MatchedLayerCount >= 0)
                    return string.Format(Loc.S("DwgVg.Dwg.TargetDetail"), LayerCount, MatchedLayerCount);
                return string.Format(Loc.S("DwgVg.Dwg.PlainDetail"), LayerCount);
            }
        }
    }

    /// <summary>移行先ビュー一覧の 1 行（チェックボックスで複数選択できる）。</summary>
    public sealed class TargetViewRow : NotifyBase
    {
        public ViewEntry Entry { get; set; }

        public string Name => Entry?.Name ?? "";

        /// <summary>移行先として使えない理由（使える場合は null）</summary>
        public string BlockReason => Entry?.BlockReason;

        public bool IsApplicable => BlockReason == null;

        /// <summary>適用できない行か（状態列を警告色にするかの判定に使う）</summary>
        public bool IsBlocked => !IsApplicable;

        /// <summary>ブロック理由、無ければ補足（テンプレートへ振り替える旨など）</summary>
        public string StatusText => BlockReason ?? Entry?.Note ?? "";

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                // 適用できない行はチェックできない
                bool v = value && IsApplicable;
                if (_isSelected != v) { _isSelected = v; OnChanged(nameof(IsSelected)); }
            }
        }
    }
}

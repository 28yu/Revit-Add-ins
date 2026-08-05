using System.Collections.Generic;
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
    /// 移行元 DWG と移行先 DWG の対応 1件分。
    /// 既定では同名の DWG を自動で紐付け、違う場合はコンボボックスで手動割り当てできる。
    /// </summary>
    public sealed class DwgPairRow : NotifyBase
    {
        public DwgDefinition Source { get; set; }

        /// <summary>移行先の候補（「(対応なし)」を先頭に含む）</summary>
        public List<string> Candidates { get; set; } = new List<string>();

        private string _selectedTarget;
        /// <summary>選択中の移行先 DWG 名。<see cref="NoMatch"/> は対応なし。</summary>
        public string SelectedTarget
        {
            get => _selectedTarget;
            set
            {
                if (_selectedTarget != value)
                {
                    _selectedTarget = value;
                    OnChanged(nameof(SelectedTarget));
                    OnChanged(nameof(StatusText));
                    OnChanged(nameof(IsResolved));
                }
            }
        }

        /// <summary>「対応なし」を表す選択肢のラベル</summary>
        public static string NoMatch => Loc.S("DwgVg.NoMatch");

        public bool IsResolved => !string.IsNullOrEmpty(SelectedTarget) && SelectedTarget != NoMatch;

        public string SourceLabel => Source?.DisplayLabel ?? "";

        /// <summary>移行先 DWG のレイヤ名集合（名前一致件数の算出に使う。ダイアログ側が設定）</summary>
        public int MatchedLayerCount { get; set; }

        public string StatusText
        {
            get
            {
                if (!IsResolved) return Loc.S("DwgVg.Status.DwgMissing");
                int total = Source?.Layers.Count ?? 0;
                return string.Format(Loc.S("DwgVg.Status.LayerMatch"), MatchedLayerCount, total);
            }
        }
    }

    /// <summary>
    /// 移行元ビュー（またはビューテンプレート）と移行先ビューの対応 1件分。
    /// </summary>
    public sealed class ViewPairRow : NotifyBase
    {
        public ViewSettingSnapshot Snapshot { get; set; }

        public string SourceName => Snapshot?.View?.Name ?? "";

        /// <summary>移行元に設定されている DWG レイヤ設定の件数</summary>
        public int SettingCount => Snapshot?.SettingCount ?? 0;

        /// <summary>移行先の候補（「(対応なし)」を先頭に含む）</summary>
        public List<string> Candidates { get; set; } = new List<string>();

        private string _selectedTarget;
        public string SelectedTarget
        {
            get => _selectedTarget;
            set
            {
                if (_selectedTarget != value)
                {
                    _selectedTarget = value;
                    OnChanged(nameof(SelectedTarget));
                    OnChanged(nameof(StatusText));
                    OnChanged(nameof(IsResolved));
                    OnChanged(nameof(IsApplicable));
                    // 対応先が変わると適用可否も変わるためチェック状態を追従させる
                    if (!IsApplicable) IsSelected = false;
                }
            }
        }

        public static string NoMatch => Loc.S("DwgVg.NoMatch");

        public bool IsResolved => !string.IsNullOrEmpty(SelectedTarget) && SelectedTarget != NoMatch;

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                // 対応先が無い行・適用できない行はチェックできない
                bool v = value && IsApplicable;
                if (_isSelected != v) { _isSelected = v; OnChanged(nameof(IsSelected)); }
            }
        }

        private string _blockReason;
        /// <summary>移行先が使用できない理由（ビューテンプレート制御下など）。使える場合は null。</summary>
        public string BlockReason
        {
            get => _blockReason;
            set
            {
                if (_blockReason != value)
                {
                    _blockReason = value;
                    if (!IsApplicable) IsSelected = false;
                    OnChanged(nameof(BlockReason));
                    OnChanged(nameof(StatusText));
                    OnChanged(nameof(IsApplicable));
                }
            }
        }

        public string StatusText
        {
            get
            {
                if (BlockReason != null) return BlockReason;
                if (!IsResolved) return Loc.S("DwgVg.Status.ViewMissing");
                return Loc.S("DwgVg.Status.Ready");
            }
        }

        /// <summary>実際に適用できる行か</summary>
        public bool IsApplicable => IsResolved && BlockReason == null;
    }
}

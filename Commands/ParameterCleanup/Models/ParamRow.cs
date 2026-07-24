using System.Collections.Generic;
using System.ComponentModel;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.ParameterCleanup.Models
{
    /// <summary>パラメータの種別</summary>
    public enum ParamKind
    {
        Project,   // 非共有プロジェクトパラメータ
        Shared,    // 共有パラメータ
        Global     // グローバルパラメータ
    }

    /// <summary>値の有無の判定状態</summary>
    public enum ValueState
    {
        Unchecked,      // 未確認
        Checking,       // 確認中
        HasValue,       // 値あり
        Empty,          // 値なし（空）
        NotApplicable   // 判定対象外（バインド無し／グローバル等）
    }

    /// <summary>
    /// 削除候補として一覧表示する1パラメータ分の情報（ダイアログ用ビューモデル）。
    /// </summary>
    public class ParamRow : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnChanged(string name)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // --- 不変の基本情報 ---
        public string Name { get; set; }
        public ParamKind Kind { get; set; }

        /// <summary>削除対象となる要素 ID（ParameterElement / SharedParameterElement / GlobalParameter）</summary>
        public ElementId Id { get; set; }

        /// <summary>値スキャン用の Definition（バインド済みパラメータのみ）</summary>
        public Definition Definition { get; set; }

        /// <summary>タイプパラメータなら true、インスタンスなら false、バインド無しは null</summary>
        public bool? IsTypeBinding { get; set; }

        /// <summary>バインド先カテゴリ（値スキャン対象）。バインド無しは空。</summary>
        public List<Category> BoundCategories { get; set; } = new List<Category>();

        /// <summary>カテゴリ表示文字列</summary>
        public string CategoriesText { get; set; } = "";

        /// <summary>グローバルパラメータの現在値（表示用）</summary>
        public string GlobalValueText { get; set; } = "";

        // --- 可変（UI バインド）---
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { if (_isSelected != value) { _isSelected = value; OnChanged(nameof(IsSelected)); } }
        }

        private ValueState _state = ValueState.Unchecked;
        public ValueState State
        {
            get => _state;
            set
            {
                if (_state != value)
                {
                    _state = value;
                    OnChanged(nameof(State));
                    OnChanged(nameof(StateText));
                }
            }
        }

        private bool _isDuplicateName;
        public bool IsDuplicateName
        {
            get => _isDuplicateName;
            set { if (_isDuplicateName != value) { _isDuplicateName = value; OnChanged(nameof(IsDuplicateName)); } }
        }

        /// <summary>値スキャン可能か（バインド済みで対象カテゴリを持つ）</summary>
        public bool IsScannable
            => Kind != ParamKind.Global && Definition != null && BoundCategories != null && BoundCategories.Count > 0;

        /// <summary>種別の表示文字列</summary>
        public string KindText
        {
            get
            {
                switch (Kind)
                {
                    case ParamKind.Project: return Loc.S("ParamCleanup.Kind.Project");
                    case ParamKind.Shared: return Loc.S("ParamCleanup.Kind.Shared");
                    case ParamKind.Global: return Loc.S("ParamCleanup.Kind.Global");
                    default: return "";
                }
            }
        }

        /// <summary>バインド種別（インスタンス／タイプ）の表示文字列</summary>
        public string ScopeText
        {
            get
            {
                if (Kind == ParamKind.Global) return "-";
                if (IsTypeBinding == true) return Loc.S("ParamCleanup.Scope.Type");
                if (IsTypeBinding == false) return Loc.S("ParamCleanup.Scope.Instance");
                return "-";
            }
        }

        /// <summary>値の有無の表示文字列</summary>
        public string StateText
        {
            get
            {
                switch (State)
                {
                    case ValueState.Unchecked: return Loc.S("ParamCleanup.State.Unchecked");
                    case ValueState.Checking: return Loc.S("ParamCleanup.State.Checking");
                    case ValueState.HasValue: return Loc.S("ParamCleanup.State.HasValue");
                    case ValueState.Empty: return Loc.S("ParamCleanup.State.Empty");
                    case ValueState.NotApplicable:
                        return Kind == ParamKind.Global
                            ? GlobalValueText
                            : Loc.S("ParamCleanup.State.NotApplicable");
                    default: return "";
                }
            }
        }
    }
}

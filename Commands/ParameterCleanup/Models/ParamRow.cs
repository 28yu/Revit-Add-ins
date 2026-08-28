using System;
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
        HasValue,       // 値あり（削除すると値が失われる）
        Empty,          // 要素に存在するが全て空
        NotFound,       // どの要素にも存在しない（定義のみ＝安全に削除可）
        NotApplicable   // 判定対象外（グローバルパラメータ）
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

        /// <summary>値スキャン用の Definition</summary>
        public Definition Definition { get; set; }

        /// <summary>共有パラメータの GUID（非共有は Guid.Empty）。バインド無しでも値を引くための鍵。</summary>
        public Guid SharedGuid { get; set; } = Guid.Empty;

        /// <summary>タイプパラメータなら true、インスタンスなら false、バインド無しは null</summary>
        public bool? IsTypeBinding { get; set; }

        /// <summary>プロジェクトパラメータとしてのバインド先カテゴリ。バインド無しは空。</summary>
        public List<Category> BoundCategories { get; set; } = new List<Category>();

        /// <summary>バインド先カテゴリの表示文字列</summary>
        public string CategoriesText { get; set; } = "";

        /// <summary>このパラメータを参照している集計表名（カンマ区切り）。無ければ空。</summary>
        public string ScheduleRefText { get; set; } = "";

        /// <summary>グローバルパラメータの現在値（表示用）</summary>
        public string GlobalValueText { get; set; } = "";

        /// <summary>プロジェクトのカテゴリにバインドされていない（＝ファミリ内部定義等）</summary>
        public bool IsNotBound => Kind != ParamKind.Global && BoundCategories.Count == 0;

        // --- スキャン結果（UI バインド）---

        private string _usageText = "";
        /// <summary>このパラメータを実際に保持している要素の内訳（カテゴリ (件数), …）</summary>
        public string UsageText
        {
            get => _usageText;
            set
            {
                if (_usageText != value)
                {
                    _usageText = value ?? "";
                    OnChanged(nameof(UsageText));
                    OnChanged(nameof(UsageDisplayText));
                    OnChanged(nameof(UsageTooltip));
                }
            }
        }

        private int _usageElementCount;
        /// <summary>このパラメータを保持している要素の総数</summary>
        public int UsageElementCount
        {
            get => _usageElementCount;
            set
            {
                if (_usageElementCount != value)
                {
                    _usageElementCount = value;
                    OnChanged(nameof(UsageElementCount));
                    OnChanged(nameof(UsageDisplayText));
                    OnChanged(nameof(UsageTooltip));
                }
            }
        }

        private string _sampleText = "";
        /// <summary>値が見つかった代表例（ファミリ名 / 要素ID : 値）</summary>
        public string SampleText
        {
            get => _sampleText;
            set
            {
                if (_sampleText != value)
                {
                    _sampleText = value ?? "";
                    OnChanged(nameof(SampleText));
                    OnChanged(nameof(StateTooltip));
                    OnChanged(nameof(UsageTooltip));
                }
            }
        }

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
                    OnChanged(nameof(StateTooltip));
                    OnChanged(nameof(ScopeText));
                    OnChanged(nameof(UsageDisplayText));
                }
            }
        }

        private bool _isDuplicateName;
        public bool IsDuplicateName
        {
            get => _isDuplicateName;
            set { if (_isDuplicateName != value) { _isDuplicateName = value; OnChanged(nameof(IsDuplicateName)); } }
        }

        /// <summary>値スキャン対象か（グローバル以外はバインドの有無にかかわらず全て対象）</summary>
        public bool IsScannable => Kind != ParamKind.Global;

        /// <summary>削除しても値が失われないと確認済みか</summary>
        public bool IsSafeToDelete => State == ValueState.NotFound || State == ValueState.Empty;

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

        /// <summary>バインド種別（インスタンス／タイプ／ファミリ内）の表示文字列</summary>
        public string ScopeText
        {
            get
            {
                if (Kind == ParamKind.Global) return "-";
                if (IsTypeBinding == true) return Loc.S("ParamCleanup.Scope.Type");
                if (IsTypeBinding == false) return Loc.S("ParamCleanup.Scope.Instance");
                // プロジェクト未バインド。要素上に実在するならファミリ内部定義とみなす。
                if (State == ValueState.HasValue || State == ValueState.Empty)
                    return Loc.S("ParamCleanup.Scope.InFamily");
                return Loc.S("ParamCleanup.Scope.NotBound");
            }
        }

        /// <summary>使用箇所列の表示文字列</summary>
        public string UsageDisplayText
        {
            get
            {
                if (Kind == ParamKind.Global) return "-";
                switch (State)
                {
                    case ValueState.Unchecked:
                    case ValueState.Checking:
                        return "";
                    case ValueState.NotFound:
                        return Loc.S("ParamCleanup.Usage.None");
                    default:
                        return string.Format(Loc.S("ParamCleanup.Usage.Summary"),
                                             UsageElementCount, UsageText);
                }
            }
        }

        /// <summary>使用箇所列のツールチップ</summary>
        public string UsageTooltip
        {
            get
            {
                if (Kind == ParamKind.Global) return Loc.S("ParamCleanup.Tip.Global");
                if (State == ValueState.NotFound) return Loc.S("ParamCleanup.Tip.NotFound");
                if (string.IsNullOrEmpty(UsageText)) return "";

                string s = string.Format(Loc.S("ParamCleanup.Usage.Summary"), UsageElementCount, UsageText);
                if (!string.IsNullOrEmpty(SampleText))
                    s += "\n" + SampleText;
                return s;
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
                    case ValueState.NotFound: return Loc.S("ParamCleanup.State.NotFound");
                    case ValueState.NotApplicable:
                        return Kind == ParamKind.Global
                            ? GlobalValueText
                            : Loc.S("ParamCleanup.State.NotApplicable");
                    default: return "";
                }
            }
        }

        /// <summary>値の状態の意味（マウスオーバー時のツールチップ用）</summary>
        public string StateTooltip
        {
            get
            {
                switch (State)
                {
                    case ValueState.HasValue:
                        {
                            string s = Loc.S("ParamCleanup.Tip.HasValue");
                            if (!string.IsNullOrEmpty(SampleText)) s += "\n" + SampleText;
                            return s;
                        }
                    case ValueState.Empty: return Loc.S("ParamCleanup.Tip.Empty");
                    case ValueState.NotFound: return Loc.S("ParamCleanup.Tip.NotFound");
                    case ValueState.Unchecked: return Loc.S("ParamCleanup.Tip.Unchecked");
                    case ValueState.Checking: return Loc.S("ParamCleanup.Tip.Checking");
                    case ValueState.NotApplicable:
                        return Kind == ParamKind.Global
                            ? Loc.S("ParamCleanup.Tip.Global")
                            : Loc.S("ParamCleanup.Tip.NotBound");
                    default: return "";
                }
            }
        }
    }
}

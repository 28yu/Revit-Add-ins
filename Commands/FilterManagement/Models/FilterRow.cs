using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.FilterManagement.Models
{
    /// <summary>フィルタの種別</summary>
    public enum FilterKind
    {
        Rule,       // ルールベースフィルタ（ParameterFilterElement）
        Selection   // 選択フィルタ（SelectionFilterElement）
    }

    /// <summary>フィルタが使用されている1ビュー分の情報</summary>
    public class ViewUsage
    {
        public string ViewName { get; set; }
        public bool IsTemplate { get; set; }

        /// <summary>表示用文字列（テンプレートには印を付ける）</summary>
        public string Display
            => IsTemplate ? $"[{Loc.S("FilterMgmt.TemplateMark")}] {ViewName}" : ViewName;
    }

    /// <summary>
    /// 一覧表示する1フィルタ分の情報（ダイアログ用ビューモデル）。
    /// </summary>
    public class FilterRow : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnChanged(string name)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // --- 基本情報 ---
        /// <summary>削除対象となるフィルタ要素の ID（ParameterFilterElement / SelectionFilterElement）</summary>
        public ElementId Id { get; set; }

        public FilterKind Kind { get; set; }

        /// <summary>読込直後のフィルタ名（名前変更の検出・取り消し用）</summary>
        public string OriginalName { get; set; }

        /// <summary>対象カテゴリの表示文字列（ルールベースのみ。選択フィルタは "-"）</summary>
        public string CategoriesText { get; set; } = "";

        /// <summary>このフィルタを使用しているビュー／ビューテンプレートの一覧</summary>
        public List<ViewUsage> UsedViews { get; set; } = new List<ViewUsage>();

        // --- 可変（UI バインド）---
        private string _name;
        /// <summary>フィルタ名（インライン編集可能）</summary>
        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnChanged(nameof(Name));
                    OnChanged(nameof(IsRenamed));
                }
            }
        }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { if (_isSelected != value) { _isSelected = value; OnChanged(nameof(IsSelected)); } }
        }

        // --- 派生（表示用）---
        /// <summary>いずれかのビューで使用されているか</summary>
        public bool IsUsed => UsedViews != null && UsedViews.Count > 0;

        /// <summary>使用ビュー数</summary>
        public int ViewCount => UsedViews?.Count ?? 0;

        /// <summary>種別の表示文字列</summary>
        public string KindText
            => Kind == FilterKind.Selection ? Loc.S("FilterMgmt.Kind.Selection") : Loc.S("FilterMgmt.Kind.Rule");

        /// <summary>使用状況の表示文字列（未使用なら注意文言、使用中ならビュー名を列挙）</summary>
        public string UsageText
        {
            get
            {
                if (!IsUsed) return Loc.S("FilterMgmt.Unused");
                return string.Join(", ", UsedViews.Select(v => v.Display));
            }
        }

        /// <summary>名前が読込時から変更されているか</summary>
        public bool IsRenamed
            => !string.IsNullOrEmpty(Name) && Name != OriginalName;
    }
}

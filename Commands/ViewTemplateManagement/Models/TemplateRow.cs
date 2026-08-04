using System.Collections.Generic;
using System.ComponentModel;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.ViewTemplateManagement.Models
{
    /// <summary>
    /// 一覧表示する1ビューテンプレート分の情報（ダイアログ用ビューモデル）。
    /// </summary>
    public class TemplateRow : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnChanged(string name)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // --- 基本情報 ---
        /// <summary>削除対象となるビューテンプレート（View）の ID</summary>
        public ElementId Id { get; set; }

        /// <summary>読込直後のテンプレート名（名前変更の検出・取り消し用）</summary>
        public string OriginalName { get; set; }

        /// <summary>ビュー種別の表示文字列（平面図／断面図／3D など）</summary>
        public string ViewTypeText { get; set; } = "";

        /// <summary>このテンプレートを適用しているビュー名の一覧</summary>
        public List<string> UsedViews { get; set; } = new List<string>();

        // --- 可変（UI バインド）---
        private string _name;
        /// <summary>テンプレート名（インライン編集可能）</summary>
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

        /// <summary>使用状況の表示文字列（未使用なら注意文言、使用中ならビュー名を列挙）</summary>
        public string UsageText
        {
            get
            {
                if (!IsUsed) return Loc.S("TemplateMgmt.Unused");
                return string.Join(", ", UsedViews);
            }
        }

        /// <summary>名前が読込時から変更されているか</summary>
        public bool IsRenamed
            => !string.IsNullOrEmpty(Name) && Name != OriginalName;
    }
}

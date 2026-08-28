using System.Collections.Generic;
using System.ComponentModel;
using Autodesk.Revit.DB;

namespace Tools28.Commands.GenericModelMerge.Models
{
    /// <summary>
    /// ダイアログのカテゴリチェックリスト 1 行分。
    /// カテゴリ名は Revit の Category.Name をそのまま単一の情報源として使う
    /// （固定翻訳は Revit のバージョン・言語で実名とずれるため持たない）。
    /// </summary>
    public class MergeCategoryRow : INotifyPropertyChanged
    {
        private bool _isSelected = true;

        public ElementId CategoryId { get; set; }

        /// <summary>Revit が返す実際のカテゴリ名。</summary>
        public string Name { get; set; }

        /// <summary>このカテゴリでビュー内に見つかった要素数。</summary>
        public int ElementCount => ElementIds.Count;

        /// <summary>このカテゴリに属するビュー内要素の Id。</summary>
        public List<ElementId> ElementIds { get; } = new List<ElementId>();

        /// <summary>チェックリストに出す表示文字列。「カテゴリ名 (要素数)」。</summary>
        public string Display => $"{Name}  ({ElementCount})";

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    /// <summary>材質コンボボックス 1 行分。</summary>
    public class MaterialRow
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public override string ToString() => Name;
    }
}

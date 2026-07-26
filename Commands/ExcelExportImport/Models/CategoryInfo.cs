using System.ComponentModel;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;

namespace Tools28.Commands.ExcelExportImport.Models
{
    /// <summary>
    /// カテゴリ情報を保持するモデル
    /// </summary>
    public class CategoryInfo : INotifyPropertyChanged
    {
        private bool _isChecked;

        public BuiltInCategory BuiltInCategory { get; set; }
        public string Name { get; set; }

        /// <summary>
        /// チェック状態。全選択/選択解除でコードから変更した際に
        /// CheckBox へ即時反映させるため INotifyPropertyChanged で通知する。
        /// </summary>
        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                if (_isChecked == value) return;
                _isChecked = value;
                OnPropertyChanged();
            }
        }

        public int ElementCount { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// 表示用ラベル。Revit が返す実際のカテゴリ名 (Name) をそのまま使う。
        /// パラメータ欄・出力欄・Excel シート名・設定も同じ Name を使うため、
        /// 全表示が Revit の表記と一致し、セクション間で食い違わない。
        /// （旧: CategoryLocalizer の固定翻訳は Revit のバージョン/言語で実名とずれるため撤去）
        /// </summary>
        public string DisplayLabel => $"{Name} ({ElementCount})";

        public CategoryInfo(BuiltInCategory builtInCategory, string name, int elementCount)
        {
            BuiltInCategory = builtInCategory;
            Name = name;
            ElementCount = elementCount;
            IsChecked = false;
        }

        public override string ToString()
        {
            return DisplayLabel;
        }
    }
}

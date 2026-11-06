using System.Collections.Generic;
using System.ComponentModel;
using Autodesk.Revit.DB;

namespace Tools28.Commands.DuplicateInstance.Models
{
    /// <summary>
    /// 「同じ場所に同一のインスタンスがあります」警告に含まれる要素1つ分の情報（ダイアログの1行）。
    /// 同じ場所で重なっている要素どうしは同じ GroupNo を持つ。
    /// </summary>
    public class DuplicateRow : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        /// <summary>組の番号（1始まり）。同じ場所で重なっている要素は同じ番号。</summary>
        public int GroupNo { get; set; }

        /// <summary>この組に含まれる要素の数</summary>
        public int GroupSize { get; set; }

        public ElementId Id { get; set; }

        /// <summary>並べ替え用の要素 ID の数値（大きいほど後から作られた要素）</summary>
        public long IdValue { get; set; }

        public string Category { get; set; } = "";
        public string Family { get; set; } = "";
        public string TypeName { get; set; } = "";
        public string Level { get; set; } = "";
        public string Workset { get; set; } = "";
        public string Mark { get; set; } = "";
        public string Comments { get; set; } = "";
        public string Host { get; set; } = "";

        /// <summary>値が入っているインスタンスパラメータの数（「値が少ないほうを選択」の判断に使う）</summary>
        public int FilledCount { get; set; }

        /// <summary>同じ組の他の要素と値が違うパラメータの数</summary>
        public int DiffCount { get; set; }

        /// <summary>インスタンスパラメータ（表示名 → 表示用の値）。比較表に使う。</summary>
        public Dictionary<string, string> Parameters { get; set; } = new Dictionary<string, string>();

        private bool _isSelected;
        /// <summary>削除対象としてチェックされているか</summary>
        public bool IsSelected
        {
            get => _isSelected;
            set { if (_isSelected != value) { _isSelected = value; OnChanged(nameof(IsSelected)); } }
        }

        /// <summary>組ごとに背景色を交互に変えるための印（偶数組なら true）</summary>
        public bool IsEvenGroup => GroupNo % 2 == 0;
    }
}

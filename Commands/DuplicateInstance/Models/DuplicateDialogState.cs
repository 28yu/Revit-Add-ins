using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.DuplicateInstance.Models
{
    /// <summary>
    /// 「画面で確認」でダイアログを一度閉じて開き直すときに、チェックや絞り込みの状態を引き継ぐための入れ物。
    /// </summary>
    public class DuplicateDialogState
    {
        public HashSet<long> CheckedIds { get; } = new HashSet<long>();
        public string SearchText { get; set; } = "";
        public string Category { get; set; }
        public bool DiffOnly { get; set; }
        public long SelectedId { get; set; } = -1;

        /// <summary>並べ替えに使っている列（"fixed:Id" や "param:マーク" など。null なら並べ替えなし）</summary>
        public string SortKey { get; set; }
        public bool SortDescending { get; set; }

        public bool HasBounds { get; set; }
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double TopPaneHeight { get; set; }

        /// <summary>ダイアログが閉じられたあと Revit 画面で表示したい要素（null なら終了）</summary>
        public List<ElementId> ShowRequest { get; set; }
    }

    /// <summary>比較表の1行（パラメータ1つ分）。Values[i] が組の i 番目の要素の値。</summary>
    public class CompareRow
    {
        public string Name { get; set; }
        public string[] Values { get; set; }
        public bool IsDiff { get; set; }
        public bool IsHeader { get; set; }
    }
}

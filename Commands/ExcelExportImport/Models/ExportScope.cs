namespace Tools28.Commands.ExcelExportImport.Models
{
    /// <summary>
    /// エクスポート対象の範囲
    /// </summary>
    public enum ExportScope
    {
        /// <summary>プロジェクト全体の要素</summary>
        EntireProject,
        /// <summary>現在のビューにある要素</summary>
        ActiveView,
        /// <summary>現在選択されている要素</summary>
        Selection
    }
}

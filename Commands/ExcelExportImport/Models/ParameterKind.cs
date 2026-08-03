namespace Tools28.Commands.ExcelExportImport.Models
{
    /// <summary>
    /// パラメータの種別（同名パラメータの区別に使う）。
    /// 既定値は Project（0）＝共有でも組み込みでもないプロジェクト/ファミリパラメータ。
    /// </summary>
    public enum ParameterKind
    {
        /// <summary>プロジェクト/ファミリパラメータ（共有でも組み込みでもない）</summary>
        Project = 0,

        /// <summary>組み込み（BuiltInParameter）パラメータ</summary>
        BuiltIn = 1,

        /// <summary>共有パラメータ（Parameter.IsShared == true）</summary>
        Shared = 2
    }
}

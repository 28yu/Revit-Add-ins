namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// エクスポート処理から進み具合を伝え、キャンセル要求を受け取るための窓口。
    /// 画面（ExportProgressWindow）とエクスポート処理（ExcelExportService）を切り離すために使う。
    /// </summary>
    public interface IExportProgress
    {
        /// <summary>キャンセルボタンが押されたか</summary>
        bool IsCancelRequested { get; }

        /// <summary>カテゴリの処理を始める（index は 1 始まり）</summary>
        void BeginCategory(int index, int total, string categoryName, int elementCount);

        /// <summary>要素 1 件の書き出しが終わった</summary>
        void ElementDone();

        /// <summary>処理の合間に呼ぶ（一定時間ごとに表示更新・キャンセル受付が行われる）</summary>
        void Tick();

        /// <summary>今いちばん時間のかかっているパラメータを伝える</summary>
        void SetSlowest(string parameterLabel, double averageMs);

        /// <summary>Excel ファイルの保存を始める（保存中はキャンセル不可）</summary>
        void BeginSaving();
    }
}

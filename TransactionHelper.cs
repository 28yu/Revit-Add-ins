using System;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28
{
    /// <summary>
    /// Transaction の確定を確実に確認するための共通ヘルパー。
    /// Commit() は例外を出さずに RolledBack 等を返すことがある
    /// （ワークシェアで他ユーザーが借用中の要素・無視できないエラー等）。
    /// 戻り値を見ないと、モデルに反映されていないのに「完了」と表示してしまう。
    /// </summary>
    internal static class TransactionHelper
    {
        /// <summary>確定できなかった場合は例外を投げる（呼び出し側の catch でエラー表示する）</summary>
        public static void CommitOrThrow(this Transaction t)
        {
            var status = t.Commit();
            if (status != TransactionStatus.Committed)
                throw new InvalidOperationException(string.Format(Loc.S("Common.CommitFailed"), status));
        }

        /// <summary>開始中のときだけ取り消す（確定失敗後に RollBack を呼ぶと例外になるため）</summary>
        public static void RollBackIfStarted(this Transaction t)
        {
            if (t != null && t.GetStatus() == TransactionStatus.Started)
                t.RollBack();
        }
    }
}

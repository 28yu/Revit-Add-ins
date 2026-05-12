using System;

namespace Tools28.Licensing
{
    /// <summary>
    /// バージョンの有効期限を管理する。
    /// 28Tools の段階的有償化戦略 Phase 1（無償版の期限管理）として、
    /// 各バージョンに有効期限を設定し、期限切れ後は最新版への更新を促す。
    ///
    /// 期限切れ判定はローカル日時で行う（サーバー不要、オフライン環境対応）。
    /// PC日付の変更で回避可能だが、建築業界向け B2B ツールとして許容範囲のリスクとする。
    /// </summary>
    internal static class ExpiryManager
    {
        /// <summary>
        /// 現バージョンの有効期限。リリース時に1年後の日付を設定する。
        /// 新バージョンリリース時にこの値を更新する。
        /// </summary>
        public static readonly DateTime ExpiryDate = new DateTime(2027, 6, 1);

        /// <summary>期限切れ前に警告を表示し始める日数。</summary>
        private const int WarningDays = 30;

        /// <summary>有効期限を過ぎているか。</summary>
        public static bool IsExpired => DateTime.Now > ExpiryDate;

        /// <summary>有効期限まで残り何日か（過ぎていれば負値）。</summary>
        public static int DaysRemaining => (int)Math.Ceiling((ExpiryDate - DateTime.Now).TotalDays);

        /// <summary>期限切れ警告を表示するべきか。</summary>
        public static bool ShouldShowWarning => !IsExpired && DaysRemaining <= WarningDays;
    }
}

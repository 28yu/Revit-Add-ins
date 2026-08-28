using System;
using System.Collections.Generic;
using System.IO;

namespace Tools28
{
    /// <summary>
    /// 診断用ログ。コマンド実行の進行を C:\temp\Tools28_debug.txt に追記する。
    /// 失敗してもユーザー体験を損ねないよう全て swallow する。
    ///
    /// 【従来の問題】
    /// 常時 ON・サイズ上限なし・無効化手段なしだったため、長期運用でログが際限なく肥大化していた。
    /// また Application / SheetCreation / FilledRegionSplitMerge / FireProtection が
    /// それぞれ独自にファイル追記していて、挙動が揃っていなかった。
    ///
    /// 【現在の仕様】
    ///  - サイズが MaxBytes を超えたら日時付きファイル名へ退避し、退避ファイルは KeepRotations 世代だけ残す
    ///  - %AppData%\Tools28\diaglog.setting に "off" と書くと無効化できる（既定は有効）
    ///  - WriteTo() で別ファイルへ書く場合も同じローテーションが効く
    /// </summary>
    internal static class DiagLog
    {
        private const string LogDir = @"C:\temp";
        private const string DefaultFileName = "Tools28_debug.txt";

        /// <summary>このサイズを超えたらローテーションする</summary>
        private const long MaxBytes = 5L * 1024 * 1024;

        /// <summary>退避ファイルを残す世代数</summary>
        private const int KeepRotations = 3;

        private static readonly object _sync = new object();
        private static readonly HashSet<string> _headerWritten = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static bool? _enabled;

        /// <summary>ログ出力が有効か。設定ファイルで無効化できる（既定は有効）。</summary>
        internal static bool Enabled
        {
            get
            {
                if (_enabled.HasValue) return _enabled.Value;
                _enabled = LoadEnabledSetting();
                return _enabled.Value;
            }
        }

        public static void Write(string message)
        {
            WriteTo(DefaultFileName, message);
        }

        public static void Cmd(string commandName, string phase)
        {
            Write($"[CMD:{commandName}] {phase}");
        }

        /// <summary>ファイル名を指定して書き込む（機能別にログを分けたい場合に使う）。</summary>
        public static void WriteTo(string fileName, string message)
        {
            if (!Enabled) return;

            try
            {
                lock (_sync)
                {
                    Directory.CreateDirectory(LogDir);
                    string path = Path.Combine(LogDir, fileName);

                    RotateIfNeeded(path);

                    if (_headerWritten.Add(fileName))
                    {
                        File.AppendAllText(path,
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] === 28 Tools 診断ログ開始 ===\n" +
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ログを止めるには %AppData%\\Tools28\\diaglog.setting に off と書いてください\n");
                    }

                    File.AppendAllText(path, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n");
                }
            }
            catch { }
        }

        /// <summary>
        /// サイズ上限を超えていたら日時付きファイル名へ退避し、古い退避ファイルを削除する。
        /// 呼び出し側で lock 済みであること。
        /// </summary>
        private static void RotateIfNeeded(string path)
        {
            try
            {
                var info = new FileInfo(path);
                if (!info.Exists || info.Length < MaxBytes) return;

                string dir = info.DirectoryName ?? LogDir;
                string stem = Path.GetFileNameWithoutExtension(path);
                string ext = Path.GetExtension(path);
                string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                File.Move(path, Path.Combine(dir, $"{stem}_{stamp}{ext}"));
                _headerWritten.Remove(Path.GetFileName(path));

                // 退避ファイルは新しい順に KeepRotations 件だけ残す
                var olds = new DirectoryInfo(dir).GetFiles($"{stem}_*{ext}");
                Array.Sort(olds, (a, b) => string.CompareOrdinal(b.Name, a.Name));
                for (int i = KeepRotations; i < olds.Length; i++)
                {
                    try { olds[i].Delete(); } catch { }
                }
            }
            catch { }
        }

        private static bool LoadEnabledSetting()
        {
            try
            {
                string path = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Tools28", "diaglog.setting");
                if (File.Exists(path))
                    return !File.ReadAllText(path).Trim().Equals("off", StringComparison.OrdinalIgnoreCase);
            }
            catch { }
            return true;   // 既定は有効（不具合報告時にログが残っている状態を優先する）
        }
    }
}

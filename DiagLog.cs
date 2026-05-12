using System;
using System.IO;

namespace Tools28
{
    /// <summary>
    /// 診断用ログ。コマンド実行の進行を C:\temp\Tools28_debug.txt に追記する。
    /// 失敗してもユーザー体験を損ねないよう全て swallow する。
    /// </summary>
    internal static class DiagLog
    {
        private const string LogPath = @"C:\temp\Tools28_debug.txt";

        public static void Write(string message)
        {
            try
            {
                Directory.CreateDirectory(@"C:\temp");
                File.AppendAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n");
            }
            catch { }
        }

        public static void Cmd(string commandName, string phase)
        {
            Write($"[CMD:{commandName}] {phase}");
        }
    }
}

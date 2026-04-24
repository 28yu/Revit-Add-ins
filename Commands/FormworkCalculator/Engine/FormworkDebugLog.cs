using System;
using System.IO;
using System.Text;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 型枠数量算出のデバッグログを C:\temp\Formwork_debug.txt に出力する。
    ///
    /// - スレッドセーフ (lock で排他制御)
    /// - 実行開始時にファイルをクリアする
    /// - 出力行数が上限 (MaxLines) を超えたら以降は書き込みを停止し、末尾に
    ///   "... truncated" を追記する。Revit が応答しなくなるのを防ぐため。
    /// - Enabled = false の場合は何もしない
    /// </summary>
    internal static class FormworkDebugLog
    {
        private const string DefaultPath = @"C:\temp\Formwork_debug.txt";
        private const int MaxLines = 200000;

        private static readonly object _lock = new object();
        private static StreamWriter _writer;
        private static int _lineCount;
        private static bool _enabled;
        private static bool _truncated;

        internal static bool Enabled => _enabled;

        internal static void Initialize(bool enable, string path = null)
        {
            lock (_lock)
            {
                CloseInternal();
                _enabled = enable;
                _lineCount = 0;
                _truncated = false;

                if (!_enabled) return;

                try
                {
                    string filePath = string.IsNullOrEmpty(path) ? DefaultPath : path;
                    string dir = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                        Directory.CreateDirectory(dir);

                    _writer = new StreamWriter(filePath, false, new UTF8Encoding(false))
                    {
                        AutoFlush = false,
                    };

                    _writer.WriteLine($"==== Formwork Debug Log Start [{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ====");
                    _lineCount++;
                }
                catch
                {
                    // ログ失敗はメイン処理に影響させない
                    _writer = null;
                    _enabled = false;
                }
            }
        }

        internal static void Log(string msg)
        {
            if (!_enabled) return;
            lock (_lock)
            {
                if (_writer == null) return;
                if (_truncated) return;
                if (_lineCount >= MaxLines)
                {
                    try { _writer.WriteLine("... truncated (MaxLines reached) ..."); } catch { }
                    _truncated = true;
                    return;
                }

                try
                {
                    _writer.WriteLine(msg);
                    _lineCount++;
                }
                catch { }
            }
        }

        internal static void Section(string title)
        {
            if (!_enabled) return;
            Log(string.Empty);
            Log($"---- {title} ----");
        }

        internal static void Flush()
        {
            if (!_enabled) return;
            lock (_lock)
            {
                try { _writer?.Flush(); } catch { }
            }
        }

        internal static void Close()
        {
            lock (_lock) { CloseInternal(); }
        }

        private static void CloseInternal()
        {
            if (_writer != null)
            {
                try
                {
                    _writer.WriteLine($"==== End [{DateTime.Now:yyyy-MM-dd HH:mm:ss}] total lines: {_lineCount} ====");
                    _writer.Flush();
                    _writer.Dispose();
                }
                catch { }
                _writer = null;
            }
            _enabled = false;
        }
    }
}

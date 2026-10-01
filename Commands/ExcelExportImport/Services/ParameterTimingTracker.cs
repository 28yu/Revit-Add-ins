using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Tools28.Commands.ExcelExportImport.Models;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// パラメータごとに「値の読み取りにかかった時間」を集計する。
    /// 大容量モデルで極端に遅いパラメータ（設備系の計算値など）を見つけ、
    /// 進み具合の画面・キャンセル時のメッセージ・ログで知らせるために使う。
    /// </summary>
    public class ParameterTimingTracker
    {
        public class Entry
        {
            public string Label;
            public long Ticks;
            public int Count;

            public double TotalSeconds => (double)Ticks / Stopwatch.Frequency;
            public double AverageMs => Count == 0 ? 0 : TotalSeconds * 1000.0 / Count;
        }

        // ParameterInfo の同一判定は ParamId＋インスタンス/タイプ＋カテゴリ
        private readonly Dictionary<ParameterInfo, Entry> _entries = new Dictionary<ParameterInfo, Entry>();

        /// <summary>合計時間が最も長いパラメータ（まだ無ければ null）</summary>
        public Entry Slowest { get; private set; }

        public void Add(ParameterInfo param, long ticks)
        {
            if (!_entries.TryGetValue(param, out var entry))
            {
                entry = new Entry { Label = param.DisplayName + " (" + param.CategoryName + ")" };
                _entries[param] = entry;
            }
            entry.Ticks += ticks;
            entry.Count++;

            if (Slowest == null || entry.Ticks > Slowest.Ticks)
                Slowest = entry;
        }

        /// <summary>合計時間の長い順に上位 count 件を返す</summary>
        public List<Entry> Top(int count)
        {
            return _entries.Values
                .OrderByDescending(e => e.Ticks)
                .Take(count)
                .ToList();
        }
    }
}

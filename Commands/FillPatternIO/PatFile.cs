using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;

namespace Tools28.Commands.FillPatternIO
{
    /// <summary>
    /// Revit の塗り潰しパターン (FillPattern) と AutoCAD/Revit 形式の .pat ファイルを
    /// 相互変換するユーティリティ。
    ///
    /// Revit 内部の長さ単位はフィート。.pat 側の単位は先頭の
    /// <c>;%UNITS=MM</c> / <c>;%UNITS=INCH</c> ヘッダーで決まる。
    /// グリッド行の書式（AutoCAD 準拠）:
    ///   角度, X原点, Y原点, シフト(delta-x), オフセット(delta-y)[, ダッシュ1, ダッシュ2, ...]
    /// </summary>
    internal static class PatFile
    {
        public const double MmPerFoot = 304.8;
        public const double InchPerFoot = 12.0;

        public class ParsedPattern
        {
            public string Name;
            public FillPatternTarget Target = FillPatternTarget.Drafting;
            public List<FillGrid> Grids = new List<FillGrid>();
        }

        private static string Fmt(double v)
            => v.ToString("0.########", CultureInfo.InvariantCulture);

        /// <summary>1つの FillPattern を .pat 形式テキストへ変換する。</summary>
        public static string BuildPatText(FillPattern fp, bool useMm)
        {
            double scale = useMm ? MmPerFoot : InchPerFoot;
            var sb = new StringBuilder();

            sb.AppendLine(useMm ? ";%UNITS=MM" : ";%UNITS=INCH");
            sb.AppendLine($"*{fp.Name},Tools28");
            sb.AppendLine(fp.Target == FillPatternTarget.Model ? ";%TYPE=MODEL" : ";%TYPE=DRAFTING");

            foreach (FillGrid grid in fp.GetFillGrids())
            {
                var parts = new List<string>
                {
                    Fmt(grid.Angle * 180.0 / Math.PI),
                    Fmt(grid.Origin.U * scale),
                    Fmt(grid.Origin.V * scale),
                    Fmt(grid.Shift * scale),
                    Fmt(grid.Offset * scale)
                };
                foreach (double seg in grid.GetSegments())
                    parts.Add(Fmt(seg * scale));

                sb.AppendLine(string.Join(",", parts));
            }

            return sb.ToString();
        }

        /// <summary>FillPatternElement を .pat ファイルへ書き出す。</summary>
        public static void ExportToFile(FillPatternElement fpe, string path, bool useMm)
        {
            string text = BuildPatText(fpe.GetFillPattern(), useMm);
            File.WriteAllText(path, text, new UTF8Encoding(false));
        }

        private static List<double> ParseNumbers(string line)
        {
            var nums = new List<double>();
            foreach (var token in line.Split(','))
            {
                var t = token.Trim();
                if (t.Length == 0) continue;
                if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out double d))
                    nums.Add(d);
                else
                    return null; // 数値でないトークンを含む行は無効とみなす
            }
            return nums;
        }

        /// <summary>.pat ファイルを解析して複数のパターン定義を取り出す。</summary>
        public static List<ParsedPattern> Parse(string path, bool defaultMm)
        {
            var result = new List<ParsedPattern>();
            bool useMm = defaultMm;
            ParsedPattern cur = null;

            foreach (var raw in File.ReadAllLines(path))
            {
                string line = raw.Trim();
                if (line.Length == 0) continue;

                if (line.StartsWith(";"))
                {
                    string up = line.ToUpperInvariant();
                    if (up.Contains("UNITS=MM")) useMm = true;
                    else if (up.Contains("UNITS=INCH")) useMm = false;
                    else if (up.Contains("TYPE=MODEL") && cur != null) cur.Target = FillPatternTarget.Model;
                    else if (up.Contains("TYPE=DRAFTING") && cur != null) cur.Target = FillPatternTarget.Drafting;
                    continue;
                }

                if (line.StartsWith("*"))
                {
                    cur = new ParsedPattern();
                    string body = line.Substring(1);
                    int comma = body.IndexOf(',');
                    cur.Name = (comma >= 0 ? body.Substring(0, comma) : body).Trim();
                    result.Add(cur);
                    continue;
                }

                if (cur == null) continue;

                var nums = ParseNumbers(line);
                if (nums == null || nums.Count < 5) continue;

                double scale = useMm ? MmPerFoot : InchPerFoot;
                var grid = new FillGrid
                {
                    Angle = nums[0] * Math.PI / 180.0,
                    Origin = new UV(nums[1] / scale, nums[2] / scale),
                    Shift = nums[3] / scale,
                    Offset = nums[4] / scale
                };

                if (nums.Count > 5)
                {
                    var segs = new List<double>();
                    for (int i = 5; i < nums.Count; i++)
                        segs.Add(nums[i] / scale);
                    grid.SetSegments(segs);
                }

                cur.Grids.Add(grid);
            }

            return result;
        }

        /// <summary>
        /// 解析済みパターンをドキュメントへ作成する。呼び出し側でトランザクションを開くこと。
        /// </summary>
        /// <param name="skipped">同名同種のパターンが既存のため作成しなかった数</param>
        /// <param name="invalid">グリッドが無い等で作成できなかった数</param>
        /// <returns>作成したパターン数</returns>
        public static int CreatePatterns(Document doc, List<ParsedPattern> patterns,
            out int skipped, out int invalid)
        {
            skipped = 0;
            invalid = 0;
            int created = 0;

            var existing = new HashSet<string>(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(FillPatternElement))
                    .Cast<FillPatternElement>()
                    .Select(e =>
                    {
                        var fp = e.GetFillPattern();
                        return Key(fp.Name, fp.Target);
                    }));

            foreach (var p in patterns)
            {
                if (p.Grids.Count == 0) { invalid++; continue; }

                string key = Key(p.Name, p.Target);
                if (existing.Contains(key)) { skipped++; continue; }

                try
                {
                    var orientation = p.Target == FillPatternTarget.Model
                        ? FillPatternHostOrientation.ToHost
                        : FillPatternHostOrientation.ToView;

                    var fp = new FillPattern(p.Name, p.Target, orientation);
                    fp.SetFillGrids(p.Grids);
                    FillPatternElement.Create(doc, fp);

                    existing.Add(key);
                    created++;
                }
                catch
                {
                    invalid++;
                }
            }

            return created;
        }

        private static string Key(string name, FillPatternTarget target)
            => (name ?? string.Empty) + "|" + target;

        /// <summary>ファイル名に使用できない文字を除去する。</summary>
        public static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "pattern";
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder();
            foreach (char c in name)
                sb.Append(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
            return sb.ToString();
        }
    }
}

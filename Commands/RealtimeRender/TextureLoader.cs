using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Tools28.Commands.RealtimeRender
{
    /// <summary>
    /// マテリアルの外観のテクスチャ画像を探して読み込む。
    /// Revit は画像の場所を「ファイル名だけ」や「1\Mats\…」のような相対パスで持っていることが多いので、
    /// Autodesk のマテリアル ライブラリのフォルダ・オプションの「追加のレンダリング外観パス」・プロジェクトのフォルダから探す。
    /// </summary>
    internal class TextureLoader
    {
        private readonly List<string> _searchDirs = new List<string>();
        private Dictionary<string, string> _fileIndex;  // ファイル名 → 見つかった場所（ライブラリの中を一度だけ調べる）
        private readonly Dictionary<string, string> _resolved = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public List<string> Missing { get; } = new List<string>();

        public TextureLoader(string documentPath, string revitVersion)
        {
            foreach (Environment.SpecialFolder folder in new[] { Environment.SpecialFolder.CommonProgramFilesX86, Environment.SpecialFolder.CommonProgramFiles })
            {
                string root = Environment.GetFolderPath(folder);
                if (!string.IsNullOrEmpty(root))
                    AddDir(Path.Combine(root, "Autodesk Shared", "Materials", "Textures"));
            }
            foreach (string dir in ReadAdditionalPaths(revitVersion)) AddDir(dir);
            try
            {
                if (!string.IsNullOrEmpty(documentPath)) AddDir(Path.GetDirectoryName(documentPath));
            }
            catch
            {
                // クラウドのモデルなど、パスとして扱えないときは使わない
            }
        }

        private void AddDir(string dir)
        {
            if (!string.IsNullOrWhiteSpace(dir) && Directory.Exists(dir) && !_searchDirs.Contains(dir)) _searchDirs.Add(dir);
        }

        /// <summary>Revit のオプション「追加のレンダリング外観パス」（Revit.ini）を読む。</summary>
        private static IEnumerable<string> ReadAdditionalPaths(string revitVersion)
        {
            var result = new List<string>();
            try
            {
                string ini = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                          "Autodesk", "Revit", "Autodesk Revit " + revitVersion, "Revit.ini");
                if (!File.Exists(ini)) return result;
                foreach (string line in File.ReadAllLines(ini))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0 || !line.Substring(0, eq).Trim().Equals("AdditionalRenderAppearancePaths", StringComparison.OrdinalIgnoreCase))
                        continue;
                    foreach (string p in line.Substring(eq + 1).Split('|', ';'))
                        if (!string.IsNullOrWhiteSpace(p)) result.Add(Environment.ExpandEnvironmentVariables(p.Trim()));
                }
            }
            catch
            {
                // 読めなければ使わない
            }
            return result;
        }

        /// <summary>画像の場所を探す。見つからなければ null。paths は「|」区切りで複数の候補が入っていることがある。</summary>
        public string Resolve(string paths)
        {
            if (string.IsNullOrWhiteSpace(paths)) return null;
            if (_resolved.TryGetValue(paths, out string cached)) return cached;
            string found = null;
            foreach (string raw in paths.Split('|'))
            {
                string p = Environment.ExpandEnvironmentVariables(raw.Trim()).Replace('/', '\\');
                if (p.Length == 0) continue;
                try
                {
                    if (Path.IsPathRooted(p) && File.Exists(p)) { found = p; break; }
                    foreach (string dir in _searchDirs)
                    {
                        string candidate = Path.Combine(dir, p.TrimStart('\\'));
                        if (File.Exists(candidate)) { found = candidate; break; }
                    }
                    if (found != null) break;
                    if (FileIndex().TryGetValue(Path.GetFileName(p), out string indexed)) { found = indexed; break; }
                }
                catch
                {
                    // パスに使えない文字などは飛ばす
                }
            }
            _resolved[paths] = found;
            if (found == null) Missing.Add(paths);
            return found;
        }

        private Dictionary<string, string> FileIndex()
        {
            if (_fileIndex != null) return _fileIndex;
            _fileIndex = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string dir in _searchDirs)
            {
                try
                {
                    foreach (string file in Directory.EnumerateFiles(dir, "*.*", SearchOption.AllDirectories))
                    {
                        string name = Path.GetFileName(file);
                        if (!_fileIndex.ContainsKey(name)) _fileIndex[name] = file;
                    }
                }
                catch
                {
                    // 読めないフォルダは飛ばす
                }
            }
            return _fileIndex;
        }

        /// <summary>
        /// 画像を読み込み、長い辺が maxSize 以下になるよう縮小して、RGBA（1 画素 4 バイト）で返す。
        /// 読めない形式（WPF が対応していない .tga / .exr など）は null。
        /// </summary>
        public static RenderTexture Decode(string path, int maxSize)
        {
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    BitmapDecoder decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.IgnoreColorProfile, BitmapCacheOption.OnLoad);
                    BitmapSource source = decoder.Frames[0];
                    int longest = Math.Max(source.PixelWidth, source.PixelHeight);
                    if (longest <= 0) return null;
                    if (longest > maxSize)
                    {
                        double s = (double)maxSize / longest;
                        source = new TransformedBitmap(source, new ScaleTransform(s, s));
                    }
                    source = new FormatConvertedBitmap(source, PixelFormats.Bgra32, null, 0);
                    int w = source.PixelWidth, h = source.PixelHeight;
                    var bgra = new byte[w * h * 4];
                    source.CopyPixels(bgra, w * 4, 0);
                    // BGRA → RGBA
                    for (int i = 0; i < bgra.Length; i += 4)
                    {
                        byte b = bgra[i];
                        bgra[i] = bgra[i + 2];
                        bgra[i + 2] = b;
                    }
                    return new RenderTexture { Width = w, Height = h, Rgba = bgra };
                }
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"画像を読めませんでした: {path} ({ex.Message})");
                return null;
            }
        }

        /// <summary>画像の明るさの平均（0〜1）。粗さの画像を 1 つの値にするのに使う。</summary>
        public static double? AverageLuminance(string path)
        {
            RenderTexture t = Decode(path, 64);
            if (t == null) return null;
            double sum = 0;
            for (int i = 0; i < t.Rgba.Length; i += 4)
                sum += (t.Rgba[i] * 0.2126 + t.Rgba[i + 1] * 0.7152 + t.Rgba[i + 2] * 0.0722) / 255.0;
            return sum / (t.Width * t.Height);
        }
    }
}

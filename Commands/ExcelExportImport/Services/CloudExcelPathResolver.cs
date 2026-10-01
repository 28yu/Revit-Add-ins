using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Win32;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// Excel で開いているクラウド上のブック（OneDrive / SharePoint）の場所を、読み込める場所に置き換える。
    ///
    /// 【背景】クラウド上のブックは Excel の FullName が「https://…」の URL になり、
    /// そのままでは ClosedXML で開けない（「指定されたパスの形式はサポートされていません」）。
    /// 一方「参照」で選ぶと OneDrive の同期フォルダ内の実ファイル（C:\Users\…\OneDrive…）になるため読める。
    /// そこで次の順に解決し、「参照」と同じファイルを読むようにする。
    ///  1. OneDrive の同期設定（レジストリ）から URL → 同期フォルダ内の実ファイルに変換
    ///  2. 見つからなければ、Excel で開いている内容を一時フォルダに複製して使う
    /// </summary>
    public static class CloudExcelPathResolver
    {
        /// <summary>URL（クラウド上のブック）か</summary>
        public static bool IsCloudPath(string path)
            => !string.IsNullOrEmpty(path)
               && (path.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
                   || path.StartsWith("http://", StringComparison.OrdinalIgnoreCase));

        /// <summary>一覧表示用のファイル名（URL の %20 などを元の文字に戻す）</summary>
        public static string GetDisplayFileName(string path)
        {
            if (!IsCloudPath(path))
                return Path.GetFileName(path);

            string decoded = SafeUnescape(path);
            int q = decoded.IndexOf('?');
            if (q >= 0) decoded = decoded.Substring(0, q);
            int slash = decoded.LastIndexOf('/');
            return slash >= 0 ? decoded.Substring(slash + 1) : decoded;
        }

        /// <summary>
        /// 読み込みに使うファイルパスを返す。ローカルのパスはそのまま返す。
        /// クラウドの URL は同期フォルダの実ファイル、無ければ一時フォルダへの複製に置き換える。
        /// どちらもできなければ null。
        /// </summary>
        public static string Resolve(string fullName)
        {
            if (!IsCloudPath(fullName))
                return fullName;

            string local = TryMapToSyncedFile(fullName);
            if (local != null)
            {
                DiagLog.Write($"[ExcelImport] クラウドのブックを同期フォルダに変換: {fullName} → {local}");
                return local;
            }

            string copy = ExcelProcessHelper.SaveOpenWorkbookCopy(fullName);
            if (copy != null)
                DiagLog.Write($"[ExcelImport] クラウドのブックを一時フォルダに複製: {fullName} → {copy}");
            else
                DiagLog.Write($"[ExcelImport] クラウドのブックを解決できません: {fullName}");
            return copy;
        }

        /// <summary>OneDrive の同期設定から、URL に対応する同期フォルダ内の実ファイルを探す。</summary>
        private static string TryMapToSyncedFile(string url)
        {
            try
            {
                string decodedUrl = SafeUnescape(url);
                int q = decodedUrl.IndexOf('?');
                if (q >= 0) decodedUrl = decodedUrl.Substring(0, q);

                // URL の前半（同期しているライブラリの場所）が長く一致するものから試す
                foreach (var mapping in GetSyncMappings()
                    .Where(m => decodedUrl.StartsWith(m.UrlNamespace, StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(m => m.UrlNamespace.Length))
                {
                    string rest = decodedUrl.Substring(mapping.UrlNamespace.Length).Trim('/');
                    var segments = rest.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

                    // ライブラリの一部フォルダだけを同期している場合は、同期フォルダが URL の途中に当たるため、
                    // 先頭のフォルダを1つずつ省きながら実在するファイルを探す
                    for (int skip = 0; skip < segments.Length; skip++)
                    {
                        string relative = string.Join(Path.DirectorySeparatorChar.ToString(), segments.Skip(skip));
                        string candidate = Path.Combine(mapping.MountPoint, relative);
                        if (File.Exists(candidate))
                            return candidate;
                    }
                }
            }
            catch (Exception ex)
            {
                DiagLog.Write($"[ExcelImport] 同期フォルダの検索で例外: {ex.Message}");
            }
            return null;
        }

        private class SyncMapping
        {
            public string UrlNamespace; // 末尾 "/" 付き・デコード済み
            public string MountPoint;
        }

        /// <summary>レジストリから「クラウドの URL ↔ 同期フォルダ」の対応を集める。</summary>
        private static List<SyncMapping> GetSyncMappings()
        {
            var list = new List<SyncMapping>();

            void Add(string urlNamespace, string mountPoint)
            {
                if (string.IsNullOrWhiteSpace(urlNamespace) || string.IsNullOrWhiteSpace(mountPoint))
                    return;
                if (!Directory.Exists(mountPoint))
                    return;
                string ns = SafeUnescape(urlNamespace.Trim());
                if (!ns.EndsWith("/")) ns += "/";
                list.Add(new SyncMapping { UrlNamespace = ns, MountPoint = mountPoint });
            }

            // 同期中のライブラリ（OneDrive 個人用／職場用、SharePoint）ごとの対応
            using (var root = Registry.CurrentUser.OpenSubKey(@"Software\SyncEngines\Providers\OneDrive"))
            {
                if (root != null)
                {
                    foreach (string name in root.GetSubKeyNames())
                    {
                        using (var key = root.OpenSubKey(name))
                        {
                            Add(key?.GetValue("UrlNamespace") as string, key?.GetValue("MountPoint") as string);
                        }
                    }
                }
            }

            // 念のため OneDrive のアカウント情報からも補う
            using (var accounts = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\OneDrive\Accounts"))
            {
                if (accounts != null)
                {
                    foreach (string name in accounts.GetSubKeyNames())
                    {
                        using (var key = accounts.OpenSubKey(name))
                        {
                            string userFolder = key?.GetValue("UserFolder") as string;
                            string cid = key?.GetValue("cid") as string;
                            string userUrl = key?.GetValue("UserUrl") as string;

                            // 個人用 OneDrive: https://d.docs.live.net/{cid}/…
                            if (!string.IsNullOrEmpty(cid))
                                Add("https://d.docs.live.net/" + cid + "/", userFolder);
                            // 職場用 OneDrive: https://{tenant}-my.sharepoint.com/personal/{user}/Documents/…
                            if (!string.IsNullOrEmpty(userUrl))
                                Add(userUrl.TrimEnd('/') + "/Documents/", userFolder);
                        }
                    }
                }
            }

            return list;
        }

        private static string SafeUnescape(string s)
        {
            try { return Uri.UnescapeDataString(s); }
            catch { return s; }
        }
    }
}

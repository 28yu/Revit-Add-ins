using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.DwgLayerTransfer.Models
{
    /// <summary>
    /// モデル内に取り込まれている DWG 1件分の定義（カテゴリ＋レイヤ）。
    ///
    /// Revit では読み込み／リンクした DWG は「DWGファイル名」のカテゴリとして登録され、
    /// その各レイヤがサブカテゴリになる。V/G ダイアログの「読み込みカテゴリ」タブは
    /// この階層をそのまま表示している。
    /// </summary>
    public sealed class DwgDefinition
    {
        /// <summary>カテゴリ名（= DWG ファイル名）</summary>
        public string Name { get; set; } = "";

        /// <summary>DWG 本体（親カテゴリ）の ID</summary>
        public ElementId CategoryId { get; set; }

        /// <summary>リンク（true）か読み込み（false）か。判定できない場合は null。</summary>
        public bool? IsLinked { get; set; }

        /// <summary>レイヤ名 -&gt; サブカテゴリ ID</summary>
        public Dictionary<string, ElementId> Layers { get; }
            = new Dictionary<string, ElementId>(StringComparer.CurrentCultureIgnoreCase);

        /// <summary>一覧表示用のラベル（リンク/読み込みの別とレイヤ数を添える）</summary>
        public string DisplayLabel
        {
            get
            {
                string kind = IsLinked == true ? Loc.S("DwgVg.Kind.Link")
                            : IsLinked == false ? Loc.S("DwgVg.Kind.Import")
                            : "";
                return string.IsNullOrEmpty(kind)
                    ? $"{Name} ({Layers.Count})"
                    : $"{Name} [{kind}] ({Layers.Count})";
            }
        }
    }

    /// <summary>ビュー／ビューテンプレート1件分の識別情報。</summary>
    public sealed class ViewEntry
    {
        public ElementId Id { get; set; }
        public string Name { get; set; } = "";
        public bool IsTemplate { get; set; }

        /// <summary>
        /// 移行先として使えない理由（使える場合は null）。
        /// 例: ビューテンプレートが「読み込みカテゴリ」の V/G を制御しているビュー。
        /// </summary>
        public string BlockReason { get; set; }

        public bool IsBlocked => BlockReason != null;
    }

    /// <summary>1ビュー分の DWG レイヤ設定スナップショット。</summary>
    public sealed class ViewSettingSnapshot
    {
        public ViewEntry View { get; set; }

        /// <summary>DWG 名 -&gt; (レイヤ名 -&gt; 設定)。レイヤ名 "" は DWG 本体。</summary>
        public Dictionary<string, Dictionary<string, LayerGraphicSetting>> ByDwg { get; }
            = new Dictionary<string, Dictionary<string, LayerGraphicSetting>>(StringComparer.CurrentCultureIgnoreCase);

        /// <summary>既定から変化している設定の件数（一覧表示用）</summary>
        public int SettingCount
        {
            get
            {
                int n = 0;
                foreach (var layers in ByDwg.Values)
                    foreach (var s in layers.Values)
                        if (s.HasAnySetting) n++;
                return n;
            }
        }
    }
}

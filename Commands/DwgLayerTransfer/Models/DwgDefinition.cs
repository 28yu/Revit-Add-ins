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

        /// <summary>
        /// モデル全体に配置された ImportInstance を持つか（＝どのビューにも現れうる）。
        /// ImportInstance が1つも見つからなかった場合も、取りこぼしを防ぐため true にする。
        /// </summary>
        public bool AppearsEverywhere { get; set; }

        /// <summary>
        /// ビュー固有に貼り付けられた ImportInstance の所有ビュー ID。
        /// 「特定のビューにだけ貼り付けた DWG」がここに入る。
        /// </summary>
        public HashSet<ElementId> OwnerViewIds { get; } = new HashSet<ElementId>();

        /// <summary>この DWG が指定ビューに現れるか。</summary>
        public bool AppearsInView(ElementId viewId)
            => AppearsEverywhere || (viewId != null && OwnerViewIds.Contains(viewId));

        /// <summary>リンク／読み込み／ビュー固有の別を表す短いラベル</summary>
        public string KindLabel
        {
            get
            {
                if (!AppearsEverywhere && OwnerViewIds.Count > 0)
                    return Loc.S("DwgVg.Kind.ViewSpecific");
                if (IsLinked == true) return Loc.S("DwgVg.Kind.Link");
                if (IsLinked == false) return Loc.S("DwgVg.Kind.Import");
                return "";
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
        /// 例: 「読み込みカテゴリ」の V/G を制御していないビューテンプレート。
        /// </summary>
        public string BlockReason { get; set; }

        public bool IsBlocked => BlockReason != null;

        /// <summary>
        /// このビューの「読み込みカテゴリ」V/G を制御しているビューテンプレート（無ければ未設定）。
        /// 制御下のビューは Revit の仕様でビュー側から直接変更できないため、
        /// 書き込み先をこのテンプレートへ振り替える。
        /// </summary>
        public ElementId TemplateId { get; set; }

        public string TemplateName { get; set; } = "";

        /// <summary>テンプレート経由で書き込むビューか。</summary>
        public bool ViaTemplate => TemplateId != null && TemplateId != ElementId.InvalidElementId;

        /// <summary>ブロックではない補足（テンプレートへ振り替える旨など）。</summary>
        public string Note { get; set; }

        /// <summary>実際に設定を書き込む要素の ID。</summary>
        public ElementId WriteTargetId => ViaTemplate ? TemplateId : Id;

        /// <summary>実際に設定を書き込む要素の名前。</summary>
        public string WriteTargetName => ViaTemplate ? TemplateName : Name;
    }

}

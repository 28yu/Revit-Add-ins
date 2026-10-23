using Autodesk.Revit.DB;

namespace Tools28.Commands.GenericModelMerge.Models
{
    /// <summary>生成する一般モデルの形式。</summary>
    internal enum MergeOutputKind
    {
        /// <summary>ダイレクトシェイプ（プロジェクト内に直接置く形状）。</summary>
        DirectShape,
        /// <summary>一般モデルファミリ (.rfa) を作成してロード・配置する。</summary>
        Family,
    }

    /// <summary>一般モデルファミリで作るときに、書き出す・配置するのどちらを行うか。</summary>
    internal enum MergeFamilyAction
    {
        /// <summary>.rfa を保存し、このプロジェクトにも読み込んで配置する。</summary>
        SaveAndPlace,
        /// <summary>.rfa を保存するだけ。プロジェクトには読み込まない。</summary>
        SaveOnly,
        /// <summary>プロジェクトに読み込んで配置するだけ。.rfa は残さない（一時フォルダに作って削除する）。</summary>
        PlaceOnly,
    }

    /// <summary>複数要素の形状を「一つのかたまり」にまとめる方法。</summary>
    internal enum MergeCombineMode
    {
        /// <summary>形状はそのまま。Revit 上の「要素」としてだけ 1 つにまとめる。</summary>
        KeepShapes,
        /// <summary>すべてを結合 (Boolean Union) して単一の形状にする。</summary>
        UnionAll,
        /// <summary>接触・交差しているものだけを結合し、離れているものは別形状のまま残す。</summary>
        UnionTouching,
    }

    /// <summary>ダイアログで決定した実行条件。</summary>
    internal class MergeOptions
    {
        public MergeOutputKind OutputKind { get; set; } = MergeOutputKind.DirectShape;
        public MergeCombineMode CombineMode { get; set; } = MergeCombineMode.KeepShapes;

        /// <summary>生成する一般モデルに割り当てる材質。InvalidElementId なら元の材質のまま。</summary>
        public ElementId MaterialId { get; set; } = ElementId.InvalidElementId;

        /// <summary>元になった要素をアクティブビューで非表示にするか。</summary>
        public bool HideSourceElements { get; set; } = true;

        /// <summary>生成する要素／ファミリの名前。</summary>
        public string Name { get; set; } = "";

        /// <summary>Family 出力時に、書き出す・配置するのどちらを行うか。</summary>
        public MergeFamilyAction FamilyAction { get; set; } = MergeFamilyAction.SaveAndPlace;

        /// <summary>Family 出力で配置する（SaveOnly 以外）か。</summary>
        public bool PlacesFamily => FamilyAction != MergeFamilyAction.SaveOnly;

        /// <summary>Family 出力で .rfa を残す（PlaceOnly 以外）か。</summary>
        public bool KeepsFamilyFile => FamilyAction != MergeFamilyAction.PlaceOnly;

        /// <summary>
        /// Family 出力時の .rfa パス。PlaceOnly のときは一時フォルダ内のパスで、処理後に削除する。
        /// （ロードしたファミリの名前はファイル名になるため、一時ファイルも作成する名前で作る）
        /// </summary>
        public string FamilyPath { get; set; } = "";
    }
}

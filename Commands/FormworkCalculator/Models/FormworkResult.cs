using System.Collections.Generic;

namespace Tools28.Commands.FormworkCalculator.Models
{
    public enum FaceType
    {
        FormworkRequired,
        DeductedTop,
        DeductedBottom,
        DeductedContact,
        DeductedBelowGL,
        Inclined,
        Error,
    }

    public enum CategoryGroup
    {
        Column,
        Beam,
        Wall,
        Slab,
        Foundation,
        Stairs,
        Roof,
        Other,
    }

    public class ElementResult
    {
        public int ElementId { get; set; }
        public string ElementName { get; set; } = string.Empty;
        public CategoryGroup Category { get; set; }
        public string CategoryName { get; set; } = string.Empty;

        /// <summary>
        /// 要素の出自 ("ホスト" or リンクファイル名)。リンクモデル対応で導入。
        /// </summary>
        public string SourceName { get; set; } = string.Empty;
        public string Zone { get; set; } = string.Empty;
        public string FormworkType { get; set; } = string.Empty;
        public double FormworkArea { get; set; }
        public double DeductedTopArea { get; set; }
        public double DeductedBottomArea { get; set; }
        public double DeductedContactArea { get; set; }
        public double InclinedArea { get; set; }
        public double OpeningAreaDeducted { get; set; }
        public double OpeningEdgeAreaAdded { get; set; }

        /// <summary>
        /// 部分接触 (T字結合で主面の一部に他要素が当たっている) が存在するか。
        /// DirectShape のマーカーパラメータや半透明表示の判定に使う。
        /// </summary>
        public bool HasPartialContact { get; set; }
    }

    public class FaceAnalysisResult
    {
        public int SourceElementId { get; set; }
        public FaceType FaceType { get; set; }
        public double Area { get; set; }
        public List<Autodesk.Revit.DB.CurveLoop> BoundaryLoops { get; set; }
            = new List<Autodesk.Revit.DB.CurveLoop>();
        public Autodesk.Revit.DB.XYZ Normal { get; set; }
        public string GroupKey { get; set; } = string.Empty;
    }

    public class CategoryResult
    {
        public CategoryGroup Category { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public double FormworkArea { get; set; }
        public double DeductedArea { get; set; }
        public int ElementCount { get; set; }
    }

    public class ZoneResult
    {
        public string Zone { get; set; } = string.Empty;
        public double FormworkArea { get; set; }
        public int ElementCount { get; set; }
    }

    public class FormworkTypeResult
    {
        public string FormworkType { get; set; } = string.Empty;
        public double FormworkArea { get; set; }
        public int ElementCount { get; set; }
    }

    public class ErrorLogEntry
    {
        public int ElementId { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public string ElementName { get; set; } = string.Empty;
        public string ErrorKind { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }

    /// <summary>
    /// 型枠算出から除外された要素の種別。
    /// </summary>
    public enum ExclusionKind
    {
        Steel,      // 鉄骨部材 (H形鋼・角形鋼管・CFT 等)
        DeckSlab,   // デッキスラブ (タイプ名に "DS" を含む床)
        WallSweep,  // 壁スイープ・リビール (壁の天端や面に付帯する成形)
        SteelStair, // 鉄骨階段 (タイプ名・マテリアルに鉄骨キーワードを含む階段)
        LgsWall,    // LGS壁・乾式壁 (壁構造に石膏ボード層があり、コンクリート層が無い壁)
    }

    /// <summary>
    /// 型枠算出から除外された要素の情報。
    /// 集計には含まれないが、解析3Dビュー上では別色 DirectShape として可視化する。
    /// </summary>
    public class ExcludedResult
    {
        public int ElementId { get; set; }
        public string ElementName { get; set; } = string.Empty;
        public CategoryGroup Category { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public string SourceName { get; set; } = string.Empty;
        public ExclusionKind Kind { get; set; }
        public string DetectionLayer { get; set; } = string.Empty;
        public string DetectionReason { get; set; } = string.Empty;
    }

    public class FormworkResult
    {
        public double TotalFormworkArea { get; set; }
        public double TotalDeductedArea { get; set; }
        public double InclinedFaceArea { get; set; }
        public int ProcessedElementCount { get; set; }

        public List<ElementResult> ElementResults { get; set; } = new List<ElementResult>();
        public List<CategoryResult> CategoryResults { get; set; } = new List<CategoryResult>();
        public List<ZoneResult> ZoneResults { get; set; } = new List<ZoneResult>();
        public List<FormworkTypeResult> TypeResults { get; set; } = new List<FormworkTypeResult>();
        public List<FaceAnalysisResult> FaceResults { get; set; } = new List<FaceAnalysisResult>();
        public List<ErrorLogEntry> Errors { get; set; } = new List<ErrorLogEntry>();
        public List<ExcludedResult> ExcludedResults { get; set; }
            = new List<ExcludedResult>();

        /// <summary>
        /// 要素の出自情報レジストリ (ホスト + リンク)。FormworkCalcEngine が設定し、
        /// FormworkVisualizer / ScheduleCreator が SurrogateId → Element/Document/Transform 解決に使う。
        /// public ではなく internal 扱いとして利用者は触らない。
        /// </summary>
        internal object SourceRegistry { get; set; }

        /// <summary>リンクモデルの統計 (実行サマリー表示用)。</summary>
        public int LinkedInstanceCount { get; set; }
    }
}

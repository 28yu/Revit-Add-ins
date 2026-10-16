using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28
{
    /// <summary>
    /// モデルを開いていなくても押せるボタン用（PushButtonData.AvailabilityClassName に指定する）。
    /// 何も指定しないと、Revit はモデル未オープンの間ボタンを押せなくする。
    /// 設定パネル（自動バックアップ・言語切替・バージョン情報・マニュアル）はモデルを使わないため常に有効にする。
    /// </summary>
    public class AlwaysAvailable : IExternalCommandAvailability
    {
        public const string ClassName = "Tools28.AlwaysAvailable";

        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            return true;
        }
    }
}

using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.GenericModelMerge.Services
{
    /// <summary>
    /// 一般モデル化のトランザクション失敗ハンドラ。
    ///
    /// これを付けないと、Revit が「エラー: 無視できません」のモーダルを出し、
    /// ユーザーがキャンセルするとトランザクションが黙ってロールバックされる。
    /// その結果「中身が空のファミリが保存される」といった分かりにくい失敗になるため、
    /// エラー文言をこちらで収集し、呼び出し側が結果ダイアログで理由を出せるようにする。
    ///
    /// - 警告は削除してコミットを妨げない
    /// - エラーは文言を記録した上でロールバック（中途半端な結果を残さない）
    /// </summary>
    internal class FamilyFailurePreprocessor : IFailuresPreprocessor
    {
        /// <summary>収集したエラー文言（重複なし）。</summary>
        public List<string> Errors { get; } = new List<string>();

        public bool HadError => Errors.Count > 0;

        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            bool hasError = false;

            foreach (var failure in failuresAccessor.GetFailureMessages())
            {
                if (failure.GetSeverity() == FailureSeverity.Warning)
                {
                    failuresAccessor.DeleteWarning(failure);
                    continue;
                }

                hasError = true;
                string text;
                try { text = failure.GetDescriptionText(); }
                catch { text = null; }
                if (!string.IsNullOrEmpty(text) && !Errors.Contains(text))
                    Errors.Add(text);
            }

            return hasError
                ? FailureProcessingResult.ProceedWithRollBack
                : FailureProcessingResult.Continue;
        }

        /// <summary>収集したエラーを 1 行にまとめる。</summary>
        public string JoinErrors()
        {
            return string.Join("\n", Errors);
        }
    }
}

using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// インポート時のトランザクション失敗ハンドラ。
    /// - 警告は自動で無視（削除）してコミットを妨げない。
    /// - 「無視できないエラー」（例: 部屋の高さは0より大きい必要がある）は
    ///   原因要素の ID を収集した上でトランザクションをロールバックする。
    ///   → 呼び出し側は失敗要素を除外して再実行し、有効な要素だけを確定できる。
    /// </summary>
    public class ImportFailurePreprocessor : IFailuresPreprocessor
    {
        /// <summary>エラーの原因となった要素ID（long 化）</summary>
        public HashSet<long> FailingElementIds { get; } = new HashSet<long>();

        /// <summary>致命的（無視できない）エラーが発生したか</summary>
        public bool HadError { get; private set; }

        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            bool hasError = false;

            foreach (var failure in failuresAccessor.GetFailureMessages())
            {
                var severity = failure.GetSeverity();
                if (severity == FailureSeverity.Warning)
                {
                    // 警告はコミットを止めないよう削除
                    failuresAccessor.DeleteWarning(failure);
                    continue;
                }

                // エラー（無視不可）→ 原因要素を記録
                hasError = true;
                foreach (var id in failure.GetFailingElementIds())
                    FailingElementIds.Add(ElementIdToLong(id));
            }

            if (hasError)
            {
                HadError = true;
                // このトランザクションはロールバック（呼び出し側が失敗要素を除外して再試行）
                return FailureProcessingResult.ProceedWithRollBack;
            }

            return FailureProcessingResult.Continue;
        }

        private static long ElementIdToLong(ElementId id)
        {
            if (id == null) return -1;
#if REVIT2026
            return id.Value;
#else
            return id.IntValue();
#endif
        }
    }
}

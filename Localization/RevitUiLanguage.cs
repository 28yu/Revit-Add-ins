namespace Tools28.Localization
{
    /// <summary>
    /// Revit 本体の UI 言語を Loc の言語コード（JP/US/CN）に対応付けるヘルパー。
    ///
    /// 【なぜアドインの言語設定と分けるのか】
    /// このアドインには性質の異なる「言語」が2つある。
    ///   - Loc.CurrentLang        … アドインの UI（ボタン・ダイアログ・エラー）の言語。ユーザーが切り替える。
    ///   - Revit 本体の UI 言語    … モデルに保存される文字列の言語。
    /// シート名のように「モデルに保存され、そのモデルを開く全員が見る」文字列は、
    /// カテゴリ名や既定ビュー名など Revit が生成する他の文字列と揃える必要があるため、
    /// アドインの言語設定ではなく Revit 本体の言語に合わせる。
    ///
    /// ⚠ ただし「モデルに保存され、かつ再実行時の検索キーになる文字列」
    ///   （フィルタ名・凡例ビュー名・共有パラメータ名など）は多言語化してはならない。
    ///   言語が変わると既存要素を見つけられず、重複作成や更新失敗を起こす。
    ///   詳細は CLAUDE.md「文字列の3分類」を参照。
    /// </summary>
    internal static class RevitUiLanguage
    {
        /// <summary>Revit の LanguageType を Loc の言語コードへ変換する。</summary>
        public static string Resolve(Autodesk.Revit.ApplicationServices.LanguageType language)
        {
            switch (language)
            {
                case Autodesk.Revit.ApplicationServices.LanguageType.Japanese:
                    return "JP";
                case Autodesk.Revit.ApplicationServices.LanguageType.Chinese_Simplified:
                    return "CN";
                default:
                    // 辞書は JP/US/CN の3つのみ。繁体字を含むその他の言語は英語にフォールバックする。
                    return "US";
            }
        }

        /// <summary>コマンドから使う簡易版。取得に失敗した場合は JP を返す。</summary>
        public static string Resolve(Autodesk.Revit.UI.UIApplication uiApp)
        {
            try
            {
                if (uiApp != null && uiApp.Application != null)
                    return Resolve(uiApp.Application.Language);
            }
            catch { }
            return "JP";
        }

        /// <summary>Document から使う簡易版（doc.Application は ApplicationServices.Application）。</summary>
        public static string Resolve(Autodesk.Revit.DB.Document doc)
        {
            try
            {
                if (doc != null && doc.Application != null)
                    return Resolve(doc.Application.Language);
            }
            catch { }
            return "JP";
        }
    }
}

## 最終セッション: 2026-07-24
新機能 **ParameterCleanup（パラメータ整理）** を追加。プロジェクト内の削除可能な
パラメータ（プロジェクト/共有/グローバル）を一覧化し、同名特定・値の有無の自動判定・
不要パラメータ削除を行う。Excel風の列並べ替え/フィルターも実装。


# 開発ステータス

> このファイルはセッション終了時に更新すること

## 現在作業中
ParameterCleanup（パラメータ整理）機能（動作確認完了・改善継続中）
ブランチ: `claude/revit-addin-parameter-cleanup-hdss4v`

### 完了（2026-07-24 ParameterCleanup セッション）
- [x] 新機能 ParameterCleanup を追加（`Commands/ParameterCleanup/`）。リボン「パラメータ」パネル＞「パラメータ整理」
  - Command / ParameterScanner / ParamRow / ParameterCleanupDialog（WPF DataGrid）
  - 対象: プロジェクト/共有パラメータ（`ParameterElement`/`SharedParameterElement`）＋グローバルパラメータ
- [x] 大容量モデルのフリーズ回避設計（列挙は軽量／値判定はカテゴリ限定・キャッシュ・early-exit／Stopwatch で約50ms毎に UI へ制御を返し進捗＋中止対応）
- [x] バインド解決を `get_Item` から `ForwardIterator`（名前キー）へ変更（多数が誤って「対象外」になる不具合を修正）
- [x] 値判定の意味を明確化：「空（未使用）」＝バインド済みだが全要素で値なし／「バインドなし」＝カテゴリ未バインド。値セルにツールチップで説明
- [x] 値の有無をダイアログ表示時に自動確認（削除後も自動再確認）。確認後に「値あり◯/空◯」サマリー表示
- [x] 集計表参照列を追加（`ScheduleField.ParameterId` で軽量取得。フィルタ/タグ/数式はAPI制約により対象外）
- [x] 種別ラジオ（すべて/プロジェクト/共有/グローバル）で絞り込み
- [x] Excel風の列メニュー（各見出し「▾」）：昇順/降順並べ替え＋値チェックリストで絞り込み（検索・全選択/選択解除・長い値は…省略＋ツールチップ・幅上限420）
- [x] 一覧下に「全選択」「選択解除」ボタン（フィルターで表示中の行のみ対象に削除用チェックを一括操作）
- [x] 行間の水平罫線を薄いグレー（#ECECEC）に
- [x] 3言語（JP/EN/CN）対応、`Docs/features.json`（added_in 2.2）・マニュアル `Docs/Features/ParameterCleanup.md` 追加
- ⚠️ ビルド時のハマりどころ（DEVLOG参照）: `Window` のインスタンスプロパティと同名の列挙型（`Visibility`/`HorizontalAlignment`/`VerticalAlignment`）は CS0176、`TextBox` は `Autodesk.Revit.UI.TextBox` と衝突し CS0104 → いずれも完全修飾で解決

---

### 旧・現在作業中（参考）: Excel エクスポート／インポートの改善
ブランチ: `claude/excel-export-improvements-r2q928`

### 完了（2026-07-23 セッション）
- [x] エクスポートダイアログのパラメータ欄に `I-`（インスタンス）/`T-`（タイプ）の凡例を追加（背景色なし・2行表示、JP/EN/CN）
- [x] 大容量モデルでパラメータのチェック／ホバーが重い問題を解消（グループ化時のUI仮想化を有効化）
- [x] エクスポートの高速化（タイプ値キャッシュ・`LookupParameter`・列幅のインライン集計・オートフィルタ範囲直接指定）
- [x] 書き出したExcelのヘッダー行（1行目）を固定（`FreezeRows(1)`）
- [x] 「カテゴリ毎にシートを分ける」→「出力Excelをカテゴリ毎にシートに分ける」に文言改善
- [x] インポートの高速化（`LookupParameter`・`ImportFromPreview` で変更セルのみ書込み・タイプ値キャッシュ）
- [x] 開いているExcelの色付けフリーズを解消（`MarkCellsViaCom` を `Range.Value2` 一括読取に・`EnableEvents=false`）
- [x] インポート読込の二重オープン解消（`GeneratePreview` にシート名 out 版を追加）
- [x] 設定ファイル(.json)読込の高速化（`ApplySettings` のパラメータ二重取得を解消・HashSet/Dictionary照合）
- [x] AutoBuild: コミット件名マーカー `[build:XXXX]` で対象Revitバージョンを切替可能に
- [x] AutoBuild: `RestartAutoBuild.ps1`（停止＋再起動を1コマンド）を追加

### 完了 (本セッション)
- [x] ワークセット '28Tools_型枠' の「全ビューに表示」チェックが入る問題を修正
  - `IsWorksetVisible(wsId)` は新規WSでも `false` を返す（UI上はチェック入り）→ガード削除、常に `SetWorksetVisibility(false)` を呼ぶ
  - 既存WS・新規WS両方で毎回設定するよう変更
- [x] Excel エクスポート時 `System.Runtime.CompilerServices.Unsafe` エラーを修正
  - `OnAssemblyResolve` が全 `System.*` をスキップしていた → ClosedXML NuGet 依存DLLをホワイトリスト化
- [x] 複数ビュー選択時のインデックス不一致修正（防御的実装）
  - `sourceViews[i]` → `perViewSources[i]` に統一
- [x] EnableSectionBox の空BBox問題修正（第1次）
  - `IsSectionBoxActive=True` を BBox なしで呼ばないよう変更
- [x] IsSectionBoxActive=False ソースビューに要素BBoxから切断ボックスを算出・設定（第2次）
  - 修正前: ソースに切断ボックスなし→解析ビューも切断ボックスなし→全体表示でDSが極小
  - 修正後: EnableSectionBox を呼んで型枠要素のBBoxから切断ボックスを算出

### ⚠️ 未解決（次セッションで継続）
**問題**: 複数ビュー選択実行時、1つの3Dビューに型枠が表示されない
- 対象: `**型枠：工作物擁壁`（IsSectionBoxActive=False のソースビュー）
- 試みた修正: 3回（上記参照） → ユーザー確認でまだ改善されず
- 最新ログ（第2回提出）から判明した事実:
  - DS は正常に413個作成されている
  - ワークセット可視性: Visible ✓
  - 要素非表示カウント: 2553件（= 2034+519、正確）✓
  - OST_GenericModel: 可視 ✓
  - フィルタ: 9件適用（壁・スラブ可視）✓
  - 切断ボックスなし → 全体表示 → DSが極小に見える（第2次修正で対処済みのはず）
- **残課題**: 第2次修正後もユーザーが改善なしと報告 → 別の原因が存在する可能性
  - 仮説A: 3Dビューの切断ボックスはビューローカル座標系だが、EnableSectionBox はワールド座標で計算している → 回転ビューでズレる可能性
  - 仮説B: EnableSectionBox が正しく動作していても、ビューポートのカメラ向きが型枠と合っていない
  - 次セッションでログ（3回目提出）を送ってもらい確認する

### デバッグログの送り方
ローカル PC で `.\Send-FormworkLog.ps1` 実行 → `.diag/Formwork_debug*.txt` が push される。Claude は `/home/user/Revit-Add-ins/.diag/Formwork_debug.txt` を Read。

または: `C:\temp\Formwork_debug.txt` と `C:\temp\Tools28_debug.txt` を直接ドラッグ&ドロップ。

## 直近の意思決定
- 2026-05-21: 第2次修正でEnableSectionBoxを再有効化（IsSectionBoxActive=False でも要素BBoxから算出）
- 2026-05-15: CLAUDE.md をスリム化し STATUS.md / TASKS.md / Docs/DEVLOG.md に分離
- 2026-05-15: FormworkCalculator は1シートにプロジェクト全体の分析ビュー+集計表を集約

## 既知のブロッカー・注意事項
- `build/test` ブランチへの push は 403 になるため `build/test2` を使う
- ローカルPCで `git pull` 後は `git log --oneline -3` で反映確認すること
- ClosedXML の AssemblyResolve ハンドラ: System.* の一律スキップは禁止（ホワイトリスト方式）

## 実装済み機能の状態

| 機能 | 状態 | 備考 |
|------|------|------|
| GridBubble（通り芯符号） | ✅ 完了 | |
| SheetCreation（シート一括作成） | ✅ 完了 | |
| ViewCopy / SectionBoxCopy / ViewportPosition / CropBoxCopy | ✅ 完了 | |
| BeamUnderLevel（梁下端色分け） | ✅ 完了・動作確認済み | |
| BeamTopLevel（梁天端色分け） | ✅ 完了・動作確認済み | |
| RoomTagCreator（部屋タグ自動配置） | ✅ 完了 | |
| FilledRegionSplitMerge（塗潰し領域分割統合） | ✅ 完了 | |
| ExcelExportImport（Excel連携） | ✅ 完了・全バージョン確認済み | |
| FireProtection（耐火被覆色分け） | ✅ 完了・動作確認済み | |
| FormworkCalculator（型枠数量算出） | 🔧 不具合修正中 | 複数ビュー時1ビュー空白問題 |
| LanguageSwitch（多言語切替） | ✅ 完了 | |
| ExpiryManager（バージョン有効期限） | ✅ 実装済み(v2.1〜) | |

## 現在のバージョン
- 最新リリース: **v2.1**（2026-05-25, `a073cb8`）
- 有効期限: 2027-06-01（`Licensing/ExpiryManager.cs`）
- 次回リリース予定: v2.2（`release/v2.2` ブランチは既に remote に存在）

## リリース履歴
| バージョン | リリース日 | コミット | 主な変更 |
|-----------|----------|---------|---------|
| v1.0 | 2026-02-04 | - | 初版（GridBubble / SheetCreation / View系コピペ） |
| v2.0 | 2026-03-27 | - | 部屋タグ自動配置・塗潰し分割統合・梁色分け・Excel連携 |
| v2.1 | 2026-05-25 | `a073cb8` | 耐火被覆色分け・型枠数量算出・有効期限管理 |

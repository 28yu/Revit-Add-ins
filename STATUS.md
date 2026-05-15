## 最終セッション: 2026-05-15 04:28
変更ファイル: .claude/update-status.sh,__test.tmp

# 開発ステータス

> このファイルはセッション終了時に更新すること

## 現在作業中
FormworkCalculator の複数3Dビュー対応 & 不具合修正
ブランチ: `claude/improve-formwork-calculation-OoBuI` (= `build/test2`)
最新コミット: `1292a36` ([1]-1 ParamSourceView バインド漏れ修正 / [6] 選択のビュー追加)

### 完了 (本セッション)
- [3] ダイアログ文言変更 (シートを作成 / 集計表を作成) - 3言語対応
- [4]-2 ビューポートタイトル非表示: `VIEWPORT_ATTR_SHOW_LABEL=0` のタイプを最優先、なければ複製作成
- [5] リンクモデルを切断ボックスでクリップ (`_sectionBoxSolid` × `BooleanOperationsType.Intersect`)
- [6] ダイアログ計算範囲に「選択のビュー」追加 (3D ビュー選択時に自動選択)
- ワークセット非表示リンクの除外 (`activeView.GetWorksetVisibility(rli.WorksetId)`)

### 未確認 (次セッションでユーザー確認待ち)
- [1] 複数3Dビュー選択 → 1シート集約 (基盤実装済み、[1]-1 修正後の動作確認待ち)
- [2] 特定ビュー再実行で他ビューはそのまま (per-view クリーンアップ実装済み、[1] 修正後に確認予定)

### 重要な実装メモ
- **DirectShape のソースビュー識別**: 共有パラメータ `ParamSourceView` (`28Tools_Formwork_ソースビュー`) で各 DirectShape にソースビュー名を保存
- **重要な落とし穴**: `FormworkParameterManager.IsAllBound()` に新パラメータの存在確認を必ず追加すること (旧バージョン文書で新パラメータがバインドされず全フィルタ無効化される)
- **シート再構築**: `FormworkSheetCreator.CreateOrUpdateSheet` が既存シートを削除→再作成 (シート名・番号は引継ぎ)。シート上の配置は `CollectAllAnalysisViewIds` + `CollectAllPerViewScheduleIds` でプロジェクト内全分析ビュー・集計表を拾う
- **HideAllFormworkShapesInOtherViews**: 全ループ後に呼ぶポストパス。`ParamSourceView` でグルーピングし、各群を対応分析ビュー以外で `HideElements`。新規分析ビューに過去ビューの DirectShape が紛れ込むのを防止
- **エンジン Scope 正規化**: `SelectedViews` モードはループ内で各ビューを `CurrentView` 扱いするため `CloneWithScope(settings, CurrentView)` でエンジンに渡す

### デバッグログの送り方
ローカル PC で `.\Send-FormworkLog.ps1` 実行 → `.diag/Formwork_debug*.txt` が push される。Claude は `/home/user/Revit-Add-ins/.diag/Formwork_debug.txt` を Read。

## 直近の意思決定
- 2026-05-15: CLAUDE.md をスリム化し STATUS.md / TASKS.md / Docs/DEVLOG.md に分離
- 2026-05-15: FormworkCalculator は1シートにプロジェクト全体の分析ビュー+集計表を集約 (1プロジェクト1シート方式)

## 既知のブロッカー・注意事項
- `build/test` ブランチへの push は 403 になるため `build/test2` を使う
- ローカルPCで `git pull` 後は `git log --oneline -3` で反映確認すること

## 実装済み機能の状態

| 機能 | 状態 | 備考 |
|------|------|------|
| GridBubble（通り芯符号） | ✅ 完了 | |
| SheetCreation（シート一括作成） | ✅ 完了 | |
| ViewCopy / SectionBoxCopy / ViewportPosition / CropBoxCopy | ✅ 完了 | |
| BeamUnderLevel（梁下端色分け） | ✅ 完了・動作確認済み | |
| BeamTopLevel（梁天端色分け） | ✅ 完了・動作確認済み | |
| RoomTagCreator（部屋タグ自動配置） | ✅ 完了 | リボン登録要確認 |
| FilledRegionSplitMerge（塗潰し領域分割統合） | ✅ 完了 | |
| ExcelExportImport（Excel連携） | ✅ 完了・全バージョン確認済み | |
| FireProtection（耐火被覆色分け） | ✅ 完了・動作確認済み | |
| FormworkCalculator（型枠数量算出） | ✅ 完了 | 一部特殊形状で接触面検出漏れあり |
| LanguageSwitch（多言語切替） | ✅ 完了 | Revit環境でのテストが必要 |
| ExpiryManager（バージョン有効期限） | ✅ 実装済み(v2.1〜) | |

## 現在のバージョン
- 最新リリース: 確認が必要（GitHub Releases 参照）
- 有効期限: ExpiryManager.cs の ExpiryDate 確認

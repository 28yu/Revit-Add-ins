## 最終セッション: 2026-05-25
v2.1 リリース完了（`release/v2.1` を `a073cb8` に force-push → GitHub Actions が起動 → `v2.1` タグ・配布ZIP 6本生成）。


# 開発ステータス

> このファイルはセッション終了時に更新すること

## 現在作業中
FormworkCalculator の 3D ビュー不具合修正（継続中）
ブランチ: `claude/fix-3d-view-issues-vqjlN`
最新コミット: `08c185f` (型枠ビューが空になる問題を修正：切断ボックスなしソースビューにBBoxを算出して設定)

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

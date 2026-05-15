## 最終セッション: 2026-05-15 04:28
変更ファイル: .claude/update-status.sh,__test.tmp

# 開発ステータス

> このファイルはセッション終了時に更新すること

## 現在作業中
（なし - 次のタスクは TASKS.md 参照）

## 直近の意思決定
- 2026-05-15: CLAUDE.md をスリム化し STATUS.md / TASKS.md / Docs/DEVLOG.md に分離

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

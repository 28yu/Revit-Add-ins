# タスクリスト・意思決定ログ

## TODO

### 優先度: 高
（完了済み項目なし）

### 優先度: 中
- [ ] 28tools-download リポジトリとの連携方法の検討
- [ ] FormworkCalculator: 特殊形状の接触面検出漏れ改善

### 優先度: 低
- [ ] Phase 2 有償化（ライセンスキー方式）の実装（有償化時）

## 完了済み
- [x] 2026-05-15: CLAUDE.md スリム化・ドキュメント分離
- [x] 2026-05-15: Stop hook で STATUS.md 自動更新を設定
- [x] 2026-05-15: アイコンジェネレーターを Docs/tools/ に移動し GitHub Pages 公開
- [x] 2026-05-15: LanguageSwitch: Application.cs に登録済みで対応完了
- [x] 2026-05-15: RoomTagCreator: Application.cs へのリボン登録済みを確認
- [x] 2026-05-15: 未使用の IconGenerator.html を削除

## 意思決定ログ

| 日付 | 決定事項 | 理由 |
|------|---------|------|
| 2026-05-15 | CLAUDE.md を分割（STATUS/TASKS/DEVLOG）| セッション開始のトークン消費削減 |
| 2026-05-15 | Web版 Claude Code 継続（CLI非導入） | 会社PC制約・マルチデバイス対応 |
| 2026-05-14 | 集計表列幅計算を係数2.6+12mmに変更 | Revit実描画に合わせるため |
| 2026-05-07 | FormworkCalculatorのマテリアルベース算出を中止 | CFT等の誤判定リスクのため |

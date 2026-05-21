# タスクリスト・意思決定ログ

## TODO

### 優先度: 高
- [ ] **FormworkCalculator: 複数ビュー選択時に1ビューが空白になる問題** (`claude/fix-3d-view-issues-vqjlN`)
  - 対象: IsSectionBoxActive=False のソースビュー（例: `**型枠：工作物擁壁`）
  - 3回修正済みもまだ未解決（2026-05-21時点）
  - 調査ポイント1: `EnableSectionBox` の切断ボックス座標系（ワールド座標 vs. ビューローカル座標）
    - `view.GetSectionBox().Transform` を確認 → Identity でなければ座標変換が必要
    - 修正案: `BoundingBoxXYZ.Transform.Inverse.OfPoint(worldPoint)` でビューローカル座標に変換してからセット
  - 調査ポイント2: EnableSectionBox が実際に呼ばれているかログ確認（3回目ログ提出待ち）
    - ログに `EnableSectionBox: set BB` or `no elements with BBox found` が出るか確認
  - 調査ポイント3: ビューポートがシート上で正しい位置に配置されているか

### 優先度: 中
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
- [x] 2026-05-15: 28tools-download 連携を完全自動化（features.json 単一ソース化・アイコン Pages 公開・リリース本文自動生成・配布サイト動的描画）

## 意思決定ログ

| 日付 | 決定事項 | 理由 |
|------|---------|------|
| 2026-05-15 | CLAUDE.md を分割（STATUS/TASKS/DEVLOG）| セッション開始のトークン消費削減 |
| 2026-05-15 | Web版 Claude Code 継続（CLI非導入） | 会社PC制約・マルチデバイス対応 |
| 2026-05-14 | 集計表列幅計算を係数2.6+12mmに変更 | Revit実描画に合わせるため |
| 2026-05-07 | FormworkCalculatorのマテリアルベース算出を中止 | CFT等の誤判定リスクのため |

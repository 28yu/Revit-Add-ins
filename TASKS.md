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

### RenderKit（Revit → Blender → ブラウザで見る 360°パノラマ／個人用）
- [ ] **第1段階の実機確認**: Revit の FBX を `RenderKit/MakePanorama.bat` に渡して確認（2026-10-15 作成・未確認）
  - 確認点: 単位（大きさのログが建物として自然か）、パースビューのカメラが FBX に入るか、
    マテリアル名でガラス・金属が判定されるか、GPU が自動で使われるか、Windows での文字化け
- [ ] 第2段階: 歩き回れる 3D（Blender で光を焼き込み → glTF → HTML1つの three.js ビューア）
- [ ] 第3段階: Tools28 に glTF 書き出しボタン（`CustomExporter`。素材・テクスチャ・カメラ・太陽を正しく渡す）
- [ ] 第4段階: HTML を exe に包む（WebView2）

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
| 2026-10-15 | 個人用レンダリングは Blender を描画エンジンにし、結果を HTML1つ（外部ライブラリなし）で配る（`RenderKit/`） | 描画エンジンの自作は非現実的。HTML1つならどの PC・スマホでもアプリ不要で見られる |
| 2026-05-15 | CLAUDE.md を分割（STATUS/TASKS/DEVLOG）| セッション開始のトークン消費削減 |
| 2026-05-15 | Web版 Claude Code 継続（CLI非導入） | 会社PC制約・マルチデバイス対応 |
| 2026-05-14 | 集計表列幅計算を係数2.6+12mmに変更 | Revit実描画に合わせるため |
| 2026-05-07 | FormworkCalculatorのマテリアルベース算出を中止 | CFT等の誤判定リスクのため |

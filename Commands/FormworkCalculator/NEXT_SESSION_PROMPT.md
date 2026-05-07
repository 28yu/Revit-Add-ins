# 次セッション開始プロンプト（型枠数量算出 開発再開用）

このファイルは新しい Claude Code セッションを始める際に使う指示文です。
新セッションで、最初のメッセージとしてこの内容をコピペしてください。

---

## 📋 コピペする指示文（ここから）

型枠数量算出アドイン (Tools28) の開発を続けます。

**前提情報**:
- 作業ブランチ: `claude/formwork-addon-development-433Z5`
- 主要ファイル: `Commands/FormworkCalculator/`
- 詳細仕様: `Commands/FormworkCalculator/HANDOFF.md` の冒頭「次セッションへの引き継ぎ事項」を必ず最初に読んでください
- 共通仕様 (CLAUDE.md): プロジェクト全体ルール
- デバッグログ: `C:\temp\Formwork_debug.txt`

**前回セッションで完了した機能** (本番投入可能):
1. 鉄骨部材の自動除外 (4 層判定: StructuralMaterialType / 形状 / Material / 名前)
2. デッキスラブの自動除外 (タイプ名 "DS")
3. 壁スイープ・リビールの自動除外
4. 動的合計サマリ集計表「型枠数量集計_合計」
5. Excel 出力の CJK 対応列幅
6. 集計表総合計と Excel 総括表の数値一致
7. パラメータ自動候補 ComboBox (工区・型枠種別)
8. 詳細診断ログ ([ElemDiag] / [FaceDiag] / [Pair])

**今回取り組む課題**:

[53] **ContactFaceDetector の精度問題**

ユーザー報告の症状:
- 柱の一部接触で全面除外 (Stage 1 過大マッチ)
- 構造フレーム T 字接合で接触部が省かれない
- 構造基礎の床接触面が残る
- 部分接触 Yes で面積 0 になるケース

最重要ケース: **柱 E3280907** (950×950×940mm スタブ柱)
- 梁 E3286646 と 4 側面が Stage 1 (Full Contact) マッチ
- ペアログで `aArea=9.61 bArea=8.28 d=0.0 FULL_CONTACT` となっている
- スクショ確認: 梁の上に乗るスタブ柱 (本来は 4 側面が露出 = formwork 必要)
- 仮説: 梁が L 字/U 字で柱を囲んでいる、または Revit Join Geometry の影響で側面同士が幾何的に「距離 0」になっている

**修正方針 (次セッションで検討・実装)**:

1. Stage 1 (`ContactFaceDetector.EvaluateContact`) の判定強化:
   - 現状: 距離 0 + 面積比 ≤ 1.2 + a 中心が b 内
   - 追加検討: a の 4 隅を b に投影して**重なり領域**を計測
     - 既に `ProjectBCornersToA` で b → a の投影を実装済み
     - 対称的に a → b の投影を追加し、b 内に入る隅の数で判定強化
   - 重なり領域が a の面積の 80% 未満なら Stage 1 拒否、Stage 2 (Partial) へ流す

2. 上記でも誤検出する場合: Solid 体積比較で「片方が他方に内包されているか」をチェック

3. 既存の `[Pair]` ログを活用して各ケースの修正前後を比較検証

**進め方**:
1. まず `Commands/FormworkCalculator/HANDOFF.md` の「次セッションへの引き継ぎ事項」を Read で読む
2. `Commands/FormworkCalculator/Engine/ContactFaceDetector.cs` の Stage 1 ロジックを読む
3. 修正案を提示してユーザー (私) に確認
4. 承認後に実装 → コミット → プッシュ
5. ユーザー側でビルド・動作確認 → ログ分析 → 反復

**コミット・プッシュ規則** (CLAUDE.md より):
- コミットメッセージは日本語
- `git commit` と `git push` は別の Bash 呼び出しに分ける
- PreToolUse hook が push 前に自動 rebase

最初のステップとして HANDOFF.md を読んで現状を把握してください。

## 📋 ここまで（コピペ範囲）

---

## 📁 参考: 関連ファイル一覧

開発で頻繁に参照するファイル:

```
Commands/FormworkCalculator/
├── HANDOFF.md                           # 開発記録・引き継ぎ (最初に読む)
├── FormworkCalculatorCommand.cs         # エントリポイント
├── Engine/
│   ├── ContactFaceDetector.cs           # ⚠️ 次セッションの主修正対象
│   ├── PartialContactClipper.cs         # 部分接触の 2D 矩形差分
│   ├── FaceClassifier.cs                # 面分類
│   ├── FormworkCalcEngine.cs            # 3 Pass パイプライン + 診断ログ
│   ├── ElementCollector.cs              # 鉄骨/デッキスラブ振り分け
│   ├── SteelMemberDetector.cs           # 鉄骨判定 (4 層)
│   ├── DeckSlabDetector.cs              # デッキスラブ判定
│   ├── WallSweepFaceDeductor.cs         # 壁スイープ host 面控除
│   ├── FormworkParameterManager.cs      # 共有パラメータ
│   ├── FormworkFilterManager.cs         # View Filter 色分け
│   └── ParameterCandidateScanner.cs     # パラメータ候補列挙
├── Models/
│   ├── FormworkSettings.cs
│   └── FormworkResult.cs                # ExclusionKind enum
├── Output/
│   ├── ScheduleCreator.cs               # 集計表 + サマリ集計表
│   ├── ExcelExporter.cs                 # CJK 対応列幅
│   └── FormworkVisualizer.cs            # 3D ビュー DirectShape
└── Views/
    ├── FormworkDialog.xaml              # ComboBox 化済み
    └── FormworkDialog.xaml.cs
```

CLAUDE.md (プロジェクトルート) には Revit 2022 Schedule API の制約まとめあり。

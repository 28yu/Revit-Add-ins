# 次回セッション用ハンドオフ — リンク壁スイープ問題

最終更新: 2026-05-14

## 🔴 未解決の優先課題

### リンクモデルの壁スイープ面の型枠 DirectShape が作成されない

**症状**: リンクモデル内の壁に取り付いている **壁スイープ (additive WallSweep)** の外側面（前/上/下/端部）に対する型枠 DirectShape が 3D ビュー「型枠分析」に作成されていない。ホストモデルの壁スイープは正しく作成される（ユーザー証言）。

**ユーザーの最新報告**:
> 動作確認したところ、まだリンクモデルの壁のスイープ面の型枠オブジェクトが作成されない問題が解決していませんでした。

⚠️ **重要**: 当初ユーザーは「リビール」と言っていたが、実際は「**スイープ (additive sweep)**」だった。リビール (subtractive reveal) ではない。

---

## ✅ 既に試したこと（このセッションで実施）

| 試行 | 結果 |
|---|---|
| ElementSource + Registry を新設しリンク要素を取り扱う基盤を構築 | 基盤完成 |
| `RevitLinkInstance.GetTotalTransform()` + `SolidUtils.CreateTransformed` で世界座標化 | 動作 |
| 共有パラメータ `28Tools_Formwork_ソース` 追加 | 動作 |
| 集計表をソース別に分離 (`型枠数量集計` / `型枠数量集計 - ホスト` / `型枠数量集計 - 構造_A棟` 等) | 動作 |
| サマリ集計表「型枠数量集計_合計」にソース別小計 + 全体合計 | 動作 |
| リンクの WallSweep を収集対象に含める | 効果なし（症状継続） |
| `MoveWallSweepsToExcluded` の呼び出しを除去（sweep を ElementResult として保持） | 効果なし（症状継続） |
| `WallSweepFaceDeductor` でソースマッチング (host↔host, link X↔link X のみ) | 動作 |

---

## 🤔 仮説（次セッションで検証すべき）

### 仮説 A: リンクの WallSweep の `get_Geometry()` が空を返す
- WallSweep は付帯要素であり、リンクドキュメント経由では Solid を持たない可能性
- `SolidUnionProcessor.GetSolids(linkedSweep, transform)` の返却が空 list
- → `ClassifyElementFaces` が null を返し、context が作成されない
- → ElementResult も作られない → DirectShape も作られない
- **検証方法**: `C:\temp\Formwork_debug.txt` で `[LinkedCollect]` 以降の Pass 1 ログを確認し、リンクの sweep が context として認識されているかチェック

### 仮説 B: リンクの WallSweep の Solid が変換時に消失
- `SolidUtils.CreateTransformed(solid, transform)` が WallSweep の薄い solid に対して失敗
- 鏡像変換 (Determinant < 0) だと Solid が反転して空になる場合あり
- **検証方法**: `SolidUnionProcessor.GetSolids(elem, transform)` で transformed.Count をログ出力

### 仮説 C: 集計表のフィルタが過度に絞り込み
- `ParamSource == リンク名` で絞ったが、リンク要素の DirectShape の `ParamSource` が正しく設定されていない
- **検証方法**: 作成された DirectShape をクリックしてプロパティパネルで `28Tools_Formwork_ソース` の値を確認

### 仮説 D: そもそもリンクの WallSweep がコレクションされていない
- `CollectFromDoc` が WallSweep を取りこぼしている
- `FilteredElementCollector(linkDoc).OfClass(typeof(WallSweep))` がリンク doc で機能するか不明
- **検証方法**: `[LinkedCollect]` ログで要素タイプ別カウントを出すように改修

### 仮説 E: ContactFaceDetector がリンクスイープの全面を deduct している
- リンクスイープと隣接するリンク壁が極めて密着 (0距離) → 全 sweep 面が contact 判定 → 全 DeductedContact
- 結果: FormworkRequired 面ゼロ → DirectShape なし
- **検証方法**: Pass 2 ログでリンク sweep の face type 内訳を確認

### 仮説 F: 「現在のビュー」スコープでリンクの WallSweep がフィルタアウトされている
- 「現在のビュー」モードではリンクインスタンスは `IsHidden(view)` 判定するが、WallSweep 個別のビュー可視性は考慮していない
- ただし `CollectFromDoc(linkDoc, settings, null, isLinked: true)` で `useViewFilter = false` になるので、リンク内では全 WallSweep が取れるはず
- **検証方法**: 「プロジェクト全体」モードでも同じ問題が起きるかユーザーに確認

---

## 🎯 次セッションでの推奨アプローチ

### ステップ 1: 詳細診断ログを追加してユーザーに実行してもらう

リンクの WallSweep について各 Pass でログを強化:
- **Pass 1**: solid 取得時点での solid 数、Volume、face 数
- **Pass 1**: face 分類後の各 FaceType の数 (FormworkRequired / DeductedTop / DeductedContact / etc.)
- **Pass 2**: contact 検出後の face type 変化（before/after の差分）
- **Pass 3**: ElementResult.FormworkArea
- **DirectShape 作成時**: 生成された ID と書き込んだ `ParamSource` 値

これらは `[LinkSweepDiag]` のような独自タグで出力し検索しやすくする。

### ステップ 2: ログから根本原因を特定

得られたログから仮説 A〜F のどれが該当するか判定:
- solid 数 = 0 → 仮説 A
- solid 数 > 0 だが transformed 後 = 0 → 仮説 B
- Pass 1 で face = 0 → solid 自体に face なし
- Pass 2 で FormworkRequired → 全て DeductedContact になっている → 仮説 E
- DirectShape は作成されているが集計表に出ない → 仮説 C

### ステップ 3: 該当仮説に応じた修正

**仮説 A の場合**: WallSweep の geometry 取得方法を変更
- `wallSweep.get_Geometry()` 以外の手段を調査
- host の wall solid と sweep の boundingBox から sweep 領域の solid を逆生成
- もしくは sweep の親 wall の solid から sweep 部分を切り出す

**仮説 B の場合**: transform 適用方法を改善
- 鏡像チェック `xform.Determinant > 0` で分岐
- 失敗時のフォールバック実装

**仮説 E の場合**: ContactFaceDetector の調整
- リンク間の sweep⇔wall の contact 検出を抑制
- もしくは sweep 外側面（壁から離れた向き）を保護する判定追加
- 例: 面の中心点から壁本体の中心への方向ベクトルと面法線の内積をチェック

---

## 📁 関連ファイル早見表

| ファイル | 役割 |
|---|---|
| `Engine/ElementSource.cs` | リンク要素ラッパー + レジストリ |
| `Engine/ElementCollector.cs` | 要素収集 (ホスト + リンク)、`CollectFromLinkedModels` |
| `Engine/SolidUnionProcessor.cs` | Solid 取得 + Transform 適用 (`GetSolids(elem, transform)`) |
| `Engine/FormworkCalcEngine.cs` | パイプライン本体 (Pass 1〜3) |
| `Engine/ContactFaceDetector.cs` | 接触面検出 |
| `Engine/WallSweepFaceDeductor.cs` | sweep 近傍 wall face 控除 (ソースマッチ済) |
| `Engine/FormworkParameterManager.cs` | 共有パラメータ管理 (`ParamSource` 含む) |
| `Output/FormworkVisualizer.cs` | DirectShape 作成 |
| `Output/ScheduleCreator.cs` | 集計表作成 (ソース別分離対応) |
| `FormworkCalculatorCommand.cs` | エントリポイント、per-source schedule 作成ループ |
| `Views/FormworkDialog.xaml.cs` | ダイアログ (`ChkIncludeLinks` 含む) |
| `Models/FormworkSettings.cs` | 設定 (`IncludeLinkedModels` 含む) |
| `Models/FormworkResult.cs` | 結果モデル (`SourceName`、`SourceRegistry` 含む) |

---

## 🧠 重要な実装ポイント

### `ElementSource.cs`
- `SurrogateId` (整数キー)
  - ホスト要素: `Element.Id.IntValue()` を維持（正値）
  - リンク要素: 連番マイナス値（衝突回避）
- `ElementSourceRegistry`: `Dictionary<int, ElementSource>` で全要素を索引
- `HostSourceName = "ホスト"` (定数)

### `ElementCollector.CollectFromDoc`
- ホスト・リンク共通メソッド (`isLinked` フラグで挙動切替)
- 「現在のビュー」スコープはホストのみ適用（リンクはインスタンス可視性で判定）
- **WallSweep はホスト・リンク両方から収集**（最終決定）

### `FormworkCalcEngine.ClassifyElementFaces`
- `ElementSource` を受け取り `src.Element` + `src.Transform` を使用
- `SolidUnionProcessor.GetSolids(elem, src.Transform)` で transform 適用
- ワールド BoundingBox を `ComputeWorldBoundingBox` で計算
- `ctx.ElementId` には `src.SurrogateId` を入れる（リンクは負値）

### `WallSweepFaceDeductor`
- `sourceByCtxId` で各 context にソース名を付与
- 同一ソース同士のペアのみ deduct を実行 (host↔host、link X↔link X)

### `MoveWallSweepsToExcluded`
- **このセッションで呼び出しを除去** (commit `74420a1`)
- WallSweep は ElementResult として残り、外側面が FormworkRequired として DirectShape 化される設計に変更

### `ScheduleCreator.CreateSchedule`
- `sourceFilter` 引数追加
  - null → 「型枠数量集計」(リンクなし時)
  - 値あり → 「型枠数量集計 - {ソース名}」を作成、`ParamSource` フィルタ適用
- `DeleteAllFormworkSchedules` でホスト集計表作成時に旧 per-source 集計表もまとめて削除

---

## 📚 このセッションで得た知見

### Revit API 関連
1. **`RevitLinkInstance.GetLinkDocument()`** はリンクされていない/アンロード中の場合 null を返す
2. **`SolidUtils.CreateTransformed`** は Revit 2017+ で利用可能。鏡像 transform で面法線反転に注意
3. **`RevitLinkInstance.IsHidden(View)`** で「現在のビュー」スコープのリンク可視判定
4. **DirectShape はホストドキュメントのみ作成可能** — リンク要素位置にホスト側で作る
5. **`ScheduleField.IsHidden = true`** にしても **`ScheduleFilter` は機能する**（フィルタ専用フィールドとして使える）

### auto-merge ワークフローの落とし穴
- **空コミットは auto-merge で失敗する** (`git commit` 時にステージするものがなくエラー)
- 再デプロイトリガーには必ず実差分を含むコミットが必要
- 例: コメント1行を追加するなど些細な変更を入れる

### スケジュール周り
- メイン集計表は `型枠数量集計` または `型枠数量集計 - {ソース名}` の命名規則
- リンクなし or 単一ソース時は `型枠数量集計` のみ作成
- リンクあり時は `型枠数量集計 - ホスト` + `型枠数量集計 - {各リンク}` を作成
- サマリ集計表 `型枠数量集計_合計` はソース列ありの単一集計表（小計 + 全体合計）

### 既存の関連解決済み問題
- **リビールではなくスイープが本当の問題**（ユーザー訂正済み）
- リビールについては「リンク壁のリビールカットは wall solid に baked-in されている前提」で対応済み

---

## 🚀 次セッションの最初の一手

1. **このファイルを最初に読む**
2. **`Commands/FormworkCalculator/HANDOFF.md` も合わせて確認** — 既存の API 知見・残課題が記録されている
3. ユーザーに以下を依頼:
   - 最新ビルドで実行（`IncludeLinkedModels=true` でリンク含む）
   - `C:\temp\Formwork_debug.txt` を共有
   - 具体的にどのリンクファイル、どの壁の、どのスイープが対象かスクリーンショット
   - 「プロジェクト全体」スコープでも同じ症状か確認
4. ログ内容から仮説 A〜F を絞り込み、必要なら診断ログ追加修正してユーザーに再実行を依頼
5. 根本原因が特定できたら本格修正

### 診断ログ追加コードの例

```csharp
// ClassifyElementFaces 冒頭で
if (src.IsLinked && src.Element is WallSweep)
{
    var rawSolids = SolidUnionProcessor.GetSolids(elem, null);  // transform前
    var transformedSolids = SolidUnionProcessor.GetSolids(elem, src.Transform);
    FormworkDebugLog.Log(
        $"  [LinkSweepDiag] E{elem.Id.IntValue()} src={src.SourceName} " +
        $"rawSolids={rawSolids.Count} (totalVol={rawSolids.Sum(s => s.Volume):F4}) " +
        $"transformedSolids={transformedSolids.Count} (totalVol={transformedSolids.Sum(s => s.Volume):F4}) " +
        $"transformDet={src.Transform.Determinant:F4}");
}
```

---

## 🌐 ブランチ情報

- **作業ブランチ**: `claude/improve-formwork-picking-SyP0Y`
- **最新コミット**: `c586c30` (コメント更新による再デプロイトリガー)
- **main 上の最新 squash 済みコミット**: `c76681a` (Auto: 型枠数量算出: 壁スイープ面の型枠を算出 + リンク別集計表の生成を強化)

> 次セッションでは新規ブランチを切り直すか、同ブランチを継続使用するか判断してください。同ブランチ継続なら `git pull --rebase origin main` で main に追従してから作業開始。

---

## 💡 補足: 一連のユーザーやり取りタイムライン

1. リンクモデル対応の要望 → 全パイプライン対応実装
2. 「リビール面の型枠が算出されない」報告 → リンク WallSweep 関連を調整
3. 集計表でソース別表示の要望 → サマリ集計表に小計実装
4. 「リビールではなくスイープだった」訂正 → `MoveWallSweepsToExcluded` 除去で対応試行
5. **「まだリンクのスイープ面が作成されない」← 次セッション開始地点**
6. 「集計表はホスト/リンクで分離・ソース列削除」→ 集計表を per-source 分離で実装済み

ユーザーは複数のリンクファイルで動作確認しており、ホスト壁の sweep は機能している（らしい）が、リンク壁の sweep は機能していない。

# 開発知見ログ (DEVLOG)

> このファイルは参照専用。開発中に得た知見・解決済みの問題を蓄積する。
> 各セクションはCLAUDE.mdから移動済み。

---

## 28tools-download 連携自動化（2026-05-15）

### 改善前の課題

| 課題 | 内容 |
|------|------|
| 機能情報の分散 | ボタン名・アイコン・説明・マニュアルが別々のファイルに存在 |
| リリース本文の二重管理 | `build-and-release.yml` に機能一覧を手書き（2箇所・約80行） |
| 配布サイトのアイコン | `Resources/Icons/` は Web 非公開。配布サイト側で別途用意 |
| 配布サイトの更新 | 新機能追加時に配布サイト側のコードも手動変更が必要 |

### 実施した変更

**Step A: アイコンを GitHub Pages で公開**
- `deploy-pages.yml` を改修。ビルド時に `Resources/Icons/*.png` を `_site/icons/features/` にコピー
- 公開 URL: `https://28yu.github.io/Revit-Add-ins/icons/features/<file>.png`
- トリガーに `Resources/Icons/**` を追加（アイコン更新でも Pages 再デプロイ）

**Step B: `Docs/features.json` を作成（機能カタログの単一ソース）**
- 14機能 × 7カテゴリ × 3言語のマッピングを1ファイルに集約
- 各エントリ: `id` / `category` / `icon` / `manual` / `added_in` / `names(ja/en/zh)`
- `added_in` フィールド: そのバージョンで初登場した機能に付与。リリース本文の「⭐新機能」判定に使用

**Step C: リリース本文を自動生成**
- `scripts/generate-release-body.py` を新規作成（約60行）
- `features.json` を読み、カテゴリ別機能一覧と新機能セクションを Markdown 生成
- `build-and-release.yml` の release ジョブに checkout + スクリプト実行ステップを追加
- `body:` 手書きブロック（×2箇所）を `body_path: release-body.md` に置換

**Step D: 配布サイト側の実装（別リポジトリ）**
- 実装指示書を `Docs/INTEGRATION-28tools-download.md` として作成
- 28tools-download 側で fetch → 動的カード描画 + marked.js による MD レンダリングを実装

### 改善後の運用フロー

新機能追加時にやること：
1. コマンドクラス作成
2. 多言語リソース追加（JP/EN/CN）
3. `Application.cs` にリボン登録
4. `Resources/Icons/` にアイコン追加
5. `Docs/Features/FeatureName.md` にマニュアル作成
6. **`Docs/features.json` に `added_in: "バージョン"` 付きでエントリ追加** ← 新規手順

これだけで次のリリース時に：
- GitHub Releases 本文に「⭐新機能」として自動掲載
- 配布サイトに自動でカード追加
- 配布サイトのマニュアルページから自動でアクセス可能

### 注意事項

- `features.json` の **カテゴリ ID は変更・削除禁止**（配布サイトの分類が崩れる）
- 過去バージョンで追加した機能の **`added_in` は変更禁止**
- Pages へのアイコン反映は main マージ後、数分かかる場合あり
- 配布サイト側のページ構造を大幅変更した場合は `INTEGRATION-28tools-download.md` との整合性確認が必要

---

## AutoBuild 開発知見

### 管理者権限
- `C:\ProgramData\Autodesk\Revit\Addins\` への書き込みには管理者権限が必要
- VBS の `ShellExecute "runas"` で UAC 昇格 → 管理者権限の PowerShell を起動
- 昇格後のプロセスは `C:\Windows\System32` がカレントディレクトリになるため、`Set-Location $PSScriptRoot` が必須

### PowerShell 5.1 での日本語文字化け
- Windows PowerShell 5.1 は `.ps1` ファイルをシステムロケール（Shift-JIS）で読む
- UTF-8 BOM を付けても `git reset --hard` 後に失われる場合がある
- **解決策**: 日本語文字列は Unicode エスケープで記述する
  ```powershell
  # "ビルド成功" を Unicode エスケープで
  $msg = -join([char[]]@(0x30D3,0x30EB,0x30C9,0x6210,0x529F))
  ```
- 通知ダイアログへの日本語テキスト受け渡しは JSON ファイル経由 + `[System.IO.File]::WriteAllText/ReadAllText` で UTF-8 を明示指定

### 通知ダイアログ (MessageBox)
- `Start-Process powershell -Command` で日本語を渡すと文字化け
- `-EncodedCommand`（Base64）でも日本語を含むスクリプトは失敗
- **解決策**: 日本語テキストを JSON ファイルに書き出し、`-EncodedCommand` のスクリプトは JSON を読むだけ（ASCII のみ）にする

### 外部コマンド (git) の日本語出力
- `git log` 等の出力は UTF-8 だが、PowerShell 5.1 はシステムロケール（Shift-JIS）で読み取る
- **解決策**: スクリプト冒頭で `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8` を設定

### ビルド成功判定
- `& .\QuickBuild.ps1` の `$LASTEXITCODE` は信頼できない（PowerShell スクリプト呼び出しでは正しく伝搬しない）
- **解決策**: ビルド前後の DLL タイムスタンプを比較して成功判定

---

## 解決済みバグ: 塗潰し領域ボタンの名称・アイコン変更がRevitに反映されない問題

### 症状
- ボタン名・アイコン・内部名を変更し、ビルド＆デプロイしても Revit に反映されない
- デバッグログも更新されない
- リボンの位置変更（パネル再構成）は反映されている

### 原因
**`git pull` の失敗に気づかず、古いコードでビルドしていた。**

具体的な経緯：
1. Claude Code（リモート環境）でコード変更 → コミット → push → 自動マージで main に反映
2. ユーザーのローカル（Windows）で `git pull origin main` を実行
3. **ローカルに未コミットの変更（`filled_region_32.png`）があったため `git pull` がエラーで中止された**
4. エラーメッセージを見落とし、`QuickBuild.ps1` を実行 → 古いコードのままビルド＆デプロイ
5. Revit を起動しても変更が反映されない → 様々な原因を調査（UIState.dat 削除、DLLタイムスタンプ確認等）
6. 実際にはコード自体が更新されていなかっただけだった

### 教訓・再発防止

#### 確認手順チェックリスト（変更が反映されない場合）
```
1. git status                           # 未コミット変更がないか
2. git log --oneline -5                 # 期待するコミットがHEADに含まれるか
3. git pull origin main                 # pull が成功したか（エラーなし？）
4. .\QuickBuild.ps1                     # ビルド成功を確認
5. デプロイ先DLLのタイムスタンプ確認      # 更新されているか
6. Revit 再起動して確認
```

#### git pull 失敗時の対処法
```powershell
# ローカルの未コミット変更を退避して pull
git stash
git pull origin main
git stash pop  # 必要なら退避した変更を戻す

# または、ローカル変更が不要なら破棄
git checkout -- <ファイル名>
git pull origin main
```

---

## BeamUnderLevel 開発知見

### アイコン作成
- アイコンは `Resources/Icons/{name}_32.png` の命名規則
- `Tools28.csproj` の `<Resource>` に登録が必要
- `Application.cs` の `LoadImage()` でリソースまたはファイルから読み込み（ハイブリッド方式）

### 梁ラベル (TextNote) の配置
- ビュー上の梁の位置取得には `beam.get_BoundingBox(view)` を使用（モデル座標の BoundingBox ではなくビュー固有のものを使うこと）
- 梁の幅はタイプパラメータから取得（インスタンスパラメータではない）
- ラベルのオフセット量はビュースケール (`view.Scale`) を考慮して調整する
- テキスト配置は `Center` + `Bottom` で梁との重なりを防止
- 梁の方向に合わせてラベルを回転させる

### 自動マージ (claude/* ブランチ)
- push 前の rebase は PreToolUse hook で自動実行される（`.claude/settings.json`）
- 自動マージ成功後はリモートブランチが自動削除される
- 削除後に再 push する場合は `--force-with-lease` ではなく通常の push を使用
- マージ失敗時は rebase して再 push すれば自動リトライされる

### BeamUnderLevel 設計詳細

#### 計算式
```
梁下端レベル = 階高 - 梁天端レベル - 梁高さ
```
コード上: `bottomLevel = floorHeight + topLevelOffset - beamHeight`
- `floorHeight` = 上位レベル標高 − 参照レベル標高（例: 3000mm）
- `topLevelOffset` = 梁天端パラメータ値（上位レベルからの下がりは負値。例: -300mm）
- `beamHeight` = 梁高さパラメータ値（例: 600mm）
- 結果は参照レベル基準（例: +2100mm → 参照レベルから2100mm上）

#### レベル構成
- **参照レベル**: 天井伏図の GenLevel（自動取得、変更不可）
- **上位レベル**: ユーザーが選択（参照レベルより上のレベルのみ表示）

#### ダイアログ構成（4ステップ）
1. レベル設定（参照レベル表示 + 上位レベル選択 + 階高表示）
2. 梁高さパラメータ選択（ファミリ毎）
3. 梁天端レベルパラメータ選択（ファミリ毎）
4. 処理確認・実行

#### パラメータ選択の設計
- ファミリ毎に異なるパラメータを選択可能
- 主要候補はラジオボタン（自動検出、検出数表示）
- 「その他」はComboBox（レベル・オフセット関連キーワードでフィルタしたパラメータ一覧）

#### フィルタ・色分けの設計
- **グラフィック上書き**: 投影サーフェス前景の塗り潰しのみ（断面パターン・投影線・断面線は変更しない）
- **配色**: 明るいパステル〜中間色トーンのみ使用。黒っぽい色・暗い茶色は使わない
- **フィルタ名**: `梁下_{レベル名}{±値}` 形式

---

## BeamTopLevel 開発知見

### ダイアログの設計知見
- `SizeToContent="Height"` + `MaxHeight="800"` でコンテンツに応じた自動サイズ調整（固定Heightだと隙間が生じる）
- Step1のGrid行定義で `Height="*"` を使うと不要な空間ができるため `Height="Auto"` のみにする

### アイコンのデザイン規則
- 梁下端: I型梁(上) + 上向き矢印 + ∇FL線(下) + ピンク/黄/青3色ブロック(右)
- 梁天端: ∇FL線(上) + 下向き矢印 + I型梁(下) + ピンク/黄/青3色ブロック(右)
- 色: ピンク `(255,128,148)`, 黄 `(218,185,47)`, 青 `(30,144,255)`
- Python Pillow (`ImageDraw`) で32x32 PNGを生成

---

## RoomTagCreator 開発知見

### アイコン設計
- **ファイル**: `Resources/Icons/room_tag_32.png`
- **デザイン**: 表形式アイコン（32x32 PNG、透過背景）
- **構成**: 上部1行は通し（結合セル風）、下部は3列グリッド、最下行は2列
- **色**: 黒線 `(0,0,0)` のみ、背景は透過
- 32x32の小さいアイコンでは、罫線1pxでの表現が基本
- 表形式のアイコンは「通しの行」と「分割された行」の組み合わせで表現
- ユーザーのフィードバックに応じて縦罫線の有無を調整（上部セクションは通し＝縦罫線なし）
- 透過PNGで作成し、背景色は不要

---

## FilledRegionSplitMerge 開発知見

### 統合処理のアルゴリズム（2D ブーリアン和）
単純に `FilledRegion.Create(doc, typeId, viewId, allLoops)` で全境界線を連結すると、
**領域が重なっていると Revit がエラーを出す、または重なり部分を穴として扱う**

**対策**: 各領域を薄板ソリッドに押し出し → ブーリアン和 → 上面の境界ループを取得

1. 最初の外形ループから `GetPlane().Normal` で基準法線を決定
2. 各領域の `GetBoundaries()` を `GeometryCreationUtilities.CreateExtrusionGeometry(loops, normal, 1.0)` で薄板ソリッド化
3. `BooleanOperationsUtils.ExecuteBooleanOperation(..., BooleanOperationsType.Union)` で順に和集合
4. 結果ソリッドの **上面**（`PlanarFace.FaceNormal.IsAlmostEqualTo(normal)`）を取得
5. `topFace.GetEdgesAsCurveLoops()` で統合後の境界ループを取得
6. **`Transform.CreateTranslation(-thickness * normal)` で元平面に戻す**（押し出し分オフセットしているので戻さないと Z が 1ft ずれる）
7. 元の領域を全削除し、新しい領域を作成

**フォールバック**: Union に失敗した場合は従来の単純連結にフォールバック

#### 開発知見
- `GeometryCreationUtilities.CreateExtrusionGeometry` は CurveLoop の向き（CCW/CW）と normal の右手則が一致しないと失敗する → `GetPlane().Normal` から得た normal を使えば OK
- 上面ループの Z 座標は押し出し分オフセットするので、必ず元平面に戻す

---

## ExcelExportImport 開発知見

### Excel COM の `Interior.Color` 形式
- **`R + G*256 + B*65536` 形式**（VBA の `RGB()` 関数と同じ）
- **BGR ではない!** — 当初 `B + G*256 + R*65536` と誤解してRとBが逆になり、黄色のつもりが水色になった
- 例: `RGB(255, 255, 153)` = `255 + 255*256 + 153*65536` = `10092543`

### 数値パラメータの Excel 書き込み
- `ClosedXML` の `cell.Value = stringValue` はテキスト形式で保存される → Excelで「数値が文字列として保存されています」警告が出る
- **数値は `double` 型で書き込む**: `double.TryParse` で変換してから `cell.Value = numValue`

### テキスト/数値 混在時の値比較
- `GetString()` だけでは不十分。`cell.DataType == XLDataType.Number` をチェックし、整数なら小数点なしの文字列に変換
- 値比較は `ValuesAreEqual()` で数値比較にフォールバック（`"4700"` vs `4700.0` を同一と判定）

### Revit パラメータの読み取り専用制限
- 構造柱の「長さ」など、Revitが自動計算するパラメータは `param.IsReadOnly = true`
- API 経由で `Set()` しても例外が発生するため、インポート時にスキップが必要
- プレビューでは読み取り専用パラメータを非表示にし、サマリーで件数と理由を表示

### `AsValueString()` の戻り値
- `StorageType.Double` のパラメータは `AsValueString()` で表示単位での文字列を取得（例: 内部値 feet → 表示 "4700" mm）
- `SetValueString()` で表示単位の文字列からの設定が可能
- `AsValueString()` が `null` を返す場合があるため、`?? AsDouble().ToString()` でフォールバック

### ClosedXML でのファイル読み取り（Excel 開いている場合）
- `FileShare.ReadWrite` を指定しないと、Excelがファイルをロックしているため読み取りが失敗する
- `new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)` を使用

### ClosedXML リッチテキスト（セル内の部分書式設定）
- `cell.CreateRichText()` でリッチテキストオブジェクトを取得
- `richText.AddText("文字列")` で部分追加し、返り値に `.SetFontColor()`, `.SetBold()` で書式設定

### Excel COM の `Characters` による部分書式設定
- `cell.Characters[startPos, length]` で文字列の一部分を取得し、`Font.Color`, `Font.Bold` を設定可能
- **startPos は1ベース**（0ベースではない）

### シート統合モードでのインポート時の空セル問題
- シート統合モードでエクスポートすると、あるカテゴリに存在しないパラメータの列は空セルになる
- **対策**: `GeneratePreview` と `Import` の両方で `string.IsNullOrEmpty(newValue)` なら処理をスキップ

### ⚠️ CheckBox.Content に文字列をバインドすると `_` が消える
- WPF の `CheckBox.Content="{Binding Foo}"` では文字列中の `_` がアクセスキーとして解釈され非表示になる
- Revit のパラメータ名には `_` を含むものが多い（例: `T-H_s`, `Haunch_Calculation`）
- **対策**: `<CheckBox><TextBlock Text="{Binding Foo}"/></CheckBox>` の形に変更
- 同じ問題は `Label.Content`, `ContentControl.Content` 全般に起こる。テキスト表示はなるべく `TextBlock` で行うこと

### エクスポート範囲（スコープ）対応
- `Models/ExportScope.cs` enum: `EntireProject` / `ActiveView` / `Selection`
- `FilteredElementCollector` の生成方法を切り替える:
  - ActiveView: `new FilteredElementCollector(doc, view.Id)`
  - Selection: `new FilteredElementCollector(doc, selectionIds)`
  - EntireProject: `new FilteredElementCollector(doc)`

### 通り芯・レベル等の注釈カテゴリのエクスポート対応
- 従来 `CategoryType.Model | AnalyticalModel` のみ許可していたが、`Annotation` も対象化
- 注釈は数が多いので **ホワイトリスト** で絞る（`RevitCategoryHelper.IsUsefulAnnotationCategory`）
  - `OST_Grids`, `OST_Levels`, `OST_Sheets`, `OST_Views`, `OST_Viewports`, `OST_TextNotes`,
    `OST_GenericAnnotation`, `OST_RevisionClouds`, `OST_ScheduleGraphics`

### パラメータ全網羅取得
- 旧: `.Take(50)` で先頭 50 要素だけサンプリング → レアなタイプのパラメータが漏れる
- 新: 「各タイプにつき先頭インスタンス 1 件」+「カテゴリの全タイプ（`WhereElementIsElementType`）」で重複排除しつつ網羅

### 配布ZIP自動アップロード（設定済み情報）
- **PAT**: Classic token（`repo` スコープ）— Fine-grained token では権限不足でリリース作成が失敗する
- **Secret名**: `DOWNLOAD_SITE_TOKEN`（Revit-Add-ins リポジトリの Settings → Secrets → Actions に登録）
- **トークン指定方法**: `softprops/action-gh-release` は `with: token:` で指定（`env: GITHUB_TOKEN:` では動作しない）
- **配布サイト側**: `js/main.js` の `downloadConfig.urls` を GitHub API (`releases/latest`) で自動取得に改修済み
- リリース本文は `build-and-release.yml` の body セクション（2箇所: Revit-Add-ins用 と 28tools-download用）を更新

---

## FireProtection 開発知見

### ⚠️ アイコン PNG の DPI メタデータが必須（WPF 表示サイズ問題）
- **症状**: 自作した96px PNGをLargeImageに設定するとリボンからはみ出して表示される
- **原因**: WPF は PNG の DPI メタデータを基に論理ピクセル寸法を決定する
  - 既存の動作するアイコン（beam_top_level_96.png 等）は **DPI=288** が設定されている
  - 96px ÷ (288/96) = 32 論理ピクセル として表示される
  - DPI が未設定だと 96px がそのまま 96 論理ピクセルとして扱われ、ボタン枠を超える
- **対策**: Pillow で保存時に `dpi=(288.0106, 288.0106)` を指定する
  ```python
  img96.save(path, dpi=(288.0106, 288.0106))
  ```
- 32px版は `dpi=(96.012, 96.012)` で OK

#### アイコンの右側カラーパレット仕様（既存と統一）
- **8×8 ピクセルの塗潰しのみ、囲い線なし**
- x範囲: 22-29、y範囲: 1-8 / 11-18 / 21-28（2px間隔）
- 色: ピンク`(255,128,148)`, 黄`(218,185,47)`, 青`(30,144,255)`
- ピクセル単位で `beam_under_level_32.png` と一致させること

### ⚠️ トランザクションブロック内で宣言した変数のスコープ
- `using (Transaction trans) { ... }` ブロック内で宣言した変数は、ブロック終了後に使えない
- 凡例シート自動配置（別トランザクション）で `hasColumnFrame` 等のフラグを使う場合は、**トランザクション開始前に宣言**して、ブロック内で代入のみ行う

### Viewport.Create のサイズ取得問題
- 凡例ビューのサイズを `view.GetBoxOutline()` で取得しようとすると失敗するケースあり（ビューが直前のトランザクションでコミットされたばかりで内部キャッシュが未更新）
- **対策**: 項目数からサイズを推定して直接配置する方式が確実

### 自動配色アルゴリズムの注意点
- ベース色を3色だけにして残りを明度シフトで生成すると、cycle≥1 のときに元色とほぼ同色になる
- **対策**: ベースパレットを十分な数（12色程度）まで拡張し、超過時のみ明度を大きく変化させる

### 凡例の Excel セル表形式
- タイトル行 + 各種類の色行 + 柱行 + 注記行を、表形式（横線・縦線）で構成
- 横線: 各行の境界 + 上下端
- 縦線: 左端 / 色四角列右端（colDivX=20mm）/ 右端
- 色四角は 18×8mm（行内に1mm余白）
- 注記は `※` ごとに別 TextNote、行間隔 `textHeight * 1.6 * lineCount + textHeight * 2.5`

### 柱枠線の非表示
- 柱の枠型の塗潰しは外周線が出るため、`SetLineStyleId()` で「非表示」または「Invisible」スタイルを適用して枠線を消す
  ```csharp
  var invisStyle = ... GraphicsStyle ... where Name.Contains("非表示") ...;
  if (invisStyle != null) fr.SetLineStyleId(invisStyle.Id);
  ```

### 凡例シート自動配置の設計
- アクティブビューがシートで凡例ビューが作成済みの場合のみ動作
- **別トランザクション**で実行（凡例ビューをコミット後に Viewport.Create する必要があるため）
- 配置位置: 右上角固定
  - `estW = 85mm`（色四角列20mm + テキスト列65mm）
  - `estH = titleH + totalRows * rowH + noteH` （`noteH=45mm`は注記セクション概算）
  - 余白 `margin = 50mm` + 微調整 `upOffset = 25mm`, `rightOffset = 30mm`
- デバッグログ: `C:\temp\FireProtection_debug.txt` に VP ID と座標を記録

---

## FormworkCalculator 開発知見

### ワークセット可視性の落とし穴（2026-05-21）

#### `IsWorksetVisible` は新規WSで常に `false` を返す（Revit API バグ）
- `WorksetDefaultVisibilitySettings.IsWorksetVisible(wsId)` が新規作成直後のWSで `false` を返す
- にもかかわらず Revit UI では「全ビューに表示」チェックが入っている
- **対策**: ガード (`if (!IsWorksetVisible) SetWorksetVisibility(false)`) は不要。常に `SetWorksetVisibility(false)` を呼ぶ
- 既存WS・新規WS の両方で毎回 `SetWorksetVisibility(false)` を呼ぶことでUI表示が一致する

#### per-view ワークセット可視性の優先順位
- Global default (`WorksetDefaultVisibilitySettings`) よりも per-view 設定 (`view.SetWorksetVisibility()`) が優先
- グローバルを Hidden にしても、per-view で Visible に設定していれば表示される
- `EnsureFormworkWorksetsVisible` を Global set の**後に**呼ぶことで解析ビューのみ Visible を保証

### ClosedXML AssemblyResolve ハンドラの注意点（2026-05-21）

`System.*` アセンブリを一律スキップすると ClosedXML の NuGet 推移的依存が壊れる。
ホワイトリスト方式に変更が必要:

```csharp
bool isNuGetDependency =
    assemblyName.Equals("System.Runtime.CompilerServices.Unsafe", StringComparison.OrdinalIgnoreCase) ||
    assemblyName.Equals("System.Memory", StringComparison.OrdinalIgnoreCase) ||
    assemblyName.Equals("System.Buffers", StringComparison.OrdinalIgnoreCase) ||
    assemblyName.Equals("System.Numerics.Vectors", StringComparison.OrdinalIgnoreCase) ||
    assemblyName.Equals("System.IO.Packaging", StringComparison.OrdinalIgnoreCase) ||
    assemblyName.Equals("System.Threading.Tasks.Extensions", StringComparison.OrdinalIgnoreCase) ||
    assemblyName.Equals("System.ValueTuple", StringComparison.OrdinalIgnoreCase);
if (!isNuGetDependency && assemblyName.StartsWith("System.", ...))
    return null; // スキップ
```

### 切断ボックス（SectionBox）操作の注意点（2026-05-21）

#### `IsSectionBoxActive = true` を `SetSectionBox` なしで呼ぶと空のBBoxが有効化される
- 空のセクションボックスが有効化されると全 DirectShape がクリップアウトされて何も見えなくなる
- **正しい順序**: 必ず `view.SetSectionBox(validBB)` → `view.IsSectionBoxActive = true` の順で呼ぶ
- BBoxが計算できない場合は `IsSectionBoxActive = true` を呼ばない（何もしない）

#### ソースビューに切断ボックスがない場合の正しい挙動
- NG: 解析ビューにも切断ボックスを設定しない → 全体表示になり型枠DSがシート上で極小
- OK: `EnableSectionBox` で型枠対象要素の BoundingBox から切断ボックスを算出して設定
- `elem.get_BoundingBox(null)` は `doc.Regenerate()` 前でも元の構造要素なら取得可能

#### `EnableSectionBox` の安全な実装パターン
```csharp
// result.ElementResults の BBox を合算して切断ボックスを設定
// 要素が一つもない場合は IsSectionBoxActive を呼ばない（空BBox有効化を防止）
if (minP != null && maxP != null)
{
    view.SetSectionBox(new BoundingBoxXYZ { Min = minP - margin, Max = maxP + margin });
    view.IsSectionBoxActive = true;
}
// else: 何もしない（安全）
```

#### ⚠️ 未解決: 3Dビューの切断ボックス座標系（2026-05-21 時点）
- `EnableSectionBox` はワールド座標 (`elem.get_BoundingBox(null)`) で BBox を計算
- Revit の切断ボックスはビューローカル座標系（`BoundingBoxXYZ.Transform`）で定義される
- 回転・傾斜のある 3D ビューではワールド座標でセットした切断ボックスがズレる可能性
- 特に `**型枠：` 系の Legacy3 ビューで発生リスクあり（カメラ向き不明）
- **次セッションで調査が必要**: `view.GetSectionBox().Transform` を確認し座標変換が必要か検証

### Legacy3 ビュー（`**型枠：` プレフィックス）の扱い（2026-05-21）

旧アドインバージョンが作成した `**型枠：{sourceViewName}` 形式の 3D ビュー:
- `IsAnalysisViewName()` には含めない（シートに余計なビューが貼り付けられるため）
- `HideAllFormworkShapesInOtherViews` の hide ループには含める（DS を非表示にするため）
- `CollectAllAnalysisViewIds` では除外し、現バージョンの解析ビュー (`3D_型枠数量 -`) のみ返す

---

## FormworkCalculator 開発知見（旧）

### 処理パイプライン（3 Pass）
```
Pass 1: 要素毎に Solid 取得 → FaceClassifier で分類
Pass 2: ContactFaceDetector で接触面を DeductedContact に変更（Full + Partial 両対応）
Pass 3: 開口加算 + ElementResult 作成 + Aggregate
```

### 共有パラメータ（OST_GenericModel にバインド）
- `28Tools_FormworkMarker` (Text): 識別用
  - 通常 formwork: `"28Tools_Formwork"`
  - 除外 (鉄骨・デッキスラブ): `"28Tools_Formwork_Excluded"`
  - クリーンアップは `StartsWith("28Tools_Formwork")` で両方カバー
- `28Tools_Formwork_部位` (Text): 柱/梁/壁/スラブ/基礎/階段/鉄骨(除外)/デッキスラブ(除外)
- `28Tools_Formwork_レベル` (Text): 参照レベル名
- `28Tools_Formwork_区分` (Text): 色分けグループキー
- `28Tools_Formwork_面積` (Area): 要素単位の最終 FormworkArea（最初の FormworkRequired DirectShape に持たせる）
- `28Tools_Formwork_部分接触` (Text "Yes"/"No"): 一部消されている面の識別

### Revit 2022 Schedule API の制約まとめ

#### 1. `TableSectionData.SetCellStyle` は限定的
スタイル上書き許可セル:
- ✓ Header セクションの全セル
- ✓ Body セクションの **行 0（列ヘッダー）のみ**
- ✗ Body のデータ行・グループフッタ・**総合計行はスタイル変更不可**

エラー: `ArgumentException: Only allow to override cell style for header section or column header in body section.`

#### 2. `TableCellStyle.FontSize` は Revit 2022 に存在しない
- Revit 2024+ で追加。リフレクションで設定しても silently fail
- → **Revit 2022 ではプログラム経由でスケジュールフォントサイズ変更は不可能**

#### 3. `ScheduleDefinition.GrandTotalTitle` 設定の前提条件
- `ShowGrandTotalTitle = true` を**先に設定しないと setter が TargetInvocationException を投げる**

#### 4. ⚠️ `doc.Regenerate()` を呼ばないと新規 DirectShape のジオメトリは認識されない
- `DirectShape.SetShape` 直後の `get_BoundingBox(null)` は `null`、`get_Geometry(opts)` は Solid 数 0
- 結果として **3D ビューに描画されない**
- 対策: 全 DirectShape 作成後に `doc.Regenerate()` を必ず呼ぶ（同じトランザクション内 OK）

#### 5. `TableSectionData.SetColumnWidth` の単位
- Revit 内部単位 = feet
- `0.167 ft ≈ 50mm`、`0.5 ft ≈ 152mm`

#### 6. 動的合計表示パターン
集計表の Body 総合計行はスタイル変更不可なので、styled な合計を実現するには:
- 別途サマリ集計表を作成し `IsItemized = false`
- 件数 + 面積（`DisplayType=Totals`）の 2 フィールド
- Body 行 0（列ヘッダー）の各セルにスタイル設定

#### 7. ClosedXML CJK 文字幅
- `Column.AdjustToContents()` は半角換算でしか計算せず日本語が見切れる
- 自前で `MeasureWidth(string)` を実装（CJK 全角=2.0、半角=1.1）して `Column.Width` を直接設定

### Revit 2022 では `ScheduleField.HasTotals` が存在しない
- **正解 API**: `ScheduleField.DisplayType = ScheduleFieldDisplayType.Totals`
- `ScheduleFieldDisplayType` enum 値は Revit 2021-2026 全バージョンで利用可能

| 値 | UI ラベル |
|---|---|
| `Standard` (0) | 計算しない |
| `Totals` (1) | 合計を計算 |
| `MinMax` (2) | 最小値と最大値を計算 |
| `Max` (3) | 最大値を計算 |
| `Min` (4) | 最小値を計算 |

### View3D の視点コピー
```csharp
if (sourceView is View3D src)
    targetView.SetOrientation(src.GetOrientation());
```

### 解析3Dビューで非表示にすべきカテゴリ
- `OST_SectionBox`（切断ボックスのアウトライン）
- `OST_Levels`（レベル線）
- `view.SetCategoryHidden(catId, true)` で非表示化

### シート自動配置のレイアウト（2026-05-14 更新）

配置ロジック (`PlaceScheduleAt`):
- 配置インスタンスを作成した後 `inst.get_BoundingBox(sheet)` から実 BB を取得して右端 X / 下端 Y を返す
- 取得失敗時は概算値 (幅 ≈ 213mm、高さ ≈ 100mm) でフォールバック
- 集計表間のギャップは `gap = 0.05 ft (≈15mm)`、シート余白は `margin = 0.082 ft (≈25mm)`

折り返しロジック:
- 配置後の右端 (`placedRight`) がシート右マージン (`rightX`) を超え、かつ既に行頭ではない場合、配置済みのインスタンスを `doc.Delete()` で削除し次行に移動

### 集計表の列幅自動調整（2026-05-14 更新）

改善後の計算式:
```
widthMm = max(headerUnits, maxValueUnits) * 2.6 + 12.0
```

パラメータの根拠:
| 項目 | 旧 | 新 | 理由 |
|---|---|---|---|
| 文字幅係数 (mm/単位) | 2.0 | **2.6** | Revit 既定フォントの実描画幅に合わせる |
| パディング (両側合計 mm) | 7 | **12** | 罫線とテキストの間に余裕を持たせる |
| 最大幅キャップ (mm) | 200 | **250** | 長いタイプ名対応 |

全角文字の幅換算 (`MeasureTextUnits`):
- CJK (全角) = 半角の 2.0 倍幅
- 半角 = 1.0 単位

### マテリアルベース算出 — 中止決定（2026-05-07）
**中止理由**: マテリアル単独では SRC/CFT を区別できない（CFT を Concrete マテリアルで作るケースがあるため、マテリアル属性だけでは型枠要否を判断不可）。代わりに **鉄骨除外** (`SteelMemberDetector` の 4 層判定) と **デッキスラブ除外** (`DeckSlabDetector`) で実用上の課題は解消した。

---

## 多言語UI(LocSystem) 開発知見

### リボンボタンのランタイム更新
- `RibbonItem.ItemText` でボタンテキストを変更可能
- `PulldownButton.Image` でアイコンを動的に差し替え可能
- パネルタイトルは `RibbonPanel.Title` で変更可能
- ボタンとパネルの参照は `Application.cs` のフィールドに保持しておく必要がある

### ボタン名⇔ローカライゼーションキーのマッピング
- `_buttonTextKeys` / `_buttonTipKeys` で内部ボタン名とキーを対応付け
- `_panelKeys` はパネルのインデックスベースの配列
- ボタン追加時にマッピングも同時に追加しないと言語切替で更新されない

### 国旗アイコンの動的差し替え
- `LoadImage()` で `pack://application:,,,/` URI から読み込み → `BitmapImage.Freeze()` 必須
- 言語コード → ファイル名の変換: `$"flag_{Loc.CurrentLang.ToLower()}_16.png"`
- 16px版をスタックボタンの `Image` に設定、32px版はプルダウンサブボタンの `LargeImage`

### 設定パネルの3段スタック
- `AddStackedItems()` は2個または3個の `RibbonItemData` を受け取る
- 3段スタック時はアイコンサイズ16px、テキストは短めにする
- プルダウンボタンもスタックアイテムとして配置可能

---

## 多言語バグ過去事例 {#多言語バグ過去事例}

過去に発生したキー名の不一致例:

| ダイアログ内の誤ったキー | 正しいキー |
|---|---|
| `Export.SelectCategory.Header` | `Export.Category` |
| `Export.SplitByCategory` | `Export.SeparateSheets` |
| `Export.ResetSettings` | `Export.RestoreSettings` |
| `Import.OpenFiles` | `Import.OpenFile` |
| `Import.SelectExcelFile` | `Import.SelectFile` |
| `Import.ChangePreview` | `Import.Preview` |
| `Import.Column.ElementId` | `Import.ColElementId` |
| `Import.Column.Category` | `Import.ColCategory` |
| `Import.Column.Parameter` | `Import.ColParameter` |
| `Import.Column.CurrentValue` | `Import.ColCurrentValue` |
| `Import.Column.NewValue` | `Import.ColNewValue` |

---

## GitHub Pages 配色テーブル

| 色 | コード | 用途 |
|---|---|---|
| Blue-green | `#5F968E` | フェーズヘッダー、矢印 |
| Mint | `#BFDCCF` | あなた（You）ノード背景 |
| Oatmeal | `#D5C9B1` | AI・自動（Auto）ノード背景、結果ノード、ブラウザフレーム |

派生色（背景色・ボーダー等）はメインカラーから明度を調整して生成。

---

## squash merge によるオートマージ連続失敗の既知問題と対策

自動マージは **squash merge** で行われるため、main 上のコミットハッシュはブランチのものと異なる。

**根本原因**:
1. push → squash merge 成功 → main に「Auto: xxx」コミットが作成される
2. 次の変更を push → ブランチには旧コミット + 新コミットが乗っている
3. `git merge --squash` が旧コミットの変更と main の squash コミットでコンフリクト

**対策（hookで自動化済み）**:
- PreToolUse hook が push 前に `git rebase origin/main` を自動実行
- rebase により旧コミットは `skipped previously applied commit` としてスキップされる
- hook は dirty working tree も `git stash` で対応済み

---

## 型枠数量算出 v2.1.1 修正の知見（2026-05-27）

v2.1 リリース後に発見された 5 件の不具合と、それぞれの修正方針・教訓を記録する。

### 1. 接触面に型枠 DS が形成される不具合（BuriedFaceDetector の新設）

**症状**:
- 床と梁の段差フラッシュ接触部や、床と布基礎の体積重なり部分に、不要な型枠 DS が両面分作成される

**原因**:
- 既存の `ContactFaceDetector` は「anti-parallel + UV-on-face」パターンの直接接触しか検出しない
- 以下 2 つの盲点があった:
  1. **SpatialGrid ペアフィルタの取りこぼし**: 床の `ctx.BB.Y` が実体より小さく算出され (例: 5650×600×900mm と報告されるが face 詳細では 5650×2700×900mm)、隣接梁の BB と overlap せずペアテスト自体が走らない
  2. **「面が他要素のソリッド内部に埋もれている」ケース**: 対向面が存在しないため anti-parallel 判定が成立しない (例: 接合されずに体積が重なっている床×布基礎)

**修正**: `Engine/BuriedFaceDetector.cs` を新設し Pass 2 直後に実行
- 各 `FormworkRequired` 面の中心点を外向き法線方向に 5mm オフセットしたサンプル点 `p_out` を作る
- `p_out` が他要素のソリッド内部にあれば `DeductedContact` に降格
- 内包判定は「凸ソリッド前提の平面署名距離テスト」(高速) → 失敗時は非整列方向へのレイキャスト (非凸対応) の二段構え
- `ctx.BB` を信頼せず、ソリッドの全エッジ端点 + 全 face UV 中心点から再構築した堅牢 BB で候補絞り込み

**回帰の落とし穴 (柱+梁の取り合いで型枠が消える)**:
- 柱 950x950 と梁 600x1160 の標準取り合いで、柱の +X 側面 (大面) の中心が梁体積内に入る
- 単純に「中心が内部 → 全面 DeductedContact」と判定すると、梁からはみ出した柱の上下端 (合計 ~355mm) まで型枠不要扱いになる
- **対策**: 対象面が相手要素との `PartialContact` を既に保持している場合は、その相手要素についての埋没判定をスキップする (既存の部分接触クリッパーに処理を委ねる)

### 2. 型枠 DS の面積が負値になる不具合（按分スケーリング）

**症状**:
- 開口の多い壁 (工作物擁壁 t1000～1600 等) で、特定の DS の面積パラメータが -8.299m² 等の負値になる

**原因**:
- `FormworkVisualizer` 内で、開口部の控除量 `openingDelta = OpeningEdgeAreaAdded - OpeningAreaDeducted` を**最初の `FormworkRequired` DS に一発で全部乗せていた** (`areaM2 += openingDelta`)
- 開口控除が当該 DS の素の面積より大きいケースで結果が負値に
- 例: face[0] 29.90m² + openingDelta -38.20m² = -8.30m²

**修正**: 「最初の DS に全部乗せる」方式を廃止し、**スケーリング係数で全 `FormworkRequired` DS に按分**する方式に変更
- `areaScale = er.FormworkArea / sum(全 FormworkRequired 面の EffectiveAreaM2)`
- 各 DS area = `fi.EffectiveAreaM2 × areaScale`
- 個々の DS は決して負にならず、合計はぴったり `er.FormworkArea` に一致

**教訓**: 集合の合計を末端の 1 要素で帳尻合わせする実装は、調整量が大きいと末端が破綻する。集合全体に按分する方が頑健。

### 3. 分析ビューのフィルタが正しく引き継がれない問題

**症状の変遷 (3 段階のイテレーション)**:
1. 初期: 派生フィルタ `28T_FW_{fid}_GM` / `28T_FW_{fid}_Other` が作成され、名前が不透明
2. 第1版 (a18ebd5): 派生フィルタ名を `{元名}_型枠除外GM` / `{元名}_型枠除外` に。一部のフィルタだけ別名になる不整合
3. 最終版 (9525afd〜6574432〜2d64b54): 派生フィルタを廃止。ソースフィルタを直接参照させる

**最終解 — カテゴリ変更による根本回避**:
- 旧: 型枠 DS は OST_GenericModel カテゴリ。ユーザー独自フィルタが「一般モデル」を含むと型枠 DS にもヒットして visible=false 干渉が起きる
- 新: 型枠 DS のカテゴリを `OST_NurseCallDevices` (ナースコール装置) に変更。一般建築モデルでは 100% 使われないため、ユーザーフィルタが衝突する確率がほぼ 0

**注意点**:
- DirectShape のカテゴリ変更は影響範囲が広い (約 25 箇所): DirectShape 作成 / DirectShapeType / フィルタ / 集計表 / カテゴリ可視性 / V/G オーバーライド / CleanupExistingFormworkShapes / FilterMatchesFormwork etc.
- 旧 GenericModel DS との互換のため `FormworkParameterManager.LegacyFormworkCategory` 定数を導入し、共有パラメータは新旧両カテゴリにバインド
- `CleanupExistingFormworkShapes` も新旧両カテゴリを走査して旧 DS をマイグレーション
- `28T_型枠_全非表示` フィルタや `型枠_柱` 等の色フィルタも、旧カテゴリのみ対象の場合は作り直す
- **ディシプリン**: `OST_NurseCallDevices` は Electrical discipline 所属のため、分析ビューの `BuiltInParameter.VIEW_DISCIPLINE` を Coordination (=4095, 全ビット ON) に明示設定して構造系ビューでも表示できるようにする

### 4. 更新モードで既存ビューのフィルタ設定が変更される

**修正**:
- 更新モード (`reusedView == true`) では `FormworkFilterManager.ApplyColorFilters` を呼ばないようガード
- ユーザーが手動で調整した色・可視性設定を保持
- 新規キーが発生した場合の追加処理は割り切ってスキップ (ユーザーは更新モード = 既存維持の意図)

### 5. シートに過去ビューもレイアウトされる

**修正**:
- シート作成時のビュー収集を `CollectAllAnalysisViewIds(doc)` (プロジェクト全体) → `perViewAnalysisViewIds` (今回実行分のみ) に変更
- 「複数の3Dビューを選択して実行」した時、過去の他ソースビュー由来の分析ビューはシートに載せない (ユーザーの自然な期待動作に合致)

### 6. その他の知見

- **Revit フィルタの可視性は AND 結合**: ある要素に複数フィルタがマッチする時、いずれかが `visible=false` なら要素は隠れる。順序やプライオリティは無関係 (グラフィックオーバーライドの優先順位とは別概念)
- **共有パラメータのカテゴリ追加**: `BindingMap.Insert` は既存バインドに対しては no-op。新カテゴリを追加するには `BindingMap.ReInsert` を使う必要がある
- **ベリファイ用ログ**: `[Buried]`, `area reconcile`, `[Filter]` 等のタグを debug log に残しておくことで、ユーザー報告時の問題箇所が即座に特定できた

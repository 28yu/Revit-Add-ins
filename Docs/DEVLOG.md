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

### コミット件名マーカーによる対象バージョン切替（`[build:XXXX]`）
- **目的**: デフォルトは `dev-config.json`（通常 2022）のまま、特定のコミットだけ別バージョンでオートビルドしたい
- **仕組み**: `Get-BuildVersions` が HEAD コミットメッセージ（`git log HEAD -1 --format="%B"`）を正規表現
  `\[build:\s*([0-9,\s]+|all)\s*\]` で解析。マーカーがあればそのバージョンのみ、なければ
  `Get-DefaultRevitVersion`（dev-config.json）にフォールバック。`Run-Build` は解決した各バージョンを
  `QuickBuild.ps1 -RevitVersion` で**明示ビルド**（dev-config には依存しない）
- **⚠️ マーカーは必ずコミット件名（1行目）に置く**: `auto-merge-claude.yml` の squash は
  `head -n 1` で件名のみを `main` に残すため、本文マーカーは `origin/main` に伝わらず機能しない
- `[build:all]` で全6バージョンを順次ビルド。複数指定 `[build:2024,2025]` も可
- 状態を持たない設計（コミット単位）なので、次のコミットで自動的にデフォルトへ戻る

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

### ヘッダー行の固定（ウィンドウ枠固定）
- 書き出した Excel はスクロールしても1行目（ヘッダー）が常に見えるように固定する
- ClosedXML: ヘッダー書き込み後に `worksheet.SheetView.FreezeRows(1)` を呼ぶ
- 分割シート／単一シートの両方の書き出しパスに追加

### ⚠️ 開いているExcelを読み込む/色付けするとフリーズする問題（COMのセル単位アクセス）
- **症状**: インポート対象のExcelが**開いている**とき、読み込み〜色付けでRevitがフリーズする
- **原因**: `ExcelProcessHelper.MarkCellsViaCom`（開いているExcelにCOMで色付け）が
  `sheet.Cells[row, col]` を**1セルずつ**アクセスしていた。COMはプロセス間呼び出しのため、
  大きな表では全セル分の往復（O(行×列)）が発生し数十秒〜フリーズになる
  （※Excelが閉じている場合は ClosedXML でディスク上のファイルを編集するのでこの経路には入らない）
- **修正**:
  1. ヘッダー行・ID列を **`Range.Value2` で一括読み取り**（1シートあたり数回のCOM呼び出しに削減）
  2. メモリ上で対象行を判定し、**色付けが必要な行だけ**COM操作。背景色は行範囲を1回で着色
  3. `Application.ScreenUpdating=false` に加え `EnableEvents=false` も設定（finallyで復元）
- **読み込み側**: `GeneratePreview` にシート名一覧の out 版を追加し、`GetSheetNames` の
  別途オープン（ClosedXML二重パース）を解消。プレビュー生成でファイルを1回だけ開く

### ⚠️ 設定ファイル(.json)読込が遅い問題（パラメータ二重取得の解消）
- **原因**: `ExportDialog.ApplySettings` が `UpdateParameterList()`（選択カテゴリの
  Revitパラメータ取得）を**2回**呼んでいた。JSONパース自体は軽く、遅さの主因はこの二重取得
- **修正**:
  1. 2回目は Revit 再取得ではなく、取得済み `_allParameters` から表示だけ更新する `FilterParameterList(null)` に変更
  2. カテゴリ照合を `HashSet`、出力パラメータ照合を `Dictionary` にして線形探索を排除

### パラメータ欄のカテゴリ見出しを折りたたみ可能に（▼/▶）
- **目的**: 複数カテゴリ選択時にパラメータが大量に並びスクロールが煩雑 → カテゴリ単位で開閉したい
- **実装**: `ParameterListBox` の `GroupStyle` を `HeaderTemplate` から `ContainerStyle`（`GroupItem` を
  Expander で再テンプレート）に変更。カテゴリ名の左に `▼`（展開）/`▶`（折りたたみ）を表示し、
  見出しクリックで開閉。既定は全て展開（`IsExpanded=True`）
- Expander は独自 `ControlTemplate` でテーマ非依存化。`ToggleButton.IsChecked` を `Expander.IsExpanded` に
  TwoWay バインド（`ToggleButton.IsChecked` は既定で双方向）
- グループ化時のUI仮想化（`IsVirtualizingWhenGrouping`）はそのまま。折りたたみ中の行は描画されないため
  スクロール対象も減る

### 削除（クリア）成功セルの色付け（青字ではなく青塗り）
- 変更・追加の成功セルは**青字**（文字色）で示すが、削除（空欄化）の成功セルは**文字が無い**ため青字にならない
- そこで削除成功セルは**セルを青で塗りつぶす**（背景色）ことで示す
- `MarkImportedCells` が `changedSet`（値あり成功=青字）と `clearedSet`（削除成功=青塗り）を分けて構築し、
  `MarkCellsViaCom`（COM）/`MarkCellsViaClosedXml`（ClosedXML）の両方へ渡す
- 凡例を「青字・青セルはインポート成功（青セルは値の削除）、赤字はインポート失敗」に更新
  （COM側は `Characters[start,len]` で位置指定のため、文言変更時は位置の再計算が必要）

### ⚠️ 値を削除（空欄化）したセルがインポートで反映されない不具合
- **症状**: 書き出したExcelの値を**削除**してインポートしても、その削除（クリア）が反映されない。
  追加・変更は反映されるが、空欄にした情報だけ無視される
- **原因**: `GeneratePreview` で空セルを**一律スキップ**していた（`if (IsNullOrEmpty(newValue)) continue;`）。
  これは元々シート統合モードの「他カテゴリのパラメータ列（N/A）」を無視するための処理だったが、
  「ユーザーが値を消したセル」も同じく無視してしまっていた
- **修正**: 空セルを「N/A」か「値の削除」かで判別する
  - `FindParameter` が **null**（要素にそのパラメータが無い）→ N/A としてスキップ（従来通り）
  - パラメータが存在し **文字列型（StorageType.String）** で現在値が空でない → 「クリア」変更として
    プレビュー行を生成（`NewValue=""`, `HasChange=true`）。`ImportFromPreview` が `SetParameterValue(param, "")`
    で空文字をセットしてクリアする
  - 数値・ElementId 型は Revit 上で空にできないためスキップ（クリア対象外）
- **補足**: この判定は空セルごとに `FindParameter` を呼ぶ。既定の分割シートモードでは空セルは少なく影響軽微。
  シート統合モードで空セルが多い場合は取得コストが増えるが、`LookupParameter` 化済みで許容範囲

### ⚠️ 大容量モデルでインポートが遅い問題（二重全走査の解消）
- **症状**: エクスポート同様、要素数の多いモデルでエクセルインポートが重い
- **原因**:
  1. `ParameterService.FindParameter` が要素の全パラメータを線形走査（エクスポートと同じ）
  2. **同じ全セル走査を2回**実行: `GeneratePreview`（ダイアログ）と `Import`（トランザクション）で
     Excel を再度開き、全セルに対し `FindParameter` + 値取得を繰り返していた
- **修正**:
  1. `FindParameterByName` を `Element.LookupParameter()`（ネイティブ）に変更 → 線形走査を廃止（プレビュー・インポート両方が高速化）
  2. **`ImportFromPreview(doc, previewRows)` を新設**。プレビューで計算済みの「変更あり かつ 書込み可能」
     セルのみを書き込む。トランザクション側での Excel 再読込・全セル走査が不要になり、
     実際に変更したセル数だけ処理する（例: 38万セル中100セル変更なら100セルのみ処理）
  3. `GeneratePreview` はタイプパラメータの現在値・読取専用フラグを `(タイプID|名前)` でキャッシュ
     （読み取りのみで安全。同一タイプの全インスタンスで共通の値を再計算しない）
- **カウント表示の変化**: `ImportFromPreview` は変更セル基準のため、旧 `Import` が空セル等も
  スキップ数に含めて水増ししていた「スキップ」件数がプレビュー表示と整合する値になる（成功≒変更あり件数）
- 旧 `Import(doc, filePath)`（ファイル直読みの全走査版）はフォールバックとして残置

### 配布ZIP自動アップロード（設定済み情報）
- **PAT**: Classic token（`repo` スコープ）— Fine-grained token では権限不足でリリース作成が失敗する
- **Secret名**: `DOWNLOAD_SITE_TOKEN`（Revit-Add-ins リポジトリの Settings → Secrets → Actions に登録）
- **トークン指定方法**: `softprops/action-gh-release` は `with: token:` で指定（`env: GITHUB_TOKEN:` では動作しない）
- **配布サイト側**: `js/main.js` の `downloadConfig.urls` を GitHub API (`releases/latest`) で自動取得に改修済み
- リリース本文は `build-and-release.yml` の body セクション（2箇所: Revit-Add-ins用 と 28tools-download用）を更新

### パラメータ名の `I-` / `T-` プレフィックスの意味
- `ParameterInfo` コンストラクタで付与（`Models/ParameterInfo.cs`）
  - `I-` = インスタンス（要素）パラメータ（`isTypeParameter == false`）
  - `T-` = タイプパラメータ（`isTypeParameter == true`）
- エクスポートダイアログのパラメータ欄下部に凡例を常時表示（`Export.ParamPrefixLegend`）
  - 背景色なし・`I-` と `T-` を改行して2行表示（ローカライズ文字列に `\n` を埋め込む）

### ⚠️ 大容量モデルでエクスポートに時間がかかりRevitがフリーズする問題（`ExcelExportService`）
- **症状**: 要素数の多いモデルをエクセルエクスポートすると書き出しに長時間かかり、その間Revitが応答なしになる
- **原因（3点、いずれも要素数×パラメータ数でスケール）**:
  1. **タイプパラメータの再計算**: セルごとに `ParameterService.FindParameter` を呼び、タイプパラメータでも
     毎回 `GetTypeId()` + `doc.GetElement()` + 全パラメータの線形走査をしていた。
     タイプパラメータは同一タイプの全インスタンスで値が同じなのに、38749要素分すべて再計算していた
  2. **線形走査**: `FindParameterByName` が要素の全パラメータを名前一致まで毎回走査
  3. **列幅の2度読み**: `AdjustColumnWidths` が書き込み後に全セルを `GetString()` で再走査（O(行×列)の二重コスト）
- **修正**:
  1. `ResolveParameterValue` を新設。タイプパラメータは `(タイプID, パラメータ名)` でキャッシュ（`typeValueCache`）し、
     ユニークなタイプ数×パラメータ数まで計算量を削減
  2. インスタンスパラメータは `Element.LookupParameter()`（ネイティブ検索）で取得し線形走査を廃止
  3. 列幅は書き込みループ内で `UpdateColWidth` により同時集計し、`ApplyColumnWidths` で適用（2度読みを廃止）
  4. オートフィルタは `RangeUsed()`（全走査）ではなく既知の範囲 `Range(1,1,row-1,totalCols)` を直接指定
- **注意**: Revit API はシングルスレッドのためデータ取得を別スレッドに逃がせない。
  上記は計算量削減による短縮であり、フリーズ時間そのものを短くするアプローチ

### ⚠️ 大容量RVTでパラメータのチェック/ホバーが重い問題（グループ化と仮想化）
- **症状**: 要素数の多いカテゴリ（例: 構造フレーム 38749 要素）を選ぶと、パラメータ欄の
  チェックボックスにカーソルを合わせただけでカーソルがクルクル回り、チェック操作も重い
- **原因**: `ParameterListBox` / `OutputListBox` は `ListCollectionView` + `PropertyGroupDescription`
  でカテゴリ別グループ表示している。**WPF はグループ化すると UI 仮想化が既定で無効になる**ため、
  全パラメータのチェックボックスが一度に実体化され、ヒットテスト・レイアウトが重くなる
- **修正**: 両 ListBox に以下を付与し、グループ化時も仮想化を有効化
  - `VirtualizingPanel.IsVirtualizingWhenGrouping="True"`（.NET 4.5+ で追加、これが肝）
  - `VirtualizingPanel.VirtualizationMode="Recycling"` / `IsVirtualizing="True"` / `ScrollUnit="Pixel"`
  - `ScrollViewer.CanContentScroll="True"`
- **注意**: `ParameterInfo` はチェック状態（`IsChecked`）をモデル側に保持するため、
  コンテナのリサイクルが起きても選択状態は失われない

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

---

## 型枠数量算出 v2.1.2 修正の知見（2026-05-28）

v2.1.1 リリース後に発見された 5 件の不具合と対応を記録する。

### 1. DirectShape 面積が負値になる不具合（按分スケーリング）

**症状**: 工作物擁壁 t1000～1600 等の開口の多い壁で、特定の DS の面積パラメータが
-8.299m² 等の負値になる。

**原因**: `FormworkVisualizer` が `openingDelta = OpeningEdgeAreaAdded - OpeningAreaDeducted`
を最初の `FormworkRequired` DS に一発で全部乗せていた (`areaM2 += openingDelta`)。
開口控除が当該 DS の素の面積より大きいケースで結果が負値に。
例: face[0] 29.90m² + (-38.20m²) = -8.30m²。

**修正**: 「最初の DS に全部乗せる」方式を廃止し、スケーリング係数で全 DS に按分。
```
areaScale = er.FormworkArea / sum(全 FormworkRequired 面の EffectiveAreaM2)
各 DS area = fi.EffectiveAreaM2 × areaScale
```
個々の DS は決して負にならず、合計はぴったり `er.FormworkArea` に一致。

**教訓**: 集合の合計を末端の 1 要素で帳尻合わせする実装は、調整量が大きいと末端が
破綻する。集合全体に按分する方が頑健。

### 2. マルチピース面の 2 個目以降の DS が 0m² になる

**症状**: `PartialContactClipper` が 1 面を複数 Solid に分割して作成した DS のうち、
2 個目以降の面積が 0m² になる。

**原因**: 旧設計で `bool firstPieceForThisFace = true` フラグを使って「最初のピース
だけに面の有効面積を全部乗せる」実装になっていた。1 番目の按分スケーリング修正で
合計が厳密に一致するようになった結果、0m² ピースが目立つようになった。

**修正**: 各ピースに `piece.Volume / sum_of_piece_volumes` で按分。
```
pieceShare = piece.Volume / sum_of_piece_volumes
areaM2 = fi.EffectiveAreaM2 × areaScale × pieceShare
```
ピースの体積は「ピース表面積 × 厚さ」なので、体積比 = 面積比。各ピースが幾何学的
サイズに比例した面積を持つ。

### 3. 一般モデル接触面の自動控除（BuriedFaceDetector 拡張）

**要望**: RC躯体が一般モデル (設備・什器・埋込物等) と接している面の型枠を省きたい。

**実装**: `BuriedFaceDetector.Run` に `Document doc` 引数を追加し、`OST_GenericModel`
カテゴリの障害物を別途収集して判定に組み込む。`FormworkMarker` パラメータで自前の
formwork DS (新 `OST_NurseCallDevices` および旧 `OST_GenericModel`) は除外。

### 4. 微小面 (0.01m² 以下) の自動除外

**要望**: 0.01m² 以下の面は集計から省きたい。

**実装**: `ComputeAndSetEffectiveArea` に絶対閾値を追加。
```
double minEffectiveFeetSq = UnitUtils.ConvertToInternalUnits(0.01, UnitTypeId.SquareMeters);
if (effectiveFeetSq > 0 && effectiveFeetSq < minEffectiveFeetSq) {
    effectiveFeetSq = 0;
    fi.FaceType = FaceType.DeductedContact;  // demoted-tiny
}
```
ログタグ `demoted-tiny(<0.01m²)` で識別可能。

### 5. エンベロープ GenericModel の誤検出と二段階フィルタリング

**症状**: 一般モデル障害物を有効にしたところ、1 つの巨大 GM (建物全体を覆うサイト
要素) が 2214 件の RC 面を誤って降格させていた。

**第1版の対策 (失敗)**: 体積 > 50m³ で除外。しかし「薄いシート状エンベロープ」
(例: 床全体を覆う厚さ 1mm の GM = 体積 25m³) は捕捉できなかった。

**第2版の対策 (採用)**: 体積基準 + BB 対角線基準の OR 条件。
- 体積 > 50m³ または
- BB 対角線 > 30m
のいずれかを満たす GM は「建物スケール」として除外。これで薄い大型シートも捕捉。

**教訓**: 物理的サイズ閾値は単一軸 (体積だけ等) では不十分。形状の多様性を考慮して
複数軸 (体積・対角線・最大辺等) の OR で判定するべき。

### 6. ログから問題を特定する勘所

- **頻度の集計** (`grep ... | awk '{print $X}' | sort | uniq -c | sort -rn`) で
  「1 つの要素 (E5325358) が 2214 件の降格を生んでいた」のような偏りを即座に発見
- `[Buried] skip GenericModel E... vol=... diag=...` のように除外時もログを残すと、
  「フィルタがどう判断したか」が後から検証可能 (今回の第1版→第2版改善の根拠になった)

---

## 運用知見: パッチリリースのフロー（2026-05-28 追記）

v2.1 (2026-05-25) → v2.1.1 (2026-05-27) → v2.1.2 (2026-05-28) と 3 日間で連続 2 つの
パッチをリリースしたフローから得た運用知見。

### 診断ログのやり取り (.diag/ フォルダ運用)

実機モデルが大きく Revit 環境でしか再現しないバグは、ユーザーから debug log を
受け取って原因特定する必要がある。チャットの添付サイズ制限 (~10MB) を超える場合や
複数ログを共有したい場合は、**リポジトリの `.diag/` フォルダに push してもらう**
方式が便利。

- ユーザー側: `Revit-Add-ins/.diag/Formwork_debug_YYYY-MM-DD_HH-MM-SS.txt` に push
- 開発側: `git fetch origin main && git show origin/main:.diag/<filename> > /tmp/log.txt`
- `.diag/` は GitHub 上でも 100MB まで普通にコミット可能
- 過去ログがそのまま履歴として残るので、修正の前後比較に便利

### 頻度集計で「主犯」を特定する

ログから「どの要素が問題を引き起こしているか」を素早く特定する基本テクニック:
```bash
grep "GenericModel E" log.txt | awk -F'GenericModel E' '{print $2}' | awk -F')' '{print $1}' \
  | sort | uniq -c | sort -rn | head
#   2214 5325358   ← 1要素で 95% の降格を生んでいた → 即座にエンベロープと判明
#     28 4639691
#     14 6175348
```

「平均」や「合計」より「分布の偏り」を見ることで真因が即時に見える。v2.1.2 の
エンベロープ問題はこの分析で 1 分で判明した。

### ヒューリスティック調整は試行錯誤

「サイト/エンベロープ要素を障害物から除外する」フィルタの調整:
- **第1版**: 体積 > 50m³ のみ → 薄いシート状エンベロープ (体積 25m³ 程度) が漏れた
- **第2版**: 体積 > 50m³ OR BB対角線 > 30m → ✅ カバー

教訓: 物理オブジェクトの形状は多様 (細長い・薄い・塊状)。サイズフィルタは
複数の幾何指標 (体積・対角線・最大辺等) を OR で組み合わせるべき。
**「除外しなかったケース」をログに残さない**と第1版で見落としに気付けなかった
ので、ログ駆動の試行錯誤を意識的に組む。

### パッチリリース時のファイル更新チェックリスト

CLAUDE.md の「リリース時の配布パッケージ更新チェックリスト」と重複するが、
v2.1.x のパッチではバージョン番号 + 修正点追記がメインなので簡略化版:

1. `Properties/AssemblyInfo.cs`: `AssemblyVersion` / `AssemblyFileVersion` 更新
2. `Packages/{2021..2026}/README.txt`: タイトル `vX.Y.Z` 更新 + 「【vX.Y.Z 修正点】」
   セクションを `【一つ前のバージョン 修正点】` の前に挿入 (履歴を残す)
3. `Packages/{2021..2026}/install.bat`: ヘッダー `vX.Y.Z` 更新
4. `Docs/Features/<機能名>.md`: 冒頭の修正サマリブロックに新エントリ追加
5. `Docs/DEVLOG.md`: セッションの知見を体系的に記録
6. claude/* ブランチに push (auto-merge で main 反映)
7. `release/vX.Y.Z` ブランチに force-push してリリースワークフロー起動

パッケージファイルの一括更新は Python の文字列置換で `【vX.Y 修正点】` の前に
新セクションを挿入する方式が安全 (sed では多言語混在で扱いにくい)。

### AutoBuild デプロイのタイムラグに注意

ユーザーの「修正されていない」報告を受けたとき、まず疑うべきは「AutoBuild が
最新の DLL をまだピックアップしていない」可能性。判別方法:

- ログの先頭にある実行開始時刻 (`Run() 開始` の行) と
- 最新コミット時刻
を比較する。デプロイ時刻のほうが遅ければユーザーは旧版で動作確認している。

その場合は「再度実行してログを送ってください」ではなく、**「数分待ってから
再度実行してください、AutoBuild が `git pull` した DLL を Revit が次回起動時に
読み込みます」**と案内する。Revit 起動中は DLL がロックされるため Revit 再起動も
必要。

---

## FillPatternIO（塗潰しパターン入出力）開発知見

### 機能概要
- 開いている rvt 内の塗り潰しパターンを一覧表示し、`.pat` ファイルへの
  書き出し／読み込みを行う（パネル: 注釈・詳細 / ボタン: 塗潰しパターン）。
- コード構成: `Commands/FillPatternIO/`
  - `FillPatternIOCommand.cs` — コマンド（ダイアログ起動）
  - `FillPatternIODialog.xaml(.cs)` — 一覧・検索・入出力ダイアログ
  - `FillPatternItem.cs` — 一覧行のビューモデル（チェック状態・プレビュー遅延生成）
  - `PatternData.cs` — 描画用に FillGrid から抽出したプレーンデータ（Revit非依存）
  - `PatternPreview.cs` — PatternData → プレビュー画像の描画（Revit非依存＝別スレッド可）
  - `PreviewRenderQueue.cs` — 専用STAスレッドでプレビューを非同期描画
  - `PatFile.cs` — `.pat` の書き出し／解析／モデルへの作成

### ⚠️ 最重要バグ: FillGrid セグメントは「位置で交互」（符号ではない）
- Revit の `FillGrid.GetSegments()` は **偶数index=ダッシュ、奇数index=ギャップ**
  の**全て正の値**を返す（`.pat` 標準の符号付き=負がギャップ、とは異なる）。
- 当初「正値=線」と解釈したため、ギャップまで線として描画され、複雑パターン
  （Concrete 等）が実線化・黒潰れした。
- 対応: 描画・書き出し・読み込みをすべて「位置で交互」規約に統一
  - 描画: `i % 2 == 0` をダッシュ、それ以外をギャップ。長さ0のダッシュ=点(ドット)。
  - 書き出し: `.pat` 標準準拠のため奇数indexを負値で出力。
  - 読み込み: `.pat` の符号付き値を絶対値化して SetSegments（往復整合を維持）。

### ⚠️ .pat の文字コードは ANSI（Shift-JIS）
- Revit/AutoCAD の `.pat` は ANSI 前提。UTF-8 で書くと日本語パターン名が Revit で文字化け。
- `CultureInfo.CurrentCulture.TextInfo.ANSICodePage` に合わせて読み書き。
- .NET 8（Revit 2025/2026）は既定でコードページ無効 → `System.Text.Encoding.CodePages`
  パッケージ（net8のみ参照）＋ `Encoding.RegisterProvider` が必要。
  `Application.cs` の AssemblyResolve の isNuGetDependency にも追加済み。
- 読み込みは UTF-8 BOM 付きも判定して対応。

### プレビュー描画の設計
- **固定実寸ウィンドウ方式**（Revit ダイアログと同様）。間隔の細/粗を忠実に表示。
  - 製図パターン: 紙面基準の固定 12mm ウィンドウ（`PreviewHeightMm`）。
  - モデルパターン: 実寸（数十〜数百mm）で固定窓に収まらないため、
    最も細かいグリッドが縦に約6本収まる**適応スケール**にズームアウト（固定窓より拡大はしない）。
    判定は `PatternData.IsModel`（`fp.Target == Model`）。
- 点(ドット)は長さ0ダッシュ→円で描画。ペンはラウンドキャップ（細い線=0.5px）。
- サブピクセル(0.5px未満)のダッシュ周期は実線化、線間引きは真のサブピクセル時のみ（忠実性維持）。
- 各線はウィンドウ4隅への射影で描画区間を厳密クリップ（点系パターンの負荷削減）。

### ⚠️ 大量パターン時のフリーズ対策（性能）
- 原因: 起動時に全パターンのサムネイルをUIスレッドで一括ラスタライズしていた。
- 対応:
  1. **プレビュー遅延生成** — `FillPatternItem.Preview` は初回アクセス時のみ描画依頼。
  2. **ListView 仮想化** — 画面内の行だけ実体化（列幅は固定にすること。Auto幅は全行測定で仮想化が無効化）。
  3. **バックグラウンド描画** — `PreviewRenderQueue` が専用STAスレッド(Dispatcher)で描画し、
     凍結済み ImageSource を UI へ marshal。UIスレッドを一切ブロックしない。
- **スレッド安全**: Revit API はスレッド非安全。描画に必要な数値は UI スレッドで
  `PatternData` に抽出し、描画スレッドは Revit API に触れない構成にすること。

### UI（検索・ホバー拡大）
- 検索欄: パターン名の部分一致（大文字小文字無視）。タイプ絞込と併用。
- サムネイルにホバーで拡大プレビュー(440x140)をツールチップ表示。拡大画像も遅延＋非同期生成。
  ツールチップ内の `{Binding LargePreview}` は PlacementTarget（=行の DataContext）を継承して解決。

### ソリッド塗り潰し
- グリッドを持たず `.pat` で表現不可。書き出しからは除外し件数を通知。プレビューは黒帯表示。

### WPF と Revit の型衝突（コンパイルエラー）
- `System.Windows.Point` と `Autodesk.Revit.DB.Point` が衝突（CS0104）。
  描画クラスは Revit非依存にして解消（`Autodesk.Revit.DB` を using しない）。

## ParameterCleanup（パラメータ整理）開発知見

新機能。プロジェクト内の削除可能なパラメータを一覧化し、同名の特定・値の有無の自動判定・不要パラメータの削除を行う。`Commands/ParameterCleanup/`（Command / Services/ParameterScanner / Models/ParamRow / Views/ParameterCleanupDialog）。

### 対象パラメータ（ユーザーが削除できるすべて）
- プロジェクト/共有: `FilteredElementCollector.OfClass(typeof(ParameterElement))` と `OfClass(typeof(SharedParameterElement))` を Id でマージ（`OfClass` が派生を含むか環境差があるため両方走査）。`GlobalParameter` は `is GlobalParameter` で除外して別処理。
- グローバル: `GlobalParametersManager.GetAllGlobalParameters(doc)`。値は `GlobalParameter.GetValue()` を `StringParameterValue`/`DoubleParameterValue`/... で分岐。
- 組み込み(Built-in)パラメータは `ParameterElement` を持たないため一覧に出ない＝削除不可。ファミリ内部パラメータは別ドキュメントのため対象外。

### バインド解決は `get_Item` ではなく `ForwardIterator`
- `doc.ParameterBindings.get_Item(definition)` は多くのパラメータで解決に失敗し、本来スキャン可能なものまで「バインドなし(NotApplicable)」になった。
- 解決策: `ParameterBindings.ForwardIterator()` で一度全走査し、`it.Key.Name`→(インスタンス/タイプ, カテゴリ) の辞書を作り、パラメータ名でひく。プロジェクトにバインド済みの名前は実質一意なので名前キーで安全。

### 大容量モデルのフリーズ回避（重要）
- Revit API はシングルスレッド。値判定（要素走査）は裏スレッド不可。→「走査量を減らす」＋「UIへ制御を返す」で対処。
- 列挙(EnumerateParameters)は `ParameterBindings`＋パラメータ要素の走査のみで軽量。
- 値判定(ScanRow)は**反復子(IEnumerable<int>)** にして一定件数(2000)ごとに `yield`。バインド先カテゴリの要素だけを `OfCategoryId` で収集し `(catId, inst/type)` でキャッシュ、値が1件でも見つかれば **early-exit**。
- UI側は `Stopwatch` で約50msごとにだけ `await Task.Delay(1)`（`Task.Delay` は Background 優先度で復帰するため Input=中止クリックが確実に処理される）。毎 yield 毎に await すると待機オーバーヘッドが累積するので時間ゲートで間引く。
- 値の有無はダイアログ表示時に `Loaded` で自動実行（削除後も再実行）。

### 値判定のセマンティクス
- 文字列: 空白のみは「空」。ElementId: 無効IDは「空」。数値/整数(Yes/No)は常に値ありとみなす（誤って数値パラメータを削除候補にしない安全側）。
- 表示の意味: 「空（未使用）」＝カテゴリにバインド済みだが全要素で値なし。「バインドなし」＝どのカテゴリにもバインドされておらず判定対象が無い（定義のみ。削除は可能）。値セルにツールチップで説明。

### 参照検出（削除前の安全確認）
- 集計表: `ViewSchedule.Definition.GetField(i).ParameterId` で軽量・確実に取得可（集計表を舐めるだけ）。ユーザーパラメータのみ（`id.IntValue() > 0`）。
- ビューフィルタ: `FilterRule` から参照パラメータを確実に取得するAPIが無く（既知の制約）未対応。
- タグ/数式: 参照情報がファミリ文書内にあり全ファミリを開く必要があるため重く実質不可＝未対応。
- よって「集計表参照」列のみ提供し、空でも他所で使用の可能性が残るため削除前警告は継続。

### Excel風の列並べ替え/フィルター（動的WPF）
- 各列ヘッダーを `DockPanel`(テキスト＋「▾」ボタン)に差し替え、`Popup` で昇順/降順＋値チェックリスト（検索・全選択/選択解除・(空白)）を表示。列ごとの許可値 `HashSet<string>` を `RowFilter` に合成。
- 長い値は `CheckBox.Content` を `TextBlock`(TextTrimming=CharacterEllipsis, MaxWidth) にして…省略＋全文ツールチップ、`Popup` の Border に MaxWidth を設定して横長化を防止。
- 一覧下の「全選択/選択解除」は `_view.Cast<object>()`（フィルター適用後の表示行）だけを対象に `ParamRow.IsSelected` を一括設定。

### ⚠️ 型シャドウイング/衝突のコンパイルエラー（ハマった）
- **CS0176**: `ParameterCleanupDialog : Window` は `Visibility` / `HorizontalAlignment` / `VerticalAlignment` という**インスタンスプロパティ**を継承する。これらと同名の列挙型を単純名(`Visibility.Visible` 等)で参照するとインスタンス経由の静的メンバーアクセス扱いで CS0176。→ `System.Windows.Visibility.Visible` のように**完全修飾**。
- **CS0104**: `using System.Windows.Controls;` と `using Autodesk.Revit.UI;` の併用で **`TextBox` が曖昧**（`System.Windows.Controls.TextBox` vs `Autodesk.Revit.UI.TextBox`）。`ComboBox` も同様に衝突しうる。→ 完全修飾。`Button`/`CheckBox`/`TextBlock` 等は Revit.UI に無く安全。
- `Color` は `System.Windows.Media.Color` と `Autodesk.Revit.DB.Color` が衝突。使う時は完全修飾。
- 教訓: Window 派生の WPF コードビハインドで列挙型/コントロール型を動的生成する際は、Revit の using と衝突しやすい名前（Visibility/HorizontalAlignment/VerticalAlignment/TextBox/ComboBox/Color/Point 等）を完全修飾する。

### ParameterCleanup: 削除時の Revit 警告ダイアログ抑制
- 複数パラメータ削除時、`Transaction.Commit()` で Revit が「これらの要素が削除されます」等の警告を**削除数だけ**モーダル表示する。
- 対策: 削除トランザクションに失敗ハンドリングを設定する。
  ```csharp
  var fho = t.GetFailureHandlingOptions();
  fho = fho.SetForcedModalHandling(false);
  fho = fho.SetClearAfterRollback(true);
  fho = fho.SetFailuresPreprocessor(new WarningSwallower());
  t.SetFailureHandlingOptions(fho);
  ```
  `WarningSwallower : IFailuresPreprocessor` は `PreprocessFailures` で `a.DeleteAllWarnings(); return FailureProcessingResult.Continue;`。
- `Set*` 系は結果を返す（コピーを返すことがある）ので必ず再代入すること。

## ExcelExportImport: カテゴリ名は Revit の Category.Name を単一の情報源にする（2026-07-24）

### 症状
エクスポートダイアログで、同一カテゴリなのに「カテゴリを選択」欄は「敷地」、「パラメータ／出力」欄と Excel シート名は「外構」と食い違う。

### 原因
- 「カテゴリ選択」欄だけ `CategoryInfo.DisplayLabel` → `CategoryLocalizer.GetLocalizedName(BuiltInCategory, Name)` という**固定翻訳表**を使っていた（例: `OST_Site`→Loc "Category.Site"→「敷地」）。
- 一方、パラメータ/出力欄は `ExportDialog.xaml.cs` が `cat.Name`（Revit 実名）を `GetParametersForCategory` に渡しており、Excel 出力（`ExcelExportService`）もシート名・ヘッダーに `category.Name` を使用。
- Revit の実際のカテゴリ名は**バージョン/言語パックで変わる**（このモデルの実名は「外構」）。ハードコードの固定翻訳「敷地」がそれとずれ、選択欄だけ別名になった。

### 修正
- `CategoryInfo.DisplayLabel` を `=> $"{Name} ({ElementCount})"`（Revit 実名）に変更。
- 唯一の利用箇所が消えた `CategoryLocalizer.cs` を撤去（Strings の `Category.*` キーは未使用として残置・無害）。

### 教訓・再発防止
- **カテゴリ名は Revit の `Category.Name` を単一の情報源にする。** 固定翻訳表で二重管理すると、Revit のバージョン/言語差で必ずどこかがずれる。
- パラメータ値・ファミリ名など他の表示も Revit 言語で出るため、カテゴリ名だけ別言語に訳すのはむしろ不整合。実名統一が最も安全。

## ExcelExportImport: エクスポートダイアログに全選択/選択解除ボタンを追加（2026-07-26）

### 追加内容
- 「カテゴリを選択」欄・「パラメータ」欄の検索ボックス下に **全選択 / 選択解除** ボタンを設置。いずれも**現在の検索絞り込みで表示中の項目だけ**を対象に一括チェック/解除する。
- ※当初はパラメータ欄に種別フィルタ用コンボ（すべて/インスタンス/タイプ）も追加したが、既存のテキスト検索で足りるため削除した。

### 実装メモ
- `CategoryInfo` / `ParameterInfo` の `IsChecked` を **`INotifyPropertyChanged`** 化。コードから一括変更した際に CheckBox へ即時反映させるため（再バインドに頼らず、仮想化された `ParameterListBox` でも整合）。
- カテゴリの一括変更中は `_suppressCategoryUpdate` フラグで `CategoryCheckBox_Changed`→`UpdateParameterList()`（Revit へのパラメータ取得）の連続発火を抑制し、最後に1回だけ更新。
- 絞り込みロジックは `GetVisibleCategories()` / `GetBaseParameters()`＋`ApplyParameterFilters()`（テキスト検索）に集約し、表示と全選択対象で同じ判定を共有。

## ExcelExportImport: 出力欄の▲▼で選択と違うパラメータが動く不具合を修正（2026-07-27）

### 症状
出力セクションでパラメータを選択し▲/▼を押すと、選択したものではなく別のパラメータが移動することがある。

### 原因
- 出力リストは `RefreshOutputList()` で `CategoryName` によりグループ化して表示している。
- 一方 `MoveUpButton_Click`/`MoveDownButton_Click` は **表示上のインデックス** `OutputListBox.SelectedIndex` で内部リスト `_outputParameters` を直接 `RemoveAt/Insert` していた。
- カテゴリが複数あると表示順（グループ順）と `_outputParameters` の並びがずれるため、`SelectedIndex` の位置には別カテゴリの要素があり、それが動いてしまう。テキストフィルタ適用時も同様にずれる。

### 修正
- `MoveSelectedOutput(direction)` に集約。`OutputListBox.SelectedItem`（選択中の要素そのもの）を基準に `_outputParameters.IndexOf` で実インデックスを取得。
- 表示がカテゴリ単位でグループ化されているため、**同一カテゴリ内で移動方向の隣にある要素**を探して入れ替える（グループをまたいで“何も動かない”誤操作感を防ぐ）。
- 移動後は同じ要素を再選択し `ScrollIntoView` する。

### 教訓
- グループ化/フィルタ表示のリストで並べ替えるときは、**表示インデックスで元データを触らない**。必ず選択オブジェクト参照から元データの位置を引き直す。

## ExcelExportImport: 出力欄の▲▼移動でスクロール位置が先頭に飛ぶ不具合を修正（2026-07-27）

### 症状
出力欄で上下移動すると、表示位置がずれてスクロールバーが動いたように見える（先頭付近へジャンプ）。

### 原因
- `MoveSelectedOutput` が移動のたびに `RefreshOutputList()` を呼び、**毎回 `OutputListBox.ItemsSource` に新しい `ListCollectionView` を代入**していた。
- ItemsSource を差し替えるとスクロールは先頭(offset 0)にリセットされ、その直後の `ScrollIntoView(selected)` が選択項目まで大きくスクロールし直すため、視覚的に大きく飛ぶ。

### 修正
- 出力欄の表示を **`ObservableCollection<ParameterInfo> _outputDisplay`** で保持し、`ListCollectionView`（グループ化）は**一度だけ生成して使い回す**（`_outputView`）。`RefreshOutputList()` は中身を差し替えるだけにした。
- 上下移動時は ItemsSource を作り直さず、**`_outputDisplay.Move()` で該当2要素だけを移動**。ObservableCollection の Move 通知でコンテナが移動するだけなので**スクロール位置が保持**される。
- `ScrollIntoView` は「表示中なら何もしない／枠外なら最小限スクロール」の挙動になるため、そのまま呼んで**枠の上下に出たときだけ追従**するようにした。
- フィルタ等で相手要素が非表示のときのみ従来どおり `RefreshOutputList()` にフォールバック。

### メモ
- `_outputDisplay` のフラット順は、複数カテゴリをまたぐ Move で `_outputParameters` とずれ得るが、表示はカテゴリでグループ化され**カテゴリ内の相対順のみ**が意味を持つため表示は一致する。エクスポート順の情報源はあくまで `_outputParameters`。

## ExcelExportImport: 書き出し済みExcelから出力設定を読み込む機能を追加（2026-07-27）

### 背景・目的
出力設定の再利用は従来 JSON 保存/読込のみ。設定JSONを保存し忘れた場合や、他ユーザーが書き出した Excel の並びをそのまま使いたい場合に不便だった。そこで**エクスポート済み Excel から出力設定（対象カテゴリ・出力パラメータ・順序）を復元**できるようにした。

### 実装
- `Services/ExportSettingsExcelReader.cs` を新規追加。`ExcelExportService` の出力形式を前提に `.xlsx` を読み、`ExportSettings` を組み立てる。
  - 1行目ヘッダー 3列目以降を `I-`/`T-` プレフィックスでパースし `RawName`/`IsTypeParameter` を復元（`(*変更不可)` サフィックスは除去）。
  - カテゴリ名は**2列目（実カテゴリ名）**を情報源にする（シート名はサニタイズ/31文字切り詰めされ得るため）。
  - カテゴリ毎シート分割: シート内カテゴリ＝1つ → ヘッダー全列をそのカテゴリに割り当て。
  - 1シート統合（"データ"）: カテゴリ複数 → 「そのカテゴリの行に非空セルがある列」をそのカテゴリのパラメータとみなす。
- `ExportDialog` に「Excel設定読込」ボタンを追加。読込後は既存の `ApplySettings()` に流し込むため、現在ドキュメントのカテゴリ/パラメータに対して照合・復元される（JSON読込と同じ経路）。
- 出力列が1つも解釈できなければ警告（アドイン以外の Excel 対策）。現在の対象に存在せず復元できなかった分があれば `{0}/{1}` 件で通知。

### 既知の制限
- 1シート統合形式で、あるカテゴリの列が**全行空**だと非空判定で漏れて復元されない（分割形式では問題なし）。実運用では分割形式が既定のため影響は小さい。

## ExcelExportImport: 設定ボタンを整理（設定復元を廃止・設定系を出力欄の行へ集約）（2026-07-27）

### 変更
- **「設定復元」ボタンを廃止**。実態は「全カテゴリのチェックを外し出力欄を空にする＝初期化」で、出力欄の「クリア」と役割が重複していたため（ユーザー判断で統合）。
- 「設定読込」「Excel設定読込」「設定保存」の3ボタンを**出力欄の「クリア」ボタンの右**に横並びで集約（コンパクトな `SmallButtonStyle`・短いラベル）。ダイアログ下部の設定エリアは撤去し OK/キャンセルのみに。
- 各ボタンの説明は**常時表示のラベルをやめ、ホバー時のツールチップ**に変更（`ToolTip` に `*.Desc` を設定）。
- 縦並びエリア撤去に伴いウィンドウ高さを 600→520 に戻す。

### 後片付け
- `ResetSettingsButton_Click` と `Export.RestoreSettings` / `Export.RestoreSettings.Desc` / `Export.LoadSettingsFromExcel.Tip` を削除。ラベルキー（LoadSettings/LoadSettingsFromExcel/SaveSettings）は短縮値に更新。

### 追記（2026-07-27）
- 「出力Excelをカテゴリ毎にシートに分ける」チェックボックスを、ダイアログ下部のオプション行から**出力セクション内（クリア/読込/Excel読込/保存ボタンの下）へ移動**。出力関係の機能を出力セクションに集約する意図。
- 出力欄は横幅が狭いためラベルは `TextBlock`＋`TextWrapping` で折り返し。多言語更新は `SplitByCategoryLabel.Text`（旧 `SplitByCategoryCheckBox.Content`）へ。

## ExcelExportImport: エクスポート実行への改名・Excel設定読込の高速化・インポートのシート情報削除（2026-07-27）

### 1. OK ボタンを「エクスポート実行」に改名
- インポート側（「インポート実行」）と表現を統一。専用キー `Export.RunExport`（JP=エクスポート実行 / EN=Export / CN=执行导出）。

### 2. Excel設定読込（`ExportSettingsExcelReader`）の高速化
- 症状: 大きな書き出し済み Excel だと読込が非常に遅い。
- 原因: カテゴリ毎シート（既定）でも**全データ行の2列目をループ**しており、ClosedXML の重いセルアクセスが行数分走っていた。
- 対処: 既定の分割形式では **1行目ヘッダー＋先頭データ行(2行目)の2列目だけ**を読む方式に変更（`Row(1).CellsUsed()` で1行目のみ、カテゴリ名は `Cell(2,2)` 1回）。全行走査を排除。
- 1シート統合（"データ"）だけは複数カテゴリ混在のため従来どおり走査（`ReadSingleSheet`）。既定ではないため影響は限定的。

### 3. Excelインポートダイアログの「シート情報」を削除
- 実用上不要な情報のため `grpSheetInfo` / `SheetInfoText` を撤去。`LoadPreview` は `GeneratePreview(doc, path)`（out なし）を使用。未使用化した `Import.SheetInfo` / `Import.SelectFile` キーも削除。

## ExcelExportImport: Excel設定読込が列入れ替え時に読めない不具合を修正（2026-07-27）

### 症状
書き出し済み Excel の列を手動で入れ替え・並べ替えたファイルから「Excel読込」すると、出力パラメータが読み込まれない。

### 原因
`ExportSettingsExcelReader` が**列位置を固定で仮定**していた。
- `col < 3` で1・2列目をスキップ（＝1列目=要素ID, 2列目=カテゴリ 前提）。
- カテゴリ名を常に2列目 `Cell(2,2)` から読む。

列を入れ替えると「カテゴリ」列が2列目からずれ、別の値（パラメータ値や要素ID）をカテゴリ名として読む → `ApplySettings` の照合キー `RawName|IsType|CategoryName` が一致せず、カテゴリも選択されないため何も復元されない。パラメータ列が1・2列目に移動した場合も取りこぼしていた。

### 修正
- 列位置ではなく**見出し文字列で列を特定**する方式に変更。
  - 1行目の各セルを走査し、`"要素ID"` は除外、`"カテゴリ"` の列番号を記録、それ以外で `I-`/`T-` プレフィックスを持つ列をパラメータ列とする（位置非依存）。
  - カテゴリ名は記録した「カテゴリ列」から読む。見つからない場合のみシート名（分割形式）／2列目（統合形式）にフォールバック。
- パラメータの並び順は1行目の左→右の並び（＝ユーザーが入れ替えた順）で復元される。

## ExcelExportImport: 出力欄で複数選択したパラメータの上下移動に対応（2026-07-28）

### 内容
- 出力欄 `OutputListBox` の `SelectionMode` を `Single`→`Extended` に変更し、複数選択を可能に。
- `MoveSelectedOutput` を複数選択対応に刷新。
  - 選択に含まれる**カテゴリごと**に、そのカテゴリ内の並び(`seq`)だけを対象に「選択ブロックを1つ分ずらす」`ReorderBlock` を適用（隣が選択なら詰めずブロック保持、端で停止）。
  - 並べ替え結果は元の位置へ書き戻すため、他カテゴリの位置は不変。
- 表示同期 `SyncDisplayToParameters` を追加。フィルタ無しなら `_outputDisplay` を `Move` で `_outputParameters` の並びへ合わせ**スクロール位置を維持**。フィルタ中のみ作り直し。
- ついでに「出力から削除(＜)」も複数選択一括削除に対応。

### 補足
- 選択は要素参照で追跡するため、表示位置とのズレや飛び選択でも常に選択したものが動く。
- 移動は従来どおりカテゴリ（＝出力シート）内での順序調整として機能する。

## ExcelExportImport: 長さ等の数値インポートが単位変換されず反映されない不具合を修正（2026-07-28）

### 症状
部屋カテゴリの「上限オフセット（上部レベルからのオフセット）」を Excel で 0→-100 に変更してインポートしても反映されない。Revit 上で直接 -100 と入力すれば変更できるため、変更不可の値ではない。

### 原因
`ParameterService.SetParameterValue` の `StorageType.Double` 処理:
1. まず `param.SetValueString(value)` を試す。
2. 失敗したら `double.TryParse` して `param.Set(dblVal)` を**生値のまま**設定。

部屋の上限オフセット等の一部パラメータは `SetValueString` が正しく反映されず（false 相当）、フォールバックで **Excel の数値（表示単位=mm）を内部単位（feet）として** `Set` してしまう。`-100` は「-100ft ≒ -30480mm」となり部屋として無効なため Revit に拒否され、結果として値が変わらない＝反映されない状態になっていた。

### 修正
- 数値として解釈できる場合は **表示単位→内部単位へ変換してから設定**する方式に変更。
  - `param.GetUnitTypeId()` で表示単位を取得し `UnitUtils.ConvertToInternalUnits(値, 単位)` で内部単位へ変換して `param.Set()`。
  - 単位を持たない数値パラメータ（比率・個数等）は変換せず生値で設定（`GetUnitTypeId` が空/例外の場合）。
  - `param.Set()` の戻り値をそのまま返し、実際に失敗した場合は失敗として報告。
- 数値として解釈できない書式付き文字列のみ従来どおり `SetValueString` にフォールバック。
- エクスポートは `AsValueString()`（表示単位）で書き出しているため、インポートの表示単位→内部単位変換と対称になり整合する。

### 補足
- 長さ・面積・体積・角度など単位を持つ全 Double パラメータに効く（部屋オフセットに限らない）。
- 単位 API（`GetUnitTypeId` / `ForgeTypeId` / `UnitUtils.ConvertToInternalUnits`）は Revit 2021 以降共通のため全対応バージョンで動作。

## ExcelExportImport: 制約エラー(部屋の高さ>0 等)で全体ロールバックされる問題を改善（2026-07-28）

### 症状
部屋の上限オフセットに -100 をインポートすると「部屋 の高さは 0 よりも大きい必要があります」（無視できないエラー）が多数出て、**1トランザクションのため全体がロールバック**され、有効な行も含め何も反映されない（Revit のエラーダイアログも表示される）。

### 原因
- インポートは 1 つの `Transaction` で全変更を書き込み、`SuccessCount>0` なら `Commit`。
- Revit の「無視できないエラー」（例: 部屋高さ ≤ 0）が commit 時に発生すると、そのトランザクションは**全体がロールバック**される。1件でも該当すると全件反映されない。

### 修正
- 失敗ハンドラ `ImportFailurePreprocessor`（`IFailuresPreprocessor`）を追加。
  - 警告は自動削除してコミットを妨げない。
  - 無視できないエラーは原因要素ID を収集し `ProceedWithRollBack` を返す（Revit のダイアログも抑止）。
- `ExcelImportCommand` を**多段パス**方式に変更。
  - 各パスでプリプロセッサ付きトランザクションを開始→書き込み→コミット。
  - エラーでロールバックされたら、収集した失敗要素を除外して再試行（最大5パス、新たに除外できる要素が無くなれば終了）。
  - これにより**制約を満たす要素だけがコミット**され、満たせない要素はスキップして結果に「◯個の要素は Revit の制約により変更できずスキップ」と報告・赤色マーキング。
- `ExcelImportService.ImportFromPreview` に除外要素ID指定のオーバーロードを追加。

### 単位変換のフォールバック強化（同時対応）
- Double パラメータ設定を「単位変換 Set →（不可なら）SetValueString →（不可なら）生値 Set」の順に整理。生値 Set（内部単位扱い）を最後の手段に下げ、単位付き数値が誤変換されないようにした。

## ExcelExportImport: ElementId型（部屋の上部レベル等）の名前設定に対応＋インポート診断ログ追加（2026-07-28）

### 症状
部屋の「上部レベル」をExcelで変更してインポートすると「パラメータ 'I-上部レベル' の値設定に失敗（値: '1FL(2FL)'）」となる。

### 原因
`ParameterService.SetParameterValue` の `StorageType.ElementId` 処理が「`SetValueString` か 数値ID」のみ対応で、**レベル名（"1FL(2FL)" 等）から Level 要素を検索して設定する処理が無かった**。エクスポートは ElementId 参照を要素名（`AsValueString`）で書き出すため、名前解決が必須。

### 修正
- `SetParameterValue(param, value, doc)` に `Document` を追加（要素検索に必要）。呼び出し側（`ExcelImportService`）も更新。
- `SetElementIdParameter`：数値ID → 名前検索 → `SetValueString` の順に解決。
- `ResolveElementIdByName`：現在値の要素クラス（上部レベルなら `Level`）に絞って名前一致を検索し、無ければ `Level` を検索。`Element.Name` の例外は `SafeName` で吸収。
- これで「上部レベル＝レベル名」のインポートが反映される（設定後に高さ制約エラーになる部屋は、既存の失敗プリプロセッサでスキップ）。

### 診断ログの追加
ユーザー環境での原因切り分けのため、`DiagLog`（`C:\temp\Tools28_debug.txt`）へ詳細を出力。
- `ExcelImportCommand`：変更対象件数、各パスの `commit` 状態・成功/失敗数・制約エラー要素数、最終結果。
- `ExcelImportService.ImportFromPreview`：各セルの `elem/param/storage/current/new` と設定結果。
- `ParameterService`：ElementId 名前解決の成否、例外内容。
- 問題再現時はこのログを送ってもらえば失敗理由（名前不一致・制約エラー等）を特定できる。

<a id="Room3DColor"></a>
## Room3DColor 開発知見

### 背景・Revit の仕様上の制約
- Revit では **Room（部屋）は 3D ビューに標準では描画されない**（面・タグは平面/断面用）。3D で部屋を可視化するには立体形状を別要素として生成する必要がある。
- 実現手法として **`Room.GetClosedShell()`（watertight な閉じたソリッドを返す）→ `DirectShape.SetShape()`** を採用。Autodesk 公式（Jeremy Tammik / RoomVolumeDirectShape）で確立された方式。
- 生成した DirectShape は **汎用モデル（OST_GenericModel）** カテゴリになる。

### 実装の要点（`Commands/Room3DColor/`）
- `RoomSolidGenerator`：部屋収集（`Location != null && Area > 0`）、体積計算の有効化（`AreaVolumeSettings.ComputeVolumes`）、DirectShape 生成、専用等角3Dビュー作成（`View3D.CreateIsometric`）、表示絞り込み。
- `RoomColorManager`：色分け基準（部屋名/レベル/パラメータ/部屋ごと）でのグルーピングと色パレット生成（BeamTopLevel と同系統のパステル）、`OverrideGraphicSettings`（サーフェス前景ソリッド塗り＋色）作成。
- 色付けは **DirectShape へ要素単位で `view.SetElementOverrides`** を適用（フィルタではなく直接オーバーライド）。生成要素を自分で把握しているため。
- 「部屋のみ表示」：専用ビューで汎用モデル以外の全カテゴリを `SetCategoryHidden`、さらに今回生成分以外の汎用モデル要素を `HideElements` で個別非表示。
- 再生成用に DirectShape の Comments に識別マーカー `Tools28_Room3DColor` を保存し、削除オプションで一括削除。

### トランザクション構成
- `TransactionGroup` 内で「体積計算の有効化」を**別トランザクション**で先に commit（ジオメトリ再計算を確定してから `GetClosedShell` を呼ぶため）→ 続くトランザクションで生成・色分け・凡例。

### 注意点
- 非囲繞・体積0の部屋は `GetClosedShell` が有効なソリッドを返さずスキップ（完了メッセージに件数表示）。
- 体積計算はモデル全体設定のため、有効化した場合はその旨を通知。
- 表示スタイルは `DisplayStyleType.HLR` に設定（サーフェス前景パターンの色が確実に見える）。

### ワークセットによる表示制御（追加対応）
永続要素（DirectShape）が他ビューに出ないよう、専用ワークセットで制御する。
- **ワークシェアあり**: `Workset.Create`（要トランザクション）で専用ワークセット「部屋3D色分け」を作成/取得し、`ELEM_PARTITION_PARAM` で各 DirectShape を割当。`WorksetDefaultVisibilitySettings.SetWorksetVisibility(id, false)` で既定を非表示にする（新規ワークセットは既存ビューもこの既定に従うため、全ビューで非表示になる）。専用色分けビューのみ `View.SetWorksetVisibility(id, WorksetVisibility.Visible)` で表示。
- **ワークシェアなし**: ワークセットが作れないため、`Element.CanBeHidden(view)` で判定しつつ専用ビュー以外の全ビューで `View.HideElements` により要素単位で非表示（同等動作）。この処理以降に作成された新規ビューには出てしまう制約あり。
- `WorksetId.IntegerValue` は 2026 でも維持されている（`ElementId.IntegerValue` のみが `Value` に変更）ため、`ELEM_PARTITION_PARAM.Set(worksetId.IntegerValue)` はバージョン分岐不要。

- 2026-07-28: Room3DColor を Revit 2022 / 2024 でオートビルド確認（DisplayStyle / ClosedShell のAPI名修正後）。

## ExcelExportImport: インポート診断ログの強化＋ログ送信スクリプト追加（2026-07-28）

### 目的
`I-イメージ`（値 '仕上表非表示'）のインポート失敗など、原因を憶測でなく実ログで特定するため。

### 内容
- `ImportFromPreview` の診断ログに **StorageType / IsShared / IsReadOnly / BuiltInParameter / 現在値 / 新値** を追加（`C:\temp\Tools28_debug.txt`）。これで失敗パラメータの正体（組込みか共有か、ElementId/String 等）を確定できる。
- `Send-Tools28Log.ps1` を追加。`C:\temp\Tools28_debug.txt` を `.diag\` にコピーして git push し、Claude が実ログを直接読めるようにする（Formwork の仕組みに倣う）。

## ExcelExportImport: 「イメージ」(ALL_MODEL_IMAGE) 等 ElementId参照の名前解決を拡張（2026-07-28）

### 根拠（実ログ C:\temp\Tools28_debug.txt）
```
param='I-イメージ' storage=ElementId shared=False readonly=False bip=ALL_MODEL_IMAGE current='' new='仕上表非表示'
[SetParam] ElementId 名前解決失敗 '仕上表非表示'（該当要素が見つからない）
```
→ `I-イメージ` は組込み `ALL_MODEL_IMAGE`（ラスター画像 `ImageType` への参照）。前回の名前解決は `Level` しか探しておらず、画像名 `仕上表非表示` に対応する `ImageType` を引き当てられず失敗していた（憶測ではなくログで確定）。

### 修正
- `ResolveElementIdByName` を対象クラス優先リスト方式に拡張:
  1. 現在値の要素クラス（値の入れ替えで最も確実）
  2. パラメータ種別からの推定：`ALL_MODEL_IMAGE`/`ALL_MODEL_TYPE_IMAGE` → `ImageType`
  3. 汎用頻出クラス：`Level` / `Material`
- `GetBuiltInParameter`（`InternalDefinition.BuiltInParameter`）で画像パラメータを判定。
- これで「イメージ＝画像名」のインポートが、該当 `ImageType` が存在すれば反映される。存在しない場合は従来どおり失敗として報告（＝データ側の問題を切り分けられる）。

## ExcelExportImport: 要素参照型（イメージ等）の失敗メッセージを明確化（2026-07-28）

### 背景
`イメージ`（`ALL_MODEL_IMAGE`）は**画像参照（ElementId型）**で、文字値は設定できない（Revitの仕様。UIでも画像ピッカーのみ）。ユーザーが文字 '仕上非表示' を入れようとして「値設定に失敗」とだけ出て理由が分かりづらかった。

### 修正
- `BuildSetFailureMessage` を追加し、失敗理由を型に応じて明示:
  - 画像参照（`ALL_MODEL_IMAGE`/`ALL_MODEL_TYPE_IMAGE`）→「画像参照型のため文字値は設定できません。その名前の画像が必要です」
  - その他 ElementId 参照 →「要素参照型のため文字値は設定できません。その名前の要素が見つかりません」
  - それ以外 → 従来メッセージ
- 文字フラグを持たせたい場合はテキスト型の共有/プロジェクトパラメータを使う、という運用切り分けが可能に。

## ExcelExportImport: 書き出し時に「文字を入れられない列」を一目で分かるよう明示（2026-07-28）

### 背景
`イメージ`（画像参照）のように文字値を入れられないパラメータに、インポート後に気づくと手戻りになる。書き出した時点で分かるようにしたい（画像だけでなく他の要素参照型・読取専用も）。

### 対応
- **ヘッダーにマーカー＋灰色**で明示（`ExcelExportService.WriteHeaderRow`）。書き出し時点で「この列は文字を入れても取り込めない/名称が必要」が一目で分かる。
  - 画像参照（`ALL_MODEL_IMAGE`/`ALL_MODEL_TYPE_IMAGE`）→ `(*画像参照/文字不可)`
  - 要素参照（`StorageType.ElementId`：レベル/材料等）→ `(*要素参照/名称必須)`（既存要素名が必要、自由文字は不可）
  - 読取専用 → `(*変更不可)`
  - 「タイプ」はタイプ名で変更できる運用のためマーカー無し
  - 対象列はヘッダーセルを灰色＋斜体にして視覚的に強調
- マーカーの定義と除去を **`ParameterHeaderMarker`** に一元化。インポート/設定読込/セル色付け（ClosedXML・COM）すべてがこの `Strip` を使うよう統一（マーカーが増えても取りこぼさない）。
- `ParameterInfo` に `IsElementReference`/`IsImage` を追加し、`ParameterService.CollectParameters` で `StorageType`・`BuiltInParameter` から判定して設定。

### 効果
文字を入れられない列が書き出し時点で分かるため、インポート実行後の手戻りを防げる。失敗時メッセージ（前回改善）と合わせて、原因の事前把握と事後説明の両方をカバー。

## ExcelExportImport: 文字設定不可の列はデータ値もグレー文字に（2026-07-28）
- 灰色ヘッダー（画像参照/要素参照/変更不可）の列について、2行目以降のデータ値の文字色もグレー（150,150,150）にして目立たなくし、編集可能な列と視覚的に区別できるようにした（`ApplyNonEditableColumnStyle`）。分割/統合の両出力に適用。

## ExcelExportImport: ヘッダーのマーカー文字を8ptに（2026-07-28）
- 灰色ヘッダーのマーカー（`(*変更不可)`/`(*画像参照…)`/`(*要素参照…)`、（）含む）だけをフォントサイズ8にして控えめに。パラメータ名は通常サイズのまま。ClosedXML のリッチテキスト（`GetRichText().AddText(...).SetFontSize(8)`）でセル内の一部だけ小さくした。

## ExcelExportImport: 同名パラメータ（例: エリアの「用途」×2）が1つしか出ない不具合を修正（2026-08-03）

### 症状
カテゴリ「エリア」には同名の「用途」パラメータが2つ存在するのに、エクスポートダイアログのパラメータ欄には「用途」が1つしか表示されなかった。

### 原因
パラメータを **名前だけをキーにして重複排除** していた（実体が別物でも名前が同じだと片方が捨てられていた）。名前ベースの排除が2箇所に存在:
1. 収集: `ParameterService.GetParametersForCategory` の `HashSet<ParameterInfo>`。`ParameterInfo.Equals/GetHashCode` が `DisplayName`(=I-/T-＋名前)＋`CategoryName` ベース → 同名の2つが同一視され2つ目が破棄。
2. 表示: `ExportDialog.GetBaseParameters` の `GroupBy(DisplayName + "|" + CategoryName)`。
   さらに書き出し(`ExcelExportService.ExportSingleSheet`)・設定復元・Excel設定読込・インポート照合(`LookupParameter` は最初の1件のみ)まで、全体が名前を識別子にしていた。

### 対応（A案: 安定した識別子で一貫して区別）
- `ParameterInfo` に **`ParamId`（Parameter.Id を long 化した安定 ID）** と **`Kind`（組み込み/共有/プロジェクト）** を追加。等価判定を `ParamId`＋種別(I-/T-)＋カテゴリに変更 → 同名でも別物なら両方残る。
- 同カテゴリ・同名グループには **曖昧性解消接尾辞** を付与（`ParameterKindHelper.BuildSuffix`）。例: `I-用途【組み込み】` / `I-用途【共有】`。同一種別で複数ある場合は ParamId 昇順で `#連番`。`DisplayName` は `Prefix + RawName + DisambigSuffix` の算出プロパティにし、ダイアログ表示・Excelヘッダー・重複排除の単一情報源に統一。
- **エクスポート**は `ParamId` で該当パラメータを厳密解決（`ParameterService.FindByIdentity`）。タイプ値キャッシュキーにも `ParamId` を追加。
- **インポート**は Excel に残るヘッダー文字列から種別・連番を解析（`ParameterService.ParseDisplayName` → `ParameterKindHelper.TryExtractSuffix`）し、種別で候補を絞って解決（`FindParameter` の種別対応オーバーロード）。接尾辞が無い通常列は従来どおり `LookupParameter`（＝旧 Excel と後方互換）。
- **種別ラベルは言語非依存で往復**: 表示は `Loc.S("Export.ParamKind.*")`（JP=組み込み/共有/プロジェクト、EN=Built-in/Shared/Project、CN=内置/共享/项目）。`ParameterKindHelper` は全言語の逆引き表を持ち、書き出し言語と取り込み言語が違っても解析可能。
- **設定保存/復元**: `ExportParameterEntry` に `DisplayName` を追加し、復元は DisplayName 一致を優先（旧設定＝DisplayName 無しは RawName にフォールバック）。Excel設定読込(`ExportSettingsExcelReader`)も DisplayName で重複排除。

### 効果
同名パラメータが全て列挙・書き出し・取り込みでき、値が一方に偏る不整合が解消。接尾辞は重複時のみ付くため、重複が無い通常パラメータのヘッダー／既存 Excel との互換は維持。

## DwgLayerTransfer: DWGレイヤ表示設定をモデル間で直接移行（2026-08-05）

### 背景
リンク/読み込みした DWG のレイヤ表示設定は、**表示/グラフィックスの上書き（V/G）の「読み込みカテゴリ」タブ**で行う。
これは**ビュー単位**の設定のため、Revit 標準の「プロジェクト標準の転送」では移せない
（転送で移せるのは**オブジェクトスタイル＞読み込まれたオブジェクト**＝モデル全体の方だけ）。
同じ DWG を使う別モデルやモデル分割時に、レイヤ表示設定を手作業でやり直す手間が発生していた。

### 方針
- **Excel などの中間ファイルを使わず**、開いている 2 つのドキュメント間で直接受け渡す
  （ユーザー要望: 手数最小化・無駄なデータ書き出しの削減）。
- 移行先は常にアクティブなドキュメント。移行元は `Application.Documents` から
  リンク・ファミリを除いた他の開いているモデルを選ぶ。

### DWG カテゴリとレイヤの取り方
- 読み込み DWG は「DWGファイル名」のカテゴリとして登録され、**各レイヤがサブカテゴリ**になる。
- 列挙は `Document.Settings.Categories` のうち **`Parent == null` かつ ID が正**のものを採る。
  組み込みカテゴリの ID は負なので、この 2 条件だけで読み込み DWG のカテゴリを抽出できる。
  （`ImportInstance` 経由でも取れるが、インスタンスが無いケースを取りこぼす）
- リンク/読み込みの別は `ImportInstance.IsLinked` から補足（表示用のみ）。
- `ElementId` の int 化は既存の `RevitCompatibility.IntValue()` を使用（2026 で `IntegerValue` 廃止）。

### 設定の読み書き
- 読み: `View.GetCategoryHidden(id)` / `View.GetCategoryOverrides(id)`
- 書き: `View.SetCategoryHidden(id, bool)`（`CanCategoryBeHidden` で事前確認）/ `View.SetCategoryOverrides(id, ogs)`
- 保持する項目は **V/G「読み込みカテゴリ」タブで実際に編集できるものだけ**に絞った:
  表示/非表示・投影線（色/線幅/線種）・切断線（色/線幅/線種）・ハーフトーン。
  **サーフェスパターン・透明度は DWG のジオメトリに効かない**ため意図的に対象外
  （移行先の設定を無用にリセットしないため）。

### ElementId はモデル間で通用しない → すべて名前で照合
| 対象 | 照合方法 |
|--|--|
| DWG | カテゴリ名（= ファイル名）。同名を自動対応、UI で手動変更可 |
| レイヤ | サブカテゴリ名の一致 |
| ビュー/テンプレート | 名前の一致。UI で手動変更可 |
| 線種 | `LinePatternElement.Name` の一致。実線は `LinePatternElement.GetSolidPatternId()` を予約名 `<solid>` で表現 |

未解決（移行先に無いレイヤ・線種）は結果ダイアログに一覧表示する。黙って落とさない。

### ⚠️ ビューテンプレートによる V/G 制御の判定（ハマりどころ）
`View.GetNonControlledTemplateParameterIds()` は「テンプレートが**制御していない**項目」を返す。
`BuiltInParameter.VIS_GRAPHICS_IMPORT` が**含まれていなければ制御下**、という反転した判定になる。

2 方向のチェックが必要で、両方とも実装している:
1. **ビュー単位で移行するとき** — 移行先ビューがテンプレート制御下だと `SetCategoryOverrides` が効かない
   → 該当ビューは状態列に理由を出しチェック不可にする（テンプレート単位への切替を促す）
2. **テンプレート単位で移行するとき** — 移行先テンプレートが読み込みカテゴリを制御していないと
   書き込んでもビューに反映されない → 同様に警告する

どちらも黙って no-op になるのが最悪なので、事前に UI へ出す設計にした。

### その他の実装メモ
- `AreGraphicsOverridesAllowed()` は**ビューテンプレートには使えない**ため、テンプレートの除外は
  `ViewSchedule` / `ViewSheet` の型判定で行う。
- DataGrid の `CellTemplate` 内 ComboBox の選択変更中に `Items.Refresh()` を呼ぶと編集トランザクションと
  衝突して例外になる → `RefreshGridDeferred()` で `CommitEdit` 後に Background 優先度で遅延実行。
  （FilterManagement のセル編集と同じ考え方）
- 同名の DWG カテゴリが複数存在しうるため、対応表の辞書化に `ToDictionary` は使わない（後勝ちで詰める）。
- 適用は移行先ドキュメントの**単一トランザクション**。実行後は Ctrl+Z で丸ごと戻せる。

## 最終セッション: 2026-08-28（リリース前の全バージョンビルド）
v2.1.2 以降、未リリースのまま 2.2〜2.6 の5世代分が溜まっている。
次回リリースは **v2.6** を想定。リリース前の動作確認のため全6バージョン
（2021〜2026）をビルドする（コミット件名に `[build:all]`）。

前回リリース（v2.1.2）以降の主な内容:
- 新機能8つ: 部屋3D色分け / 塗潰しパターン入出力 / パラメータ整理 /
  フィルタ整理 / テンプレート整理 / DWGレイヤ表示設定の移行 /
  一般モデル化 / 自動バックアップ
- 全21機能の監査と不具合修正（切断ボックス誤ペースト、言語設定が保存されない、
  シート連番、バックアップ世代削除ほか）
- 多言語化の積み残し解消（トランザクション名29・message39ほか）
- リリース時のバージョン反映を CI に一元化
- アイコンの青を Revit に合わせた（3機能のみ）

⚠️ リリース実行前に必要な作業:
- Licensing/ExpiryManager.cs の ExpiryDate をリリース日の1年後に更新
  （現在 2027-06-01。CI が残り30日以下で失敗、180日以下で警告）
- Packages/*/README.txt のリリースノート本文を手で追記
  （タイトル行のバージョンは CI が自動で書き換える）
## 前回セッション: 2026-08-28 03:18
変更ファイル: Application.cs,CLAUDE.md,Commands/GenericModelMerge/GenericModelMergeCommand.cs,Commands/GenericModelMerge/Models/MergeCategoryRow.cs,Commands/GenericModelMerge/Models/MergeOptions.cs

**GenericModelMerge（一般モデル化）** を新規追加。3Dビューに表示されている複数カテゴリの
要素から形状を読み取り、一つの一般モデルのかたまりを作成する。出力はダイレクトシェイプと
一般モデルファミリ(.rfa)の2択、まとめ方は「形はそのまま／すべてつなげる／くっついているものだけ」
の3択をダイアログで切替。リボンに新規パネル「モデル」を新設。

## 前回セッション: 2026-08-28（全機能監査）
全 21 機能を監査し、優先度の高い不具合を修正。
切断ボックスの誤ペースト（未コピーでもペーストできて解除される）、言語設定が
一般ユーザー権限で保存されない問題、コピー＆ペースト系の無通知な部分失敗と
全ビューアクティブ化、シート連番のプレフィックス無視、バックアップ世代削除の
並び順を修正。多言語化（トランザクション名29・message39ほか）と
リリース時のバージョン反映を CI に一元化。

## 前回セッション: 2026-08-28（パラメータ整理）
**ParameterCleanup（パラメータ整理）** の値判定をバインド非依存に刷新。
「バインドなし」表示のパラメータ（ファミリ内部定義の共有パラメータ）でも要素側から
実際の値を読み取るようにし、値が入ったまま削除してしまう事故を防止。
使用箇所索引（`ParameterUsageIndex`）を新設し、モデル走査は1スキャンにつき1回に集約。
値の状態を「値あり／空／未使用」の3種に分離し、「使用箇所」列と削除前警告を追加。

## 前回セッション: 2026-08-05
**DwgLayerTransfer（DWGレイヤ表示設定の移行）** の作り直しと不具合解消。
左右2ペインの段階選択UIへ変更し、移行元を常にビュー単位に固定。
「移行しても設定が変わらない」不具合を解消（詳細は `Docs/DEVLOG.md`）。
書き込みは「効いたか」を元のビューで検証する方式へ全面的に変更し、
V/G は全項目を移すよう拡張、オブジェクトスタイル（モデル全体）の移行も追加（既定OFF）。
動作確認完了。

## 前回セッション: 2026-07-27
**ExcelExportImport** の一連の改善。エクスポートダイアログに全選択/選択解除、
出力欄のクリア、書き出し済みExcelからの出力設定読込（列入れ替えにも対応・高速化）、
出力欄の▲▼移動のスクロール維持、設定ボタンの集約、シート分割チェックの出力欄への移動、
OKボタンを「エクスポート実行」に改名、Excelインポートのシート情報削除 などを実施。


# 開発ステータス

> このファイルはセッション終了時に更新すること

## 現在作業中
残項目2件の対応 — Revit での動作確認待ち
動作確認対象: **Revit 2022 / 2024**
ブランチ: `claude/parameter-organization-improvement-1eol6q`

### 完了（2026-08-28 残項目2件）
- [x] **シート一括作成にリスト入力モードを追加**（案A）
      - ラジオボタンで「枚数を指定して連番で作る」/「番号と名前のリストから作る」を切替
      - `シート番号[Tab]シート名` を複数行入力。Excel の2列をそのまま貼り付け可能
      - 空行無視／タブ無しは番号のみ／3列目以降無視／入力内の重複は先勝ち
      - 既存シート番号と重複する行はスキップし、**スキップした番号を結果ダイアログに列挙**
      - ダイアログは `SizeToContent="Height"` でモードによる高さ差を吸収
      - ⚠ `Window` 派生での `Visibility` は CS0176 になるため完全修飾（DEVLOG 参照）
- [x] **Excel ヘッダーの多言語化**
      - `ExcelHeaderNames` を新設（要素ID / カテゴリ / 統合シート名）
      - `ParameterHeaderMarker` のマーカー3種も `Loc.S` 化
      - **書き出しは現在の言語 / 読み込みは3言語すべてを候補**にするため、
        別言語で書き出した Excel も、日本語固定だった旧版の Excel も読める
      - 英語 Revit で「タイプ」列が編集不可扱いになる不具合を修正
        （`p.RawName == "タイプ"` → `BuiltInParameter.ELEM_TYPE_PARAM` の Id で判定）

### 旧・現在作業中
GenericModelMerge（一般モデル化）— 新規実装完了。Revit での動作確認待ち
動作確認対象: **Revit 2022 / 2024**（AutoBuild を `[build:2022,2024]` で実行）
ブランチ: `claude/unified-model-generation-fzdtr1`

### 完了（2026-08-28 監査対応 第2弾）
- [x] **`Cast<T>()` → `OfType<T>()`**（5箇所）: 梁天端/梁下端（構造フレーム）、型枠（図枠）、
      部屋3D色分け/部屋タグ（部屋）。カテゴリ指定だけでは型が保証されず、DirectShape 等が
      混ざると遅延評価中に `InvalidCastException` でコマンドごと落ちていた
- [x] **Excel の要素Id 幅を統一**: 書出 long / 読込 `int.TryParse` / 結果マーキング `double`→`int` の
      3通りを `TryParseElementId` + `ToElementId`（`#if REVIT2026`）に集約。
      `ImportPreviewRow.ElementId` も `long` へ。Revit 2026 で大きな Id の行が
      黙ってスキップされる問題を解消
- [x] **型枠シートの名前一致削除に確認ダイアログ**: `FindFormworkSheet(doc, out bool byNameOnly)` を
      追加し、タグで見つからず名前一致だけの場合はトランザクション開始前に確認。
      「いいえ」でシート出力全体をスキップ
- [x] **言語メニュー（JP/US/CN）のツールチップ**: `AddPushButton` の戻り値を `_buttons` に登録し、
      `_buttonTipKeys` にも追加。言語切替で更新されるようになった
- [x] **フィルタ整理／テンプレート整理**: Revit の警告を記録して結果ダイアログで通知
      （パラメータ整理と挙動を統一）
- [x] **診断ログを DiagLog に集約**: 5MB でローテーション（退避3世代）、ON/OFF スイッチを実装。
      Application / SheetCreation / FilledRegionSplitMerge / FireProtection の独自実装を委譲。
      出力先は `C:\temp` のまま、既定 ON。`%AppData%\Tools28\diaglog.setting` に `off` で無効化
- [x] **図面No 空欄でシート番号が "- 1" になる不具合**を修正。旧形式 "- 12" も採番対象として
      認識し続けるため、旧バージョンで作ったシートからも連番を継続できる
- [x] **`Docs/Features/SheetCreation.md` を全面改訂**（C案）。マニュアルは「番号と名前を
      タブ区切りで複数行入力／Excel 貼り付け／重複スキップ」と書いていたが、実装には
      その機能が一切無かった。実装（図枠＋作成枚数＋図面No）の説明に書き直し、
      「この機能でできないこと」節を追加

- [x] **AutoBackup 同期中ダイアログを許可リスト方式へ**（第2段・完了）。運用ログから
      「正常に完了する同期ではダイアログが出ない」と確認できたため、許可リスト
      `AutoDismissDialogIds`（初期は空）に無いダイアログは自動処理せず、Revit に表示させて
      ユーザーが判断する方式に変更。普段の同期は従来どおり止まらない

### 次回以降
- [ ] シート一括作成のリスト入力方式（番号・名前を1枚ずつ指定 / Excel 貼り付け）— 機能追加として検討
- [ ] Excel ヘッダー「要素ID」「カテゴリ」の多言語化 — 読み戻し互換の設計が必要

### 完了（2026-08-28 GenericModelMerge セッション）
- [x] `Commands/GenericModelMerge/` 一式を新設（Command / Models / Services / Views）
- [x] 対象は**アクティブ3Dビューに表示されている要素**。セクションボックス・ビューフィルタ・
      非表示設定がそのまま効く（`FilteredElementCollector(doc, view.Id)`）
- [x] ダイアログにカテゴリチェックリスト（要素数つき・初期は全チェック）
- [x] 出力形式を2択に：**ダイレクトシェイプ** / **一般モデルファミリ(.rfa)**。
      各選択肢に平易な説明文（`*.Hint`）を必ず併記
- [x] まとめ方を3択に：形はそのまま／すべてつなげる／くっついているものだけつなげる
- [x] 接触判定は **Union-Find** で推移的に連結（既存 `UnionByProximity` の単一パスの穴を回避）
- [x] ブーリアン失敗時に**形状を捨てず**単独形状として残し、件数を結果ダイアログで通知
- [x] 材質は単一材質を指定。ファミリは `MATERIAL_ID_PARAM`、DirectShape は `doc.Paint()`
      （20,000面超は適用を見送り通知。形状は常に正確に作成）
- [x] `.rfa` の保存先はダイアログ指定。ファミリテンプレート(.rft)は自動探索＋手動選択フォールバック
- [x] 元要素の非表示はチェックボックスでON/OFF（初期ON）。`CanBeHidden` でふるってから `HideElements`
- [x] リボンに新規パネル **「モデル」** を新設＋アイコン生成（32/16px + features 用）
- [x] 3言語（JP/US/CN）の文字列を55キー同時追加。キー一致を検証済み
- [x] `Docs/Features/GenericModelMerge.md` / `Docs/features.json`（`model` カテゴリ追加、`added_in: 2.6`）
      / `Docs/DEVLOG.md` / `CLAUDE.md` を更新

#### 動作確認1回目で判明した不具合の修正（同日）
- [x] **材質のコピー方法を修正**。`ElementTransformUtils.CopyElements` はプロジェクト↔ファミリ間で
      使えず「ファミリとプロジェクト間ではコピーできません」エラーになる。
      `Material.Create(famDoc, 同名)` でファミリ側に作り直す方式へ変更
- [x] **`Transaction.Commit()` の戻り値を必ず確認**するよう修正。エラーのモーダルを
      ユーザーがキャンセルするとロールバックされ、**空の .rfa が保存されていた**
- [x] `FamilyFailurePreprocessor` を新設。警告は削除、エラーは文言を収集してロールバック
- [x] 保存前にファミリ文書の `FreeFormElement` 数を数え直して検証
- [x] **元要素を非表示にする前に**生成物が立体形状を持つか検証し、なければ全体をロールバック
      （「空の一般モデルを作り、元要素だけ消えたビュー」を残さない）

### 完了（2026-08-28 ParameterCleanup セッション）
- [x] `Services/ParameterUsageIndex.cs` 新設。「どの要素・タイプがどのパラメータを保持しているか」を
      ドキュメント1回走査で索引化（インスタンス＝タイプ単位、タイプ＝ファミリ単位でグループ化し、
      代表要素1つだけパラメータ列挙）
- [x] 値判定を `ParameterBindings` 非依存に変更。バインドなしの共有パラメータも
      `get_Parameter(Guid)` / `get_Parameter(Definition)` / `Parameters` 列挙で値を読み取る
- [x] バインド辞書のキーを**パラメータ名 → `InternalDefinition.Id`** に変更（同名による誤判定を解消）
- [x] `ValueState` に `NotFound`（未使用＝どの要素も保持していない）を追加し、`Empty` と区別
- [x] 「使用箇所」列を追加（保持要素数＋カテゴリ内訳、ツールチップにファミリ名・要素ID・値の例）
- [x] 「バインドなしを全選択」→「未使用を全選択」に変更（値確認前は不可、`NotFound` のみ選択）
- [x] 削除前に「値あり／未確認」が含まれる場合の警告ダイアログを追加
- [x] `WarningSwallower` が Revit の警告文を記録し、削除結果ダイアログで通知するよう変更
- [x] 3言語（JP/US/CN）の文字列を同時追加。`Docs/Features/ParameterCleanup.md` / `Docs/DEVLOG.md` 更新


### 完了（2026-08-05 DwgLayerTransfer セッション）
- [x] UI を左右2ペインの段階選択に変更（① 読み取り元ビュー → ② DWG ／ ③ 反映先ビュー（複数可）→ ④ 反映先DWG（複数可））
- [x] 見出しを移行元＝淡いティール／移行先＝淡いアンバーに色分け
- [x] 移行元を常に**ビュー単位**に固定。テンプレートの V/G で「読み込み」が含まれていなくても、ビューから読めば実際に効いている値が取れるため
- [x] 反映先の単位（ビュー/ビューテンプレート）は移行先にのみ適用
- [x] テンプレート制御下のビューは、書き込み先をテンプレート／主ビューへ自動で振り替え
- [x] **検証を「書けたか」から「元のビューで効いたか」へ変更**。伝播しない書き込みは SubTransaction ごと破棄
- [x] V/G は `OverrideGraphicSettings` の**全項目**を移行（サーフェス/切断パターン・透明度・詳細レベルを追加）
- [x] オブジェクトスタイル（線の色・太さ・線種）の移行を追加。**モデル全体に効くため既定OFF**＋確認ダイアログで警告
- [x] 割り当てテンプレート名を一覧に常時表示。結果ダイアログに実際の書き込み先名を表示
- [x] 線種・塗潰しの名前解決は完全一致を優先（`2_Hidden2` と `2_HIDDEN2` は別要素）
- [x] 診断ログ（`C:\temp\Tools28_debug.txt`）と、移行元 vs 移行先の項目別突き合わせダンプを追加

### 完了（2026-07-24 ParameterCleanup セッション）
- [x] 新機能 ParameterCleanup を追加（`Commands/ParameterCleanup/`）。リボン「パラメータ」パネル＞「パラメータ整理」
  - Command / ParameterScanner / ParamRow / ParameterCleanupDialog（WPF DataGrid）
  - 対象: プロジェクト/共有パラメータ（`ParameterElement`/`SharedParameterElement`）＋グローバルパラメータ
- [x] 大容量モデルのフリーズ回避設計（列挙は軽量／値判定はカテゴリ限定・キャッシュ・early-exit／Stopwatch で約50ms毎に UI へ制御を返し進捗＋中止対応）
- [x] バインド解決を `get_Item` から `ForwardIterator`（名前キー）へ変更（多数が誤って「対象外」になる不具合を修正）
- [x] 値判定の意味を明確化：「空（未使用）」＝バインド済みだが全要素で値なし／「バインドなし」＝カテゴリ未バインド。値セルにツールチップで説明
- [x] 値の有無をダイアログ表示時に自動確認（削除後も自動再確認）。確認後に「値あり◯/空◯」サマリー表示
- [x] 集計表参照列を追加（`ScheduleField.ParameterId` で軽量取得。フィルタ/タグ/数式はAPI制約により対象外）
- [x] 種別ラジオ（すべて/プロジェクト/共有/グローバル）で絞り込み
- [x] Excel風の列メニュー（各見出し「▾」）：昇順/降順並べ替え＋値チェックリストで絞り込み（検索・全選択/選択解除・長い値は…省略＋ツールチップ・幅上限420）
- [x] 一覧下に「全選択」「選択解除」ボタン（フィルターで表示中の行のみ対象に削除用チェックを一括操作）
- [x] 行間の水平罫線を薄いグレー（#ECECEC）に
- [x] 「空のみ」チェック削除（値列フィルターで代替）・「バインドなしを全選択」ボタン追加
- [x] 列フィルターのポップアップ文字サイズをダイアログと統一（配置元ボタンからの継承で小さくなっていた）
- [x] フィルターの クリア/OK/キャンセル ボタンの文字を中央寄せに
- [x] ダイアログ上部の説明文を2行に改行（3言語）
- [x] リボンボタン名を2行（パラメータ/整理）に改行（3言語）
- [x] 削除時に Revit の警告ダイアログが数枚出る問題を修正（`IFailuresPreprocessor`＋`ForcedModalHandling=false` で警告を自動抑制）

### 完了（2026-07-24 Excel連携・カテゴリ名不整合の修正）
- [x] エクスポートダイアログのカテゴリ名がセクション間で食い違う不具合を修正
  - 原因: カテゴリ選択欄のみ `CategoryLocalizer`（固定翻訳。例 OST_Site→「敷地」）を使い、パラメータ/出力欄・Excelシート名の Revit 実名（「外構」）とずれていた
  - 対応: `CategoryInfo.DisplayLabel` を Revit 実名 `Category.Name` に統一し、二重管理の元凶 `CategoryLocalizer` を撤去。全表示が Revit 表記と一致
- [x] 3言語（JP/EN/CN）対応、`Docs/features.json`（added_in 2.2）・マニュアル `Docs/Features/ParameterCleanup.md` 追加
- ⚠️ ビルド時のハマりどころ（DEVLOG参照）: `Window` のインスタンスプロパティと同名の列挙型（`Visibility`/`HorizontalAlignment`/`VerticalAlignment`）は CS0176、`TextBox` は `Autodesk.Revit.UI.TextBox` と衝突し CS0104 → いずれも完全修飾で解決

---

### 旧・現在作業中（参考）: Excel エクスポート／インポートの改善
ブランチ: `claude/excel-export-improvements-r2q928`

### 完了（2026-07-23 セッション）
- [x] エクスポートダイアログのパラメータ欄に `I-`（インスタンス）/`T-`（タイプ）の凡例を追加（背景色なし・2行表示、JP/EN/CN）
- [x] 大容量モデルでパラメータのチェック／ホバーが重い問題を解消（グループ化時のUI仮想化を有効化）
- [x] エクスポートの高速化（タイプ値キャッシュ・`LookupParameter`・列幅のインライン集計・オートフィルタ範囲直接指定）
- [x] 書き出したExcelのヘッダー行（1行目）を固定（`FreezeRows(1)`）
- [x] 「カテゴリ毎にシートを分ける」→「出力Excelをカテゴリ毎にシートに分ける」に文言改善
- [x] インポートの高速化（`LookupParameter`・`ImportFromPreview` で変更セルのみ書込み・タイプ値キャッシュ）
- [x] 開いているExcelの色付けフリーズを解消（`MarkCellsViaCom` を `Range.Value2` 一括読取に・`EnableEvents=false`）
- [x] インポート読込の二重オープン解消（`GeneratePreview` にシート名 out 版を追加）
- [x] 設定ファイル(.json)読込の高速化（`ApplySettings` のパラメータ二重取得を解消・HashSet/Dictionary照合）
- [x] AutoBuild: コミット件名マーカー `[build:XXXX]` で対象Revitバージョンを切替可能に
- [x] AutoBuild: `RestartAutoBuild.ps1`（停止＋再起動を1コマンド）を追加
- [x] インポートで値の削除（空欄化）が反映されない不具合を修正（文字列パラメータの空セルをクリア変更として取込む）
- [x] 削除インポート成功セルを青塗りで表示（変更成功は青字、削除成功は青セル）
- [x] エクスポートのパラメータ欄カテゴリ見出しを折りたたみ可能に（▼/▶でカテゴリ単位に開閉）

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
| ParameterCleanup（パラメータ整理/未使用削除） | ✅ 完了・動作確認済み | v2.2予定・Excel風フィルター/自動値判定 |
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

<!-- rebuild trigger: Excelインポート単位変換・制約エラー対応の再ビルド -->

<!-- rebuild trigger: 書き出し時マーカー/失敗メッセージ改善の再デプロイ -->

<!-- rebuild trigger: 文字設定不可列のグレー化 再デプロイ -->

<!-- rebuild trigger: ヘッダーマーカー8pt化 再デプロイ -->

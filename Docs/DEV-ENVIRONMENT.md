# 開発環境の説明（他リポジトリへ移植するための仕様書）

このリポジトリ（Revit-Add-ins / Tools28）で成立している
**「ユーザーは指示を出して、CAD 上で動作確認するだけ」** という開発体制を、
他リポジトリ（例: ARES Commander 用アドイン）で再現するための説明書。

- 対象読者: 新しいリポジトリで開発を始めるときの Claude と、ユーザー本人
- 前提: 個人開発（レビュー者がいない／main への直マージを許容できる）

---

## 1. Claude が動いている場所と、できないこと

| 項目 | 実際 |
|--|--|
| 実行環境 | claude.ai/code のリモートコンテナ（**Linux**）。ユーザーの PC ではない |
| リポジトリ | セッション開始時にクローンされる。コンテナは使い捨て（push しないと消える） |
| できること | ソースの閲覧・編集、git commit / push、GitHub API、Notion 等の MCP 操作 |
| **できないこと** | **MSBuild でのビルド／CAD の起動／実行時の動作確認／画面の目視確認** |

つまり Claude は **コードが「コンパイルできるか」すら自分では確認できない**。
検証はすべて、ユーザーの Windows PC 上の実機（Revit / ARES）で行われる。
この非対称性を埋めるために、以下 2〜3 の自動化がある。

---

## 2. 1 サイクルの流れ（誰が何をするか）

```
🙂 ユーザー: チャットで指示
        ↓
🤖 Claude: claude/xxx ブランチにコミット → push
        ↓ (GitHub Actions: auto-merge-claude.yml)
⚙ main へ squash merge（PR なし・レビューなし）・claude/* ブランチ削除
        ↓ (30 秒ポーリング)
⚙ 開発 PC 常駐の AutoBuild.ps1 が origin/main の変更を検知
   → git reset --hard origin/main
   → QuickBuild.ps1（MSBuild でビルド → アドイン配置先へコピー）
   → 成功／失敗を MessageBox で通知
        ↓
🙂 ユーザー: 通知を見て CAD を再起動 → 触って確認
        ↓
🙂 ユーザー: 結果（OK／不具合／ログ）をチャットで報告 → 次サイクルへ
```

**ユーザーの手作業は 4 つだけ**
1. 指示を出す
2. 通知（ビルド成功／失敗）を待つ
3. CAD を再起動して動作確認する
4. 結果を伝える（不具合ならログ・スクリーンショットを貼る）

ビルド・デプロイ・マージ・ブランチ削除・リリースは、すべて自動。

---

## 3. 構成要素の実体

### 3-1. リポジトリ側（Claude が触る）

| ファイル | 役割 |
|--|--|
| `CLAUDE.md` | Claude への常時指示。**唯一のルールブック**。コミットメッセージ規約、多言語化ルール、新機能追加手順、リリース手順をここに集約 |
| `STATUS.md` | 現在の作業状況。Stop hook で自動追記される |
| `TASKS.md` | TODO と完了ログ |
| `Docs/DEVLOG.md` | 機能ごとの設計・不具合の知見（CLAUDE.md から参照して肥大化を防ぐ） |
| `.claude/settings.json` | hook 定義（下記） |
| `dev-config.json` | 開発時の既定ビルド対象（`"defaultRevitVersion": "2022"`） |

`.claude/settings.json` の hook は 2 つ:

- **PreToolUse (Bash)**: コマンドに `git push` が含まれていたら、実行前に
  `git fetch origin main` → `stash` → `rebase origin/main` → `stash pop` を自動実行。
  main が先に進んでいても push が弾かれない。
  ⚠ この仕組みのため **commit と push は必ず別々の Bash 呼び出しで実行する**
  （同じ呼び出しに混ぜると、未コミット状態で rebase されて事故る）。
- **Stop**: セッション終了時に `.claude/update-status.sh` が変更ファイル一覧を `STATUS.md` 冒頭へ追記。

### 3-2. GitHub Actions 側

| ワークフロー | 契機 | 内容 |
|--|--|--|
| `auto-merge-claude.yml` | `claude/**` への push | main へ **squash merge**（PR を作らない）→ 開いている claude PR を close → `claude/*` ブランチを全削除 → Pages デプロイを起動 |
| `build-and-release.yml` | タグ `v*` / `release/v*` / 手動 | 全バージョンをビルド → 配布 ZIP → GitHub Releases（本文は `Docs/features.json` から自動生成） |
| `deploy-pages.yml` | main | `Docs/` を GitHub Pages へ |
| `notify-site.yml` | `Docs/Features/**` の変更 | 配布サイトのリポジトリへ `repository_dispatch` |
| `local-build-deploy.yml` | main への `.cs/.xaml/.csproj` 変更 | self-hosted runner での予備ビルド経路（現在の主経路は AutoBuild.ps1） |

⚠ squash merge は **コミット件名（1 行目）しか main に残らない**。
後述の `[build:XXXX]` マーカーを本文に書くと消えるのはこのため。

### 3-3. 開発 PC 側（ユーザーの Windows。Claude からは触れない）

| ファイル | 役割 |
|--|--|
| `StartAutoBuild.vbs` | ダブルクリック → UAC で昇格 → `AutoBuild.ps1` を非表示で常駐起動 |
| `AutoBuild.ps1` | 30 秒間隔で `origin/main` をポーリング。変更検知で `reset --hard` → ビルド → 通知。多重起動は Mutex で防止 |
| `RestartAutoBuild.ps1` | 停止＋再起動を 1 コマンドで。⚠ `AutoBuild.ps1` 自体を変更したら再起動しないと反映されない |
| `QuickBuild.ps1` | 単一バージョンのビルド＆デプロイ。vswhere で MSBuild を検索 → `bin\Release\Revit{ver}\` → `%ProgramData%\Autodesk\Revit\Addins\{ver}\28Tools\` へ DLL 一式、ルートへ `.addin` をコピー |
| `BuildAll.ps1` / `CreatePackages.ps1` | 全バージョンビルドと配布 ZIP 作成（リリース時） |
| `AutoBuild.log` / `AutoBuild_detail.log` | 監視ログ / ビルド詳細ログ。**失敗時にユーザーが貼る情報源** |

ビルド成功判定は終了コードではなく **DLL のタイムスタンプが更新されたか** で行っている
（`& .\script.ps1` 経由の `$LASTEXITCODE` が信用できないため）。

デプロイ時、CAD が起動中だと DLL がロックされてコピーに失敗する。
本体 DLL の失敗は致命的エラー、依存 DLL はスキップして続行する作りになっている。

### 3-4. ビルド対象バージョンの切り替え

- 既定は `dev-config.json` の `defaultRevitVersion`（通常 2022）。
- 一時的に変えたいときは **コミット件名の末尾**にマーカーを付ける:
  `[build:2024]` / `[build:2024,2025]` / `[build:all]`
- そのコミットのビルドだけ対象が変わり、次回からは既定に戻る。
- ⚠ 必ず**件名（1 行目）**に書く。squash で本文が落ちるため。
  複数コミットを積む場合は **HEAD になる最後のコミットの件名**に入れる。

---

## 4. 失敗したときの情報経路

Claude は実機を見られないので、次の 4 つをユーザーが貼ることで初めて調査できる。

| 症状 | ユーザーが貼るもの |
|--|--|
| ビルドが失敗（通知が「ビルド失敗」） | `AutoBuild_detail.log` の末尾（コンパイルエラー行） |
| ビルドは通るが CAD に出てこない | `AutoBuild.log`、アドイン配置先のファイル一覧 |
| 実行時に落ちる／挙動がおかしい | アプリのデバッグログ（本リポジトリは `C:\temp\Tools28_debug.txt`）、エラーダイアログのスクリーンショット |
| UI が思っていたものと違う | スクリーンショット |

**「実行時ログをファイルに吐く」ことは、この体制の必須要件**。
Claude が実機を見られない代わりの唯一の目になる。

---

## 5. この方式が成立している条件（＝移植時に満たすべき前提）

1. 開発 PC が起動していて、監視スクリプトが常駐している
2. main が常に「ビルドしてよい状態」＝レビューなしの直マージを許容できる（個人開発だから成立）
3. **ビルド〜配置がコマンド 1 本**にまとまっている
4. 配置先が CAD が自動でアドインを読むパスである（手動インストール不要）
5. 成否が必ずユーザーに届く（デスクトップ通知）
6. アプリが実行時ログをファイルに残す

### 既知の弱点（承知の上で使っている）

- Claude はコンパイル検証ができないため、API シグネチャ誤りは**実機ビルドで初めて判明**する
- main へ直マージなので、壊れたコミットもそのまま main に入る（次のコミットで直す運用）
- `AutoBuild.ps1` 自体の変更は再起動するまで反映されない
- CAD 起動中のデプロイは DLL ロックで一部スキップされる（再起動後に再実行で解消）

---

## 6. ARES 版で「決めなければならないこと」

以下は CAD 依存なので、**Claude が推測で決めてはいけない**。
新リポジトリの最初のセッションでユーザーに確認して `CLAUDE.md` に固定する。

| 項目 | このリポジトリ（Revit） | ARES 側で決めること |
|--|--|--|
| API の入手方法 | NuGet `Nice3point.Revit.Api.*` | SDK の DLL 参照か、NuGet か、参照 DLL の置き場所 |
| ターゲットフレームワーク | net48（2021-2024）/ net8.0-windows（2025-2026） | ARES の対応ランタイム |
| バージョン軸 | Revit 2021〜2026 のマルチターゲット（条件付きコンパイル） | ARES はバージョン別ビルドが必要か。不要なら単一構成に簡素化 |
| ビルド | `msbuild Tools28.csproj /p:RevitVersion=XXXX` | msbuild か dotnet build か、プロジェクト形式 |
| 配置先 | `%ProgramData%\Autodesk\Revit\Addins\{ver}\28Tools\` + `.addin` マニフェスト | ARES のアドオン読み込みパス（フォルダ／レジストリ／マニフェスト） |
| 起動時ロード | `.addin` XML | ARES のロード方法（起動時自動 or コマンドでロード） |
| 実行時ログ | `C:\temp\Tools28_debug.txt` | ログ出力先のパスを決めて固定する |
| 多言語化 | JP/US/CN の 3 言語必須（`Loc.S()`） | 必要か。必要なら同じ 3 分類ルールを移植する |
| 配布 | GitHub Releases + ZIP + install.bat | 配布形態（インストーラ／ZIP／不要） |

その他、このリポジトリから**そのまま持っていける汎用ルール**:

- コミットメッセージは日本語、1 行目に要約
- 文字列の 3 分類（A: UI 表示のみ多言語化 / B: モデルに残るが検索キーでない / C: 検索キーは固定＝多言語化禁止）
- `CLAUDE.md` / `STATUS.md` / `TASKS.md` / `DEVLOG.md` の役割分担
- commit と push は別々の Bash 呼び出し

---

## 7. 新リポジトリ立ち上げの手順（ARES 側で行う作業）

1. **リポジトリ作成** し、`CLAUDE.md` を書く（第 6 節の表を埋めた内容 + 汎用ルール）
2. **`.github/workflows/auto-merge-claude.yml` をコピー**（変更不要。`claude/**` push で main へ squash merge）
3. **`QuickBuild.ps1` 相当を作る**: ビルド → アドイン配置先へコピー、まで 1 本で完結させる
4. **`AutoBuild.ps1` / `StartAutoBuild.vbs` / `RestartAutoBuild.ps1` をコピー**し、
   プロジェクト名・DLL 名・配置先・`[build:]` マーカーの扱いを ARES 用に差し替える
   （バージョン軸が不要なら `Get-BuildVersions` は削って単純化してよい）
5. **`dev-config.json`** を用意（不要なら省略）
6. **`.claude/settings.json`** をコピー（push 前 rebase hook。`update-status.sh` のパスを新リポジトリに合わせる）
7. **疎通テスト**: Claude に「起動確認用の最小コマンドを 1 個追加して」と指示 →
   通知が出て CAD のメニューに現れるところまで確認する。
   **ここが通れば、以降は指示と動作確認だけで回る**

---

## 8. 新セッションに貼る指示文（テンプレート）

新しいリポジトリのセッション冒頭で、そのまま貼れる文面。

```
このリポジトリでは、Revit-Add-ins（Tools28）と同じ開発体制を再現したい。
体制の仕様は Revit-Add-ins の Docs/DEV-ENVIRONMENT.md にまとめてある（内容は以下に要約）。

【体制】
- あなた（Claude）は claude.ai/code のリモート Linux コンテナで動く。
  ビルドも CAD の起動も動作確認もできない。コードを書いて push するところまでが担当。
- あなたが claude/xxx ブランチに push すると、GitHub Actions が main へ squash merge する。
- 私の Windows PC に常駐している監視スクリプトが main の変更を検知し、
  自動でビルドして ARES のアドイン配置先へコピーし、成否をデスクトップ通知で知らせる。
- 私は通知を見て ARES を再起動し、動作確認して結果をチャットで返す。
- したがって「ビルドが通るか」は私が実機で確認するまで分からない。
  コンパイルエラーは私が AutoBuild_detail.log を貼るので、それを見て直してほしい。

【あなたへのルール】
- 作業ブランチは claude/<内容>-<任意> を使う。main へ直接 push しない。
- commit と push は必ず別々の Bash 呼び出しで実行する（push 前 rebase hook があるため）。
- コミットメッセージは日本語。1 行目に要約（50 文字目安）。
- ビルド対象を一時的に変えたいときは、コミット件名の末尾に [build:XXXX] を付ける
  （squash で本文は消えるので必ず件名に書く）。
- 実行時ログはファイルに出す設計にする（私が貼れる唯一の手がかりになる）。
- CLAUDE.md / STATUS.md / TASKS.md / DEVLOG.md の役割分担を守る。

【最初にやること】
まず ARES 依存の以下を私に質問して、CLAUDE.md に固定してほしい。推測で決めないこと。
1. ARES の API 参照方法（SDK の DLL / NuGet / 置き場所）
2. ターゲットフレームワークと、バージョン別ビルドの要否
3. ビルドコマンド（msbuild / dotnet build）
4. アドインの配置先パスと、ARES への読み込ませ方（マニフェスト／レジストリ等）
5. 実行時ログの出力先
6. 多言語化の要否
7. 配布形態（ZIP / インストーラ / 不要）

そのうえで、
- CLAUDE.md
- .github/workflows/auto-merge-claude.yml
- ビルド＆デプロイスクリプト（QuickBuild 相当）
- 常駐監視スクリプト（AutoBuild 相当）と起動用 VBS
- .claude/settings.json（push 前 rebase hook）
を用意し、最後に「起動確認用の最小コマンドを 1 個追加」して疎通テストまで持っていってほしい。
```

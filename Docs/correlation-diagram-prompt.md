# 相関図 画像生成AIプロンプト

> 以下のプロンプトを画像生成AI（ChatGPT / DALL-E、Midjourney、Adobe Firefly など）に貼り付けて使ってください。
> 日本語テキストは画像生成AIでは文字化けしやすいため、**テキストなし版**と**テキストあり版**の2パターンを用意しました。

### 登場人物（6人）

| 登場人物 | アイコン | 役割 |
|---|---|---|
| あなた（開発者） | 人物アイコン | 監督・プロデューサー |
| Claude Code | Anthropic の Claude ロゴ（オレンジ〜茶色の丸いアイコン） | AI脚本家（コードを全て書く） |
| PowerShell | PowerShell ロゴ（青い背景に `>_` マーク） | 現場スタッフ（ビルド&デプロイ） |
| Revit | Autodesk Revit ロゴ（青い `R` マーク） | 舞台（動作確認の場所） |
| GitHub | GitHub ロゴ（黒い Octocat） | 台本倉庫 + 配達係 |
| NuGet | NuGet ロゴ（青い立方体） | 小道具係（API部品の自動調達） |

※ Visual Studio は不要です（コードは全て Claude Code が書きます）

---

## プロンプト A: テキストなし版（推奨）

画像生成AIで日本語テキストを正確に描くのは難しいため、**アイコンと矢印だけの図を生成**し、あとから PowerPoint や Canva などでテキストを載せる方法です。

```
Create a clean, professional infographic-style flowchart diagram for a software development workflow.
White background, modern flat design, 16:9 aspect ratio.

Layout (top to bottom, 3 phases):

【Phase 1 - DEVELOP (top section)】
- Top center: a PERSON icon (developer) in a large BLUE-bordered rounded rectangle
- One arrow going straight down from the person to:
- Center: the Anthropic Claude logo (a rounded orange-brown icon) in a PURPLE-bordered rounded rectangle — this represents "Claude Code", the AI that writes all the code
- One arrow going straight down from Claude to:
- Center: a DOCUMENT / CODE file icon in a light gray rounded rectangle — this represents the generated code

【Phase 2 - BUILD & TEST (middle section)】
- The PowerShell logo (blue background with ">_" terminal symbol) in a BLUE-bordered rounded rectangle, centered
- Two arrows from PowerShell splitting left and right:
  - Left arrow to: a GEAR/BUILD icon (compile) in a GRAY-bordered rectangle
  - Right arrow to: the Autodesk Revit logo (blue "R" icon) in a BLUE-bordered rounded rectangle
- A small NuGet logo (blue cube icon) in a GRAY-bordered rectangle, connected to the gear/build icon with a DOTTED arrow (auto-supply)
- A curved feedback arrow from Revit back up to the Person icon at the top, with a GREEN checkmark (OK) and RED X mark (NG)

【Phase 3 - RELEASE (bottom section)】
- The GitHub logo (black Octocat) in a DARK-bordered rounded rectangle
- Arrow from GitHub to: a ROCKET / automation icon (GitHub Actions) in a GRAY-bordered rectangle
- Arrow from the rocket to: a DOWNLOAD / package icon
- Arrow from the download icon to: multiple USER silhouettes

Each software element should display its ACTUAL RECOGNIZABLE LOGO/ICON inside the rounded rectangle.
No text or labels anywhere in the image — only logos, icons, and arrows.

Color coding for borders:
- BLUE border: things the developer does manually
- PURPLE border: AI-powered (Claude Code)
- GRAY border: runs automatically

Style: minimal, modern, flat design, slight drop shadows, pastel color palette, clean curved arrows
```

> **生成後**: PowerPoint や Canva で以下のテキストを各アイコンの横に配置してください
> - 人アイコン → 「あなた（監督）」
> - Claude ロゴ → 「Claude Code（AI脚本家）」
> - PowerShell ロゴ → 「PowerShell（現場スタッフ）」
> - Revit ロゴ → 「Revit（舞台）」
> - NuGet ロゴ → 「NuGet（小道具係）」
> - GitHub ロゴ → 「GitHub（台本倉庫）」
> - ロケットアイコン → 「GitHub Actions（自動スタッフ）」
> - 矢印の横に：「指示を出す」「コードを全て書く」「ビルド実行」「動作確認」「OK/NG」など

---

## プロンプト B: テキストあり版（英語ラベル）

画像生成AIに英語テキストを入れた状態で生成し、あとから日本語に差し替える方法です。

```
Create a professional software development workflow diagram as an infographic.
White background, modern flat design, 16:9 aspect ratio.

The diagram flows top-to-bottom with 3 phases separated by thin horizontal divider lines.
Each element is a soft rounded rectangle containing the software's ACTUAL LOGO and a short English label.

=== PHASE 1: DEVELOP ===
Top center:
  [Person icon] "YOU (Director)" — blue border, large size

One arrow going down, labeled "Give instructions":
  [Anthropic Claude logo - orange/brown rounded icon] "Claude Code (AI)" — purple border

Arrow going down from Claude Code, labeled "Writes ALL code":
  [Document/code icon] "CODE" — centered, wide rectangle, light gray

=== PHASE 2: BUILD & TEST ===
Arrow from CODE down to:
  [PowerShell logo - blue ">_" icon] "PowerShell" — blue border, labeled "Run build"

PowerShell splits into two arrows:
  Left: "Build" → [Gear icon] "Build (compile)" — gray border
  Right: "Deploy" → [Autodesk Revit logo - blue "R"] "Revit" — blue border, labeled "Test here"

Small dotted arrow into Build from:
  [NuGet logo - blue cube] "NuGet" — gray border, labeled "Auto-supply"

Curved arrow from Revit back up to YOU:
  Labeled "OK → next" (green) / "NG → tell Claude to fix" (red)

=== PHASE 3: RELEASE ===
Arrow from YOU labeled "Push tag" to:
  [GitHub Octocat logo] "GitHub" — dark border

Arrow from GitHub labeled "Triggers" to:
  [Rocket icon] "GitHub Actions" — gray border, labeled "Auto build + ZIP"

Arrow from GitHub Actions to:
  [Download icon] "Release" — gray border

Arrow from Release to:
  [Users icon] "Users download"

=== LEGEND (bottom-right corner) ===
Blue border = You do this manually
Purple border = AI does this
Gray border = Runs automatically

IMPORTANT: Use the ACTUAL recognizable logo for each software tool (Claude, PowerShell, Revit, GitHub, NuGet).
Style: clean, minimal, pastel colors, thin arrows with labels, slight shadows
```

---

## プロンプト C: ドラマ相関図風（カジュアル版）

TV番組の人物相関図のようなポップなデザインにしたい場合のプロンプトです。

```
Create a Japanese TV drama-style character relationship diagram (相関図 / soukanzu) for a software development team.

Style: colorful, pop, TV drama relationship chart style. Each character is in a circular frame connected by labeled arrows. Bright pastel background with decorative elements. Fun and approachable, like a variety show graphic.

Characters (each in a circular frame with their ACTUAL SOFTWARE LOGO prominently displayed):

1. "YOU" (top center, largest circle, gold border)
   - Icon: a smiling person at a desk with a lightbulb above their head
   - Role badge below: "Director & Producer"

2. "Claude Code" (directly below YOU, purple border)
   - Icon: the Anthropic Claude logo (orange-brown rounded icon) prominently displayed
   - Role badge below: "AI Scriptwriter — writes ALL code"

3. "PowerShell" (center-left, teal/blue border)
   - Icon: the PowerShell logo (blue background with ">_" terminal symbol)
   - Role badge below: "Build Staff"

4. "Revit" (center-right, blue border)
   - Icon: the Autodesk Revit logo (blue "R" mark)
   - Role badge below: "The Stage"

5. "GitHub" (bottom center, dark gray border)
   - Icon: the GitHub Octocat logo
   - Role badge below: "Script Vault & Delivery"

6. "NuGet" (small circle, near PowerShell, light gray border)
   - Icon: the NuGet logo (blue cube)
   - Role badge below: "Props Supplier"

Arrows between characters (with relationship labels on the arrows):
- YOU → Claude Code: "Gives instructions" (thick arrow going down)
- Claude Code → CODE (small document icon between Claude and PowerShell): "Writes all code" (arrow)
- YOU → PowerShell: "Runs build script" (arrow)
- PowerShell → Revit: "Auto deploy" (arrow)
- NuGet → PowerShell: "Auto-supply API parts" (dotted arrow)
- Revit → YOU: "OK! / NG → fix it" (curved feedback arrow, green & red)
- YOU → GitHub: "Push code & tags" (arrow)
- GitHub → Users (small user silhouettes): "Auto release → Download" (arrow)

IMPORTANT: Each circle must display the ACTUAL RECOGNIZABLE LOGO of the software, not a generic icon.

Legend in bottom-right corner:
  Blue = Manual (you do it)
  Purple = AI-powered
  Gray = Fully automatic

16:9 aspect ratio, high resolution, vibrant colors
```

---

## 使い分けガイド

| プロンプト | おすすめのAI | 特徴 |
|---|---|---|
| **A: テキストなし版** | DALL-E, Midjourney, Firefly | 最も確実。テキストは後から載せる |
| **B: テキストあり英語版** | ChatGPT (GPT-4o), DALL-E 3 | 英語ラベル付きで生成。日本語は後から差替え |
| **C: ドラマ相関図風** | ChatGPT (GPT-4o), Midjourney | ポップで楽しい見た目。プレゼン映えする |

### おすすめの作成手順

1. 上記プロンプトで画像生成AI に図のベースを作ってもらう
2. 生成された画像を **PowerPoint** や **Canva** に貼り付ける
3. 日本語のテキストラベルと矢印の説明を手動で上に配置する
4. 色の凡例（青=手動、紫=AI、灰=自動）を追加する

> **注意**: 画像生成AIはロゴを正確に再現できないことがあります。その場合は、生成された図をベースにして、PowerPoint/Canva で各ソフトの公式ロゴ画像に差し替えてください。

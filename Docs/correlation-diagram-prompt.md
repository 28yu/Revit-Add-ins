# 相関図 画像生成AIプロンプト

> 以下のプロンプトを画像生成AI（ChatGPT / DALL-E、Midjourney、Adobe Firefly など）に貼り付けて使ってください。
> 日本語テキストは画像生成AIでは文字化けしやすいため、**テキストなし版**と**テキストあり版**の2パターンを用意しました。

---

## プロンプト A: テキストなし版（推奨）

画像生成AIで日本語テキストを正確に描くのは難しいため、**アイコンと矢印だけの図を生成**し、あとから PowerPoint や Canva などでテキストを載せる方法です。

```
Create a clean, professional infographic-style flowchart diagram for a software development workflow.
The diagram should have a white background with a modern flat design style.

Layout (top to bottom, 3 phases):

【Phase 1 - Top section】
- A person icon (developer) at the top center, highlighted with a BLUE circle
- Two arrows going down-left and down-right from the person
- Down-left arrow points to: a robot/AI icon (Claude Code) with a PURPLE circle
- Down-right arrow points to: a monitor/IDE icon (Visual Studio) with a BLUE circle
- Both the robot and monitor have arrows pointing down to a central document/code icon

【Phase 2 - Middle section】
- A gear/terminal icon (PowerShell) with a BLUE circle, centered
- Two arrows from the gear: one going down-left to a package/build icon, one going down-right to a building/architecture icon (Revit) with a BLUE circle
- A small box icon (NuGet) with a GRAY circle, connected to the package icon with a dotted arrow
- A circular arrow from Revit back up to the person icon (feedback loop), with a GREEN checkmark and RED X mark

【Phase 3 - Bottom section】
- A cloud/repository icon (GitHub) with a DARK circle
- Arrow from GitHub to a rocket/automation icon (GitHub Actions) with a GRAY circle
- Arrow from the rocket to a download/release icon
- Arrow from the download icon to multiple user silhouettes

Color coding for borders/backgrounds:
- BLUE: things the developer does manually
- PURPLE: AI-powered (Claude Code)
- GRAY: runs automatically
- Use soft rounded rectangles for each item
- Arrows should be clean with slight curves
- No text or labels anywhere in the image

Style: minimal, modern, flat design, slight drop shadows, pastel color palette, 16:9 aspect ratio
```

> **生成後**: PowerPoint や Canva で以下のテキストを各アイコンの横に配置してください
> - 人アイコン → 「あなた（監督）」
> - ロボットアイコン → 「Claude Code（AI脚本家）」
> - モニターアイコン → 「Visual Studio（もう一人の脚本家）」
> - 歯車アイコン → 「PowerShell（現場スタッフ）」
> - 建物アイコン → 「Revit（舞台）」
> - 箱アイコン → 「NuGet（小道具係）」
> - クラウドアイコン → 「GitHub（台本倉庫）」
> - ロケットアイコン → 「GitHub Actions（自動スタッフ）」
> - 各矢印の横に動作説明（「指示を出す」「コードを書く」など）

---

## プロンプト B: テキストあり版（英語ラベル）

画像生成AIに英語テキストを入れた状態で生成し、あとから日本語に差し替える方法です。

```
Create a professional software development workflow diagram as an infographic.
White background, modern flat design, 16:9 aspect ratio.

The diagram flows top-to-bottom with 3 phases separated by thin horizontal lines.
Each element is a soft rounded rectangle with an icon and a short English label inside.

=== PHASE 1: DEVELOP ===
Top center:
  [Person icon] "YOU (Director)" — blue border, large size

Two arrows from YOU going down:
  Left arrow labeled "Give instructions" →
    [Robot icon] "Claude Code (AI)" — purple border
  Right arrow labeled "Write yourself" →
    [Monitor icon] "Visual Studio" — blue border

Both Claude Code and Visual Studio have arrows labeled "Write code" pointing down to:
  [Document icon] "CODE" — centered, wide rectangle, light gray

=== PHASE 2: BUILD & TEST ===
Arrow from CODE down to:
  [Gear icon] "PowerShell" — blue border, labeled "Run build"

PowerShell splits into two arrows:
  Left: "Build" → [Package icon] "Build (compile)" — gray border
  Right: "Deploy" → [Building icon] "Revit" — blue border, labeled "Test here"

Small dotted arrow into Build from:
  [Box icon] "NuGet" — gray border, labeled "Auto-supply API parts"

Curved arrow from Revit back up to YOU:
  Labeled "OK → next" (green) and "NG → fix" (red)

=== PHASE 3: RELEASE ===
Arrow from YOU labeled "Push tag" to:
  [Cloud icon] "GitHub" — dark border

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

Style: clean, minimal, flat icons, pastel colors, thin arrows with labels, slight shadows
```

---

## プロンプト C: ドラマ相関図風（カジュアル版）

TV番組の人物相関図のようなポップなデザインにしたい場合のプロンプトです。

```
Create a Japanese TV drama-style character relationship diagram (相関図 / soukanzu) for a software development team.

Style: colorful, pop, TV drama relationship chart style with character portraits in circles connected by labeled arrows. Bright pastel background with decorative elements. Fun and approachable, like a variety show graphic.

Characters (each in a circular frame with a portrait):

1. "YOU" (top center, largest circle, gold border)
   - Cartoon portrait: a smiling person at a desk with a lightbulb above their head
   - Role badge: "Director & Producer"

2. "Claude Code" (left of YOU, purple border)
   - Cartoon portrait: a friendly robot holding a pen and paper
   - Role badge: "AI Scriptwriter"

3. "Visual Studio" (right of YOU, blue border)
   - Cartoon portrait: a computer monitor with code on screen
   - Role badge: "Manual Scriptwriter"

4. "PowerShell" (center, teal border)
   - Cartoon portrait: a worker in a hard hat with a wrench
   - Role badge: "Build Staff"

5. "Revit" (right of PowerShell, orange border)
   - Cartoon portrait: a grand theater stage
   - Role badge: "The Stage"

6. "GitHub" (bottom left, dark gray border)
   - Cartoon portrait: the GitHub octocat holding a filing cabinet
   - Role badge: "Script Vault"

7. "NuGet" (small, near PowerShell, light gray border)
   - Cartoon portrait: a small delivery person carrying boxes
   - Role badge: "Props Supplier"

Arrows between characters (with relationship labels):
- YOU → Claude Code: "Gives instructions" (thick arrow)
- YOU → Visual Studio: "Writes & debugs" (thick arrow)
- Claude Code → CODE (center): "Writes code automatically"
- Visual Studio → CODE: "Writes & fixes code"
- YOU → PowerShell: "Runs build script"
- PowerShell → Revit: "Deploys automatically"
- NuGet → PowerShell: "Supplies API parts" (dotted arrow)
- Revit → YOU: "Test OK / NG feedback" (curved arrow)
- YOU → GitHub: "Pushes code & tags"
- GitHub → "Users": "Auto release & download"

Legend in corner:
  Blue labels = Manual (you do it)
  Purple labels = AI-powered
  Gray labels = Fully automatic

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

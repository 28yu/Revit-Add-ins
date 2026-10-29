# v2.8 リリース本文に書き足す文章

自動生成の本文（`scripts/generate-release-body.py`）は**新機能しか載せない**ため、今回の改善は載らない。
下書きリリースを「Publish release」する前に、GitHub のリリース編集画面で、本文の**先頭**（「Revit 2021 / … 対応」の行の下）に
以下を貼り付ける（28tools-download 側の下書きにも同じ文章を貼る）。

---

### 改善

**ビューの「〜合わせ」（ビューポート位置・視点・切断ボックス・トリミング領域）**
- コピー／ペーストの2つのボタンを、1つのボタンにまとめました（AutoCAD の MA と同じ操作: 元を選ぶ → 合わせたい先を選ぶ）
- ボタン名を「ビューポート位置合わせ」「視点合わせ」「切断ボックス合わせ」「トリミング領域合わせ」に変更しました
- コピー中はカーソルの横に印が出ます。Esc で取りやめます
- 視点合わせは Ctrl+Z で元に戻せます
- ビューポート位置合わせは、枠の中心ではなく図の位置でそろえます（Revit 2023 以降）

**一般モデル化**
- 材質を「元の材質のまま」にできるようにしました（既定）
- 切断ボックスの外側を切り取ります
- 作ったのに見えないときは、理由を結果に表示します。作ったものは選択状態になります
- ファミリで作るとき「書き出して配置／書き出すだけ／配置するだけ」を選べます

**梁レベル色分け（天端／下端）**
- その階の梁だけを色分けします（参照レベルで判定）。「すべて／水平の梁だけ／斜めの梁だけ」を選べます
- レベルの表記を自由に入力できるようにしました（既定はビューのレベル名）
- 同じレベル差は、どの階でも同じ色になります（一度決まった色は変わりません）
- 凡例を階ごとに、凡例ビューとして作成します。実行後は作成した凡例を開きます
- 梁ラベルを「文字」か「梁タグ」から選べるようにしました
- 鉛直ブレースには梁ラベルを付けません。水平ブレースはチェックボックスで選べます

**リアルタイムレンダリング**
- 「表示」の設定を追加: 太陽の光・照明器具、時間帯（朝〜夕方）、空・雲と背景色、地面・地平線と地面の色、単色表示、PNG の背景を透明に
- 3D ビューの「レンダリング設定」（照明・背景・雲の量）を最初の見た目として読み込みます
- フェーズの「上書き」表示の要素を、フェーズのマテリアルで描きます
- 大きなモデルで描画ウィンドウが閉じてしまう問題の対策

### 不具合の修正
- 梁レベル色分け・耐火被覆色分け: 凡例を上書きすると凡例のビューが消えることがある問題を修正

---

## English (optional)

### Improvements
**View "Match" tools (Viewport Position, View, Section Box, Crop Region)**
- Copy and Paste are now a single button (works like AutoCAD's MA: pick the source, then the targets)
- Renamed to Align Viewports / Match View / Match Section Box / Match Crop Region
- A marker next to the cursor shows what is being copied; press Esc to cancel
- Match View can be undone with Ctrl+Z
- Align Viewports aligns by the drawing position instead of the frame center (Revit 2023 and later)

**Generic Model Merge**
- New default material option: keep the original materials
- Geometry outside the section box is trimmed
- If the result is not visible, the reason is shown; created elements are selected
- When creating a family: export and place / export only / place only

**Beam Level Color Code (Top/Bottom)**
- Only beams on that level are colored (by reference level); choose all / horizontal only / sloped only
- The level text can now be entered freely (defaults to the view's level name)
- The same level difference now gets the same color on every floor
- Legends are created per floor as legend views and opened after the run
- Beam labels can be placed as text or as beam tags
- Vertical braces are no longer labeled; horizontal braces are optional

**Real-time Rendering**
- New display settings: sunlight / light fixtures, time of day, sky and background color, ground and ground color, single color, transparent PNG background
- The 3D view's rendering settings (lighting, background, clouds) are used as the starting look
- Elements shown as "Overridden" by the phase filter are drawn with the phase materials
- Mitigation for the render window closing on large models

### Bug fixes
- Beam Level Color Code / Fire Protection Color Code: overwriting a legend could delete the legend view

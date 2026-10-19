# Tools28Render — レイトレーシングの描画プログラム

Revit の「リアルタイムレンダリング」ボタン（`Commands/RealtimeRender/`）から起動される、
別プログラム（exe）の描画エンジンです。C++ と Vulkan で書かれています。

- GPU の**光の計算専用回路**（RT コア、`VK_KHR_ray_query`）で、影・反射・間接光を計算する
- RT コアの無い GPU（Intel の内蔵 GPU、古い GTX など）では、**CPU で作った BVH をシェーダーで自分でたどる予備の方式**に
  自動で切り替える（同じ絵になるが遅いので、半分の大きさで描いて拡大する）
- **ノイズ除去**: 前のフレームの結果を今のカメラ位置に合わせて重ね（再投影）、面の境目を守りながら周りの画素と平均する
  （SVGF 方式）。動いている間もざらつかず、止まると足し合わせが進んで正確な結果に近づく
- **空・地面**: 雲のある青空と、地平線まで続くベージュの地面（モデルの一番下の高さ）。遠くは霞む（Enscape に近い見え方）
- **照明器具**: ファミリの測光の設定（明るさ・色温度・配光・光源の形）から光源を作り、影も計算する（実際の明るさの単位なので、昼は控えめ・夜はよく見える）
- **マテリアルの外観**（Revit の外観タブ）の色・イメージ（テクスチャ）・反射率・粗さ・透明度・レリーフ パターン（凹凸）・自己発光を使う。
  テクスチャは遠くほど縮小した段（ミップマップ）を使うので、遠くの模様がちらつかない
- 速い GPU では 1 フレームで複数回（1〜4 回、約 30 fps を目安に自動調整）計算して、動いている間の画質を上げる
- **ライブ同期**: Revit でモデルを変更すると、自動で描画ウィンドウに反映される（右上の「同期中」ボタンで止める / 再開）。
  視点は連動しない（描画ウィンドウと Revit のビューは別々に動かす）
- **書き出し**（ウィンドウ右上のボタン）: 画像（PNG）／360°画像（ブラウザで見回せる HTML ＋ JPEG）／
  ビューア（シーンを中に入れた exe。Revit の無い PC でもダブルクリックで歩き回れる）
- 描画は Revit とは別のプログラムで行うので、ここで問題が起きても Revit は止まらない

## 仕組み

```
Revit（Tools28.dll）                       Tools28Render.exe
  RealtimeRenderCommand                      main.cpp        起動・設定
    SceneExportContext  … 3Dビューの形を集める   scene.cpp       シーンファイルの読み込み・確認用の部屋
    AppearanceReader    … マテリアルの外観を読む
    LightReader         … 照明器具の測光の設定を読む
    TextureLoader       … テクスチャ画像を探して読む
    SceneWriter         … %TEMP%\Tools28\Render\scene.t28s を書く
    → Tools28Render.exe --scene … --lang JP    vk_context.cpp  GPU の選択（NVIDIA などの独立 GPU を優先）
                                               bvh.cpp         BVH（予備の方式・マウス位置の物を調べる）
                                               renderer.cpp    加速構造の作成・1フレームの描画・案内の重ね合わせ
                                               shaders/        光の計算（GPU 上で動く）※下の表
                                               window_app.cpp  ウィンドウ・操作 / offscreen.cpp 画像の書き出し
                                               help_overlay.cpp 操作の案内・ツールバー・お知らせの絵（Windows のフォントを使う）
                                               export.cpp      書き出し（保存先の画面・PNG・360°の HTML・ビューアの exe）
```

シーンファイル（`.t28s`）の形式は `src/scene.cpp` の先頭に書いてあります。
**形式を変えるときは `Commands/RealtimeRender/SceneWriter.cs` も同時に直すこと。**

## シェーダー

| ファイル | 内容 |
|--|--|
| `common_head.glsl` / `common_body.glsl` | 共通部分（空・太陽・材質・テクスチャ・光の通り道） |
| `pathtrace_hw.comp` | ハード版。当たり判定を RT コアで行う（Vulkan 1.2 + ray query） |
| `pathtrace_sw.comp` | ソフト版（予備）。BVH を自分でたどる（Vulkan 1.0 で動く） |
| `temporal.comp` | ノイズ除去 その1: 前のフレームを再投影して重ねる。ばらつき（分散）も求める |
| `atrous.comp` | ノイズ除去 その2: 位置・向き・明るさが近い画素だけで平均（間隔 1〜16 画素で 5 回） |
| `composite.comp` | 仕上げ: 面の色を掛け戻し、トーンマッピング、自動露出の集計 |
| `sky.comp` | 空と雲（体積の積雲・巻雲・うろこ雲）を起動時に 1 枚の画像に描く |
| `overlay.comp` | 操作の案内を画面に重ねる |

パストレースは「最初に見えている面の色」で割った明るさを出し、仕上げで掛け戻す（色の境目をぼかさないため）。
ガラスは素通りした先の物を「最初に見えている面」にする（窓の外の景色が動いたときににじまないように）。
鏡（つるつるの金属）は、映っている物の虚像（鏡の向こうにあるとみなした位置・向き）を「最初に見えている面」にする（動いたときに映り込みが残像にならないように）。
そのためガラスは毎回同じ計算（必ず通り抜け＋空の映り込み）にしている。

GPU の選び方（`vk_context.cpp`）: RT コアのある独立 GPU → RT コアのある内蔵 GPU → RT コアの無い独立 GPU → RT コアの無い内蔵 GPU。
ノートPCの「内蔵 GPU ＋ NVIDIA」の構成でも NVIDIA が選ばれる。どれを選んだかはログに出る（画面には出さない）。

## 操作（ウィンドウ）

Revit の 3D ビューと同じ操作を基本にし、Enscape のような歩き回る操作も使える。
ウィンドウの左下に、マウス・キーの絵と短い言葉で案内を出す（`help_overlay.cpp`）。
右上の「×」をクリックすると「?」の記号だけになり、「?」をクリックすると開く（H キーでも切り替え）。

| 操作 | 内容 |
|--|--|
| ホイール | ズーム（カーソルの先の物に近づく） |
| ホイールを押してドラッグ | 移動（パン） |
| Shift ＋ ホイールを押してドラッグ | 回転（オービット。ドラッグを始めた位置の物が中心） |
| ホイールをダブルクリック | 全体表示 |
| 左ドラッグ | その場で見回す |
| W A S D / 矢印キー | 前後左右に歩く（Shift で速く） |
| Q / E | 下がる / 上がる |
| + / - | 明るさ（0.5 EV ずつ） |
| F | 自動露出の切り替え |
| R | 最初の視点に戻る |
| H | 操作の案内を隠す / 出す |
| P | 今の視点を画像で保存（右上の「画像」ボタンと同じ） |
| Esc | 閉じる |

## ライブ同期

```
Revit（LiveSync.cs）                                   Tools28Render.exe（window_app.cpp）
  DocumentChanged → 0.3 秒待って Idling で書き出し直し      0.15 秒ごとに更新時刻を調べる
    → scene_*.t28s.tmp を書いて scene_*.t28s と置き換え  →  読み直す（視点はそのまま）
                          scene_*.t28s.pause があれば休む ←  「同期」ボタンで作る / 消す
```

- シーンファイルは描画ウィンドウごとに `%TEMP%\Tools28\Render\scene_<Revit の番号>_<連番>.t28s`（2 日たったものは次の起動時に消す）
- 大きなモデルでは書き出しに時間がかかり、その間 Revit が止まる。気になるときは「同期」ボタンで止める
- 描画ウィンドウを閉じると、Revit 側の見張りも終わる

## 書き出し

| ボタン | 保存するもの | 大きさ・回数 |
|--|--|--|
| 画像（札: PNG） | PNG（今の視点） | 画面と同じ縦横比で長い辺 1920〜3840 画素。RT コアあり 512 回・なし 128 回足し合わせ |
| 360°（札: HTML） | HTML（画像入り。ブラウザで開くとドラッグで見回せる）と同じ名前の JPEG | 4096×2048（GPU のメモリ 5GB 以上・RT コアあり）または 2048×1024 |
| ビューア（札: EXE） | exe（この exe ＋ シーン） | 今の視点が最初の視点になる |

- 画像・360°は、ウィンドウの描画を一時的に書き出す大きさに切り替えて足し合わせる（進み具合を表示。Esc で中止）。
  明るさは書き出しを始めたときのまま
- ビューアの exe の形式は `src/export.h` の先頭に書いてある（exe の後ろにシーンを付け足し、最後の 16 バイトで位置を示す）。
  起動時に `--scene` が無ければ自分の exe を調べる。言語は Windows の表示言語に合わせる
- ビューアは署名の無い exe なので、受け取った PC で Windows の警告が出ることがある（「詳細情報」→「実行」）

ログ: `%LOCALAPPDATA%\Tools28\Render\render.log`（どの GPU・方式を使ったか、速さ（10 秒ごとの fps）、GPU メモリの使用量、エラーの内容が残る）。
前回の起動分は `render.prev.log` に残る（落ちたあとに開き直しても、落ちたときの記録を調べられる）

## ビルド

Windows 用の exe は **Linux 上で MinGW-w64 を使って作り**、`prebuilt/win-x64/Tools28Render.exe` に置いてリポジトリに入れる。
アドインのビルド（`Tools28.csproj`）はこの exe を出力フォルダにコピーするだけで、C++ のビルドは行わない
（ユーザーの PC に C++ の開発環境が無くても AutoBuild が動くようにするため）。

```bash
# 準備（Ubuntu）: sudo apt install cmake mingw-w64 glslc
./build-windows.sh              # → prebuilt/win-x64/Tools28Render.exe を更新

# Linux 版（動作確認用）: sudo apt install libglfw3-dev mesa-vulkan-drivers
cmake -S . -B build-linux -DT28_SYSTEM_GLFW=ON && cmake --build build-linux -j
./build-linux/Tools28Render --test-scene --output test.png --width 640 --height 360 --samples 64
```

- GPU の無い Linux でも、Mesa の **lavapipe**（CPU で Vulkan を再現する）がレイトレーシングに対応しているので、
  画像書き出しモード（`--output`）で描画結果を確かめられる（遅い）
- シェーダーは glslc があればビルド時に変換され、`shaders/generated/*.spv.inc` も更新される（これもリポジトリに入れる）
- 予備の方式は `--software` で強制できるので、RT コアのある環境でも動作確認できる
- ウィンドウ表示の確認: Xvfb ＋ xdotool でマウス操作を再現して画面を撮れる。Windows 版は Wine で動く
  （Wine には日本語フォントが無いので、`drive_c/windows/Fonts/YuGothM.ttc` に Noto CJK を置くと案内の確認ができる）
- ダウンロードは git で行う（この開発環境では GitHub の圧縮ファイルのダウンロードが通らないため）

## コマンドライン

```
Tools28Render --scene model.t28s [--lang JP|US|CN]     ウィンドウで表示
Tools28Render --test-scene                              確認用の部屋
Tools28Render --scene model.t28s --output out.png       画像を作って終わる
    [--width 1280 --height 720 --samples 256 --panorama]
その他: --view ビュー名（書き出すファイルの名前「日付_ビュー名.拡張子」に使う。無ければ --title の「 - 」より後ろ）/ --sky-size 横の画素数（空の画像。既定 4096。動作確認で速くするときは 1024）/ --live（シーンファイル・.cam の更新を見張る。Revit から起動したとき）/ --title 題名（ウィンドウと書き出すファイルの名前）/ --export-samples 回数（書き出しを短くする。動作確認用）
        --log ファイル / --gpu 番号 / --bounces 回数 / --exposure EV / --no-auto-exposure
        --render-scale 倍率（小さく描いて拡大＝速い。既定は RT コアあり 1.0・なし 0.5）
        --software（RT コアがあっても予備の方式で描く） / --exit-after フレーム数（動作確認用）
        --motion-test フレーム数 / --motion-reference フレーム数（動いている間の残像の確認。DEVLOG 参照）
```

## アイコン

`res/Tools28Render.ico`（28 Tools のサイトのファビコン `res/favicon_28.png` から `res/make_app_icon.py` で作る）を `res/app.rc` で exe に入れている。
エクスプローラー・ウィンドウ・タスクバーのアイコンになり、書き出したビューア exe も同じアイコンになる（28 Tools から書き出したと分かる）。

## 使っている外部の部品

| 部品 | 用途 | ライセンス | 入れ方 |
|--|--|--|--|
| Vulkan-Headers | Vulkan の定義 | Apache-2.0 / MIT | ビルド時に git で取得 |
| volk | Vulkan の読み込み | MIT | `third_party/volk` |
| GLFW | ウィンドウ | zlib | ビルド時に git で取得（Linux はシステムのもの） |
| stb_image_write | PNG 保存 | Public Domain / MIT | `third_party/stb` |
| stb_truetype | 案内の文字を描く（フォントは Windows のものを使う） | Public Domain / MIT | `third_party/stb` |

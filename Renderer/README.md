# Tools28Render — レイトレーシングの描画プログラム

Revit の「リアルタイムレンダリング」ボタン（`Commands/RealtimeRender/`）から起動される、
別プログラム（exe）の描画エンジンです。C++ と Vulkan で書かれています。

- GPU の**光の計算専用回路**（レイトレーシング、`VK_KHR_ray_query`）で、影・反射・間接光を計算する
- 動いている間は1回分の計算、止まると計算結果を足し合わせて画質が上がっていく
- 描画は Revit とは別のプログラムで行うので、ここで問題が起きても Revit は止まらない

## 仕組み

```
Revit（Tools28.dll）                       Tools28Render.exe
  RealtimeRenderCommand                      main.cpp        起動・設定
    SceneExportContext  … 3Dビューの形を集める   scene.cpp       シーンファイルの読み込み・確認用の部屋
    SceneWriter         … %TEMP%\Tools28\Render\scene.t28s を書く
    → Tools28Render.exe --scene … --lang JP    vk_context.cpp  GPU の選択（NVIDIA などの独立 GPU を優先）
                                               renderer.cpp    加速構造の作成・1フレームの描画
                                               shaders/pathtrace.comp  光の計算（GPU 上で動く）
                                               window_app.cpp  ウィンドウ・操作 / offscreen.cpp 画像の書き出し
```

シーンファイル（`.t28s`）の形式は `src/scene.cpp` の先頭に書いてあります。
**形式を変えるときは `Commands/RealtimeRender/SceneWriter.cs` も同時に直すこと。**

## 操作（ウィンドウ）

| 操作 | 内容 |
|--|--|
| 左ドラッグ | 見回す |
| W A S D / 矢印キー | 前後左右に移動（Shift で速く） |
| Q / E | 下がる / 上がる |
| ホイール | 移動の速さ |
| + / - | 明るさ（0.5 EV ずつ） |
| F | 自動露出の切り替え |
| R | 最初の位置に戻る |
| Esc | 閉じる |

ログ: `%LOCALAPPDATA%\Tools28\Render\render.log`（どの GPU を使ったか、エラーの内容が残る）

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
- シェーダーは glslc があればビルド時に変換され、`shaders/generated/pathtrace.spv.inc` も更新される（これもリポジトリに入れる）
- ダウンロードは git で行う（この開発環境では GitHub の圧縮ファイルのダウンロードが通らないため）

## コマンドライン

```
Tools28Render --scene model.t28s [--lang JP|US|CN]     ウィンドウで表示
Tools28Render --test-scene                              確認用の部屋
Tools28Render --scene model.t28s --output out.png       画像を作って終わる
    [--width 1280 --height 720 --samples 256 --panorama]
その他: --log ファイル / --gpu 番号 / --bounces 回数 / --exposure EV / --no-auto-exposure
        --render-scale 倍率（小さく描いて拡大＝速い） / --exit-after フレーム数（動作確認用）
```

## 使っている外部の部品

| 部品 | 用途 | ライセンス | 入れ方 |
|--|--|--|--|
| Vulkan-Headers | Vulkan の定義 | Apache-2.0 / MIT | ビルド時に git で取得 |
| volk | Vulkan の読み込み | MIT | `third_party/volk` |
| GLFW | ウィンドウ | zlib | ビルド時に git で取得（Linux はシステムのもの） |
| stb_image_write | PNG 保存 | Public Domain / MIT | `third_party/stb` |

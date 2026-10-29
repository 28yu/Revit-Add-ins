#pragma once
// 書き出し: 画像（PNG）・360°パノラマ（ブラウザで見られる HTML ＋ JPEG）・ビューア（シーンを埋め込んだ exe）。
//
// ビューアの exe は、この exe の後ろにシーンをそのまま付け足したもの:
//   [この exe] [ "T28V" | uint32 題名のバイト数 | 題名（UTF-8） | シーン（.t28s と同じ形式） ] [ uint64 付け足した位置 | "T28EMBED" ]
// 起動したときに自分の exe の最後を調べ、付け足しがあればそのシーンを開く（Revit もアドインも要らない）。
#include <cstdint>
#include <string>
#include <vector>

#include "scene.h"

namespace t28 {

// Image は JPEG（既定）か PNG（背景を透明にするとき・保存先で .png を選んだとき）。保存先の拡張子で決める
enum class ExportKind { Image, Panorama, Viewer };

// 保存先を選ぶ画面を出す（Windows）。キャンセルなら空。
// Windows 以外（動作確認用）は、環境変数 T28_EXPORT_DIR（無ければ今のフォルダ）に既定の名前で保存する
// 画像は JPEG と PNG を選べる（defaultName の拡張子のほうを先に出す）
std::string chooseSavePath(void* ownerWindow, ExportKind kind, const std::string& defaultName);
// 色を選ぶ画面を出す（Windows の標準の画面。色の調整もできる）。color は 0xBBGGRR。選んだら true。
// Windows 以外（動作確認用）は、決まった色を順に切り替える
bool chooseColor(void* ownerWindow, uint32_t& color);
// 既定のファイル名（日付_ビュー名.拡張子。例: 261016_3D 南東.png）。ファイル名に使えない文字は _ にする
// png: 画像を PNG にする（背景を透明にするとき）。それ以外の画像は .jpg
std::string defaultExportName(const std::string& viewName, ExportKind kind, bool png = false);
// パスのファイル名の部分
std::string fileNameOf(const std::string& path);

bool writePng(const std::string& path, const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height);
// 画像を保存する。拡張子が .png なら PNG、それ以外は JPEG（品質 92。PNG の 1/5〜1/10 ほどの大きさ）
bool writeImage(const std::string& path, const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height);
// 360°パノラマ: htmlPath に、画像を中に入れた HTML（ブラウザで開くと見回せる）を書く。
// 画像は HTML の中だけに持つ（以前は同じ名前の .jpg も横に置いていたが、同じ画像を 2 回持つことになるのでやめた）。
// 他の 360°ビューアや SNS で使うときは、HTML の「画像を保存」ボタンで JPEG を取り出せる
bool writePanoramaHtml(const std::string& htmlPath, const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height,
                       const std::string& title);
// ビューア: この exe に、詰めたシーン（scene_pack.h。camera が最初の視点）を付け足して path に書く
bool writeViewerExe(const std::string& path, const Scene& scene, const Camera& camera, const std::string& title, std::string& error);
// 自分の exe にシーンが付け足されていれば、その位置と題名を返す
bool findEmbeddedScene(std::string& exePath, long& sceneOffset, std::string& title);
// 付け足されたシーンを読む（詰めたシーンでも、以前の形式（.t28s そのまま）でも読める）
bool loadEmbeddedScene(const std::string& exePath, long sceneOffset, Scene& scene, std::string& error);
// Windows の表示言語（"JP" / "CN" / "US"）
const char* systemLanguage();

}  // namespace t28

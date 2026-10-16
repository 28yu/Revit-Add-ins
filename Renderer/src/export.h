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

enum class ExportKind { Image, Panorama, Viewer };

// 保存先を選ぶ画面を出す（Windows）。キャンセルなら空。
// Windows 以外（動作確認用）は、環境変数 T28_EXPORT_DIR（無ければ今のフォルダ）に既定の名前で保存する
std::string chooseSavePath(void* ownerWindow, ExportKind kind, const std::string& defaultName);
// 既定のファイル名（題名_年月日_時分.png など）。題名に使えない文字は _ にする
std::string defaultExportName(const std::string& title, ExportKind kind);
// パスのファイル名の部分
std::string fileNameOf(const std::string& path);

bool writePng(const std::string& path, const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height);
// 360°パノラマ: htmlPath に、画像を中に入れた HTML（ブラウザで開くと見回せる）を書く。
// 同じ名前の .jpg も横に置く（他の 360°ビューアでも使えるように）
bool writePanoramaHtml(const std::string& htmlPath, const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height,
                       const std::string& title);
// ビューア: この exe にシーンを付け足して path に書く
bool writeViewerExe(const std::string& path, const Scene& scene, const std::string& title, std::string& error);
// 自分の exe にシーンが付け足されていれば、その位置と題名を返す
bool findEmbeddedScene(std::string& exePath, long& sceneOffset, std::string& title);
// Windows の表示言語（"JP" / "CN" / "US"）
const char* systemLanguage();

}  // namespace t28

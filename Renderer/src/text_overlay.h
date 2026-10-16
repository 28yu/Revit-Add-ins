#pragma once
// 操作の案内を、半透明の背景つきの文字の画像（RGBA）にする。
// フォントは Windows に入っているもの（言語ごとに選ぶ）を使うので、exe にフォントを入れなくてよい。
#include <cstdint>
#include <string>
#include <vector>

namespace t28 {

struct TextImage {
    std::vector<uint8_t> rgba;  // 1画素 4 バイト。a = 不透明度
    uint32_t width = 0, height = 0;
};

// lines: UTF-8 の文字列（1つが1行）。pixelHeight: 文字の高さ（画素）
// 使えるフォントが見つからなければ false
bool renderTextPanel(const std::vector<std::string>& lines, float pixelHeight, TextImage& out);

}  // namespace t28

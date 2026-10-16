#pragma once
// 操作の案内（ウィンドウの左下に出す半透明の絵）を作る。
//  - 開いているとき: マウス・キーの絵と短い言葉を並べ、右上に閉じる記号「×」
//  - 閉じているとき: 「?」の記号だけ
// 記号はクリックでも H キーでも切り替えられる。
// 文字は Windows に入っているフォント（言語ごとに選ぶ）で描くので、exe にフォントを入れなくてよい。
#include <cstdint>
#include <vector>

namespace t28 {

struct HelpImage {
    std::vector<uint8_t> rgba;  // 1画素 4 バイト。a = 不透明度（重ねる側は mix(元の色, rgb, a) で合成する）
    uint32_t width = 0, height = 0;
    // 切り替えの記号の範囲（画像の左上からの画素）。ここをクリックすると開く / 閉じる
    int toggleX = 0, toggleY = 0, toggleW = 0, toggleH = 0;
};

// expanded: 開いた状態か。scale: 画面の拡大率（高 DPI の画面では 1.5 など）
// フォントが見つからなくても絵だけは作る（文字が欠けるだけ）
HelpImage buildHelpImage(bool expanded, float scale);

}  // namespace t28

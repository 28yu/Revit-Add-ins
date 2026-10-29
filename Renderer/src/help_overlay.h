#pragma once
// 操作の案内（ウィンドウの左下に出す半透明の絵）を作る。
//  - 開いているとき: マウス・キーの絵と短い言葉を並べ、右上に閉じる記号「×」
//  - 閉じているとき: 「?」の記号だけ
// 記号はクリックでも H キーでも切り替えられる。
// 文字は Windows に入っているフォント（言語ごとに選ぶ）で描くので、exe にフォントを入れなくてよい。
#include <cstdint>
#include <string>
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

// 右上のツールバー。ボタンは左から「画像」「360°」「ビューア（exe）」と、Revit から開いたときだけ「同期」
struct ToolbarImage {
    HelpImage image;
    struct Rect { int x, y, w, h; };
    std::vector<Rect> buttons;  // 画像の左上からの画素
};
enum ToolbarButton { kButtonImage = 0, kButtonPanorama = 1, kButtonViewer = 2, kButtonSync = 3, kButtonCount = 4 };
// hover: カーソルが乗っているボタン（-1 = 無し）。sync: 同期のボタン（-1 = 出さない、0 = 止めている、1 = 同期中）
// width: この幅より狭ければボタンを広げてこの幅にする（0 = 中身に合わせる。下の「表示」の設定と幅をそろえるため）
// imagePng: 画像の書き出しが PNG（背景を透明にするとき）。それ以外は JPG の札を出す
ToolbarImage buildToolbarImage(float scale, int hover, int sync, int width = 0, bool imagePng = false);

// ツールバーの下の「表示」の設定。上から
//   空・雲を表示（＋背景色）／地面・地平線を表示（＋地面の色）／太陽の光／照明器具／単色で表示（＋色）／
//   PNG の背景を透明に／時間帯（つまみ）
enum LookButton {
    kLookSky = 0, kLookBackground = 1, kLookMono = 2, kLookMonoColor = 3, kLookGround = 4, kLookTransparent = 5,
    kLookGroundColor = 6, kLookTimeReset = 7, kLookTime = 8, kLookSun = 9, kLookLights = 10, kLookCount = 11
};
// 時間帯のつまみの範囲（時）
const float kTimeFirstHour = 5.0f, kTimeLastHour = 19.0f;
struct LookState {
    bool sky = true;                  // 空・雲を出す（出さないときは背景色）
    uint32_t background = 0xFFFFFF;   // 背景色（0xBBGGRR）
    bool ground = true;               // 地面（と地平線）を出す
    bool mono = false;                // 物を全部 1 色にする
    uint32_t monoColor = 0xF0F0F0;    // 単色の色（0xBBGGRR）
    bool transparentPng = false;      // 画像（PNG）の書き出しで、背景を透明にする
    uint32_t groundColor = 0xA2A8B0;  // 地面の色（0xBBGGRR。既定はこれまでの地面の色）
    float timeHour = -1.0f;           // 時間帯（時。負なら Revit の「太陽の設定」のまま）
    float revitHour = 12.0f;          // Revit の太陽が何時ごろか（つまみの位置に使う）
    bool sun = true;                  // 太陽の光（直射と空の明るさ）
    bool lights = true;               // 照明器具の光
};
// hover: カーソルが乗っているボタン（-1 = 無し）
// width: この幅より狭ければ広げる（0 = 中身に合わせる）
ToolbarImage buildLookPanelImage(float scale, int hover, const LookState& look, int width = 0);

// 画面の上の中央に出すお知らせ（ボタンの説明・書き出しの進み具合・保存しました 等）。
// progress: 0〜1 なら進み具合の棒を出す（負なら出さない）。lines は 1〜2 行
HelpImage buildStatusImage(const std::vector<std::string>& lines, float progress, float scale);

}  // namespace t28

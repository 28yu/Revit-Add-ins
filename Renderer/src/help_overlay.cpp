#include "help_overlay.h"

#include <algorithm>
#include <cmath>
#include <cstdio>
#include <cstdlib>
#include <functional>
#include <string>

#include "i18n.h"
#include "log.h"

#define STB_TRUETYPE_IMPLEMENTATION
#include <stb_truetype.h>

namespace t28 {

namespace {

// --- フォント -----------------------------------------------------------------

// 言語ごとのフォントの候補（上から順に探す）。見つからなければ他の言語の候補も試す
std::vector<std::string> fontCandidates() {
    std::vector<std::string> jp, cn, us, all;
#ifdef _WIN32
    const char* win = std::getenv("WINDIR");
    std::string dir = std::string(win ? win : "C:\\Windows") + "\\Fonts\\";
    jp = {dir + "YuGothM.ttc", dir + "YuGothR.ttc", dir + "meiryo.ttc", dir + "msgothic.ttc"};
    cn = {dir + "msyh.ttc", dir + "msyh.ttf", dir + "simsun.ttc"};
    us = {dir + "segoeui.ttf", dir + "arial.ttf"};
#else
    jp = cn = {"/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
               "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"};
    us = {"/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"};
#endif
    switch (currentLanguage()) {
        case Lang::CN: all = cn; all.insert(all.end(), jp.begin(), jp.end()); break;
        case Lang::US: all = us; all.insert(all.end(), jp.begin(), jp.end()); break;
        default: all = jp; all.insert(all.end(), cn.begin(), cn.end()); break;
    }
    return all;
}

bool loadFile(const std::string& path, std::vector<unsigned char>& data) {
    FILE* f = openFile(path, "rb");
    if (!f) return false;
    fseek(f, 0, SEEK_END);
    long size = ftell(f);
    fseek(f, 0, SEEK_SET);
    data.resize(size > 0 ? static_cast<size_t>(size) : 0);
    bool ok = size > 0 && fread(data.data(), 1, data.size(), f) == data.size();
    fclose(f);
    return ok;
}

// フォントは一度だけ読む。見つからなければ nullptr
const stbtt_fontinfo* font() {
    static std::vector<unsigned char> data;
    static stbtt_fontinfo info;
    static int state = 0;  // 0 = まだ, 1 = 読めた, -1 = 無い
    if (state == 0) {
        state = -1;
        for (const std::string& path : fontCandidates()) {
            if (!loadFile(path, data)) continue;
            int offset = stbtt_GetFontOffsetForIndex(data.data(), 0);
            if (offset >= 0 && stbtt_InitFont(&info, data.data(), offset)) {
                logInfo("Help font: %s", path.c_str());
                state = 1;
                break;
            }
        }
        if (state < 0) logError("No font found for the help text");
    }
    return state > 0 ? &info : nullptr;
}

// UTF-8 の文字列を文字（コードポイント）の並びにする
std::vector<int> decodeUtf8(const std::string& s) {
    std::vector<int> out;
    for (size_t i = 0; i < s.size();) {
        unsigned char c = static_cast<unsigned char>(s[i]);
        int cp = c, extra = 0;
        if (c >= 0xF0) { cp = c & 0x07; extra = 3; }
        else if (c >= 0xE0) { cp = c & 0x0F; extra = 2; }
        else if (c >= 0xC0) { cp = c & 0x1F; extra = 1; }
        ++i;
        for (int k = 0; k < extra && i < s.size(); ++k, ++i) cp = (cp << 6) | (static_cast<unsigned char>(s[i]) & 0x3F);
        out.push_back(cp);
    }
    return out;
}

float textWidth(const std::string& text, float px) {
    const stbtt_fontinfo* f = font();
    if (!f) return 0;
    float scale = stbtt_ScaleForPixelHeight(f, px);
    std::vector<int> cps = decodeUtf8(text);
    float w = 0;
    for (size_t i = 0; i < cps.size(); ++i) {
        int adv, lsb;
        stbtt_GetCodepointHMetrics(f, cps[i], &adv, &lsb);
        w += adv * scale;
        if (i + 1 < cps.size()) w += stbtt_GetCodepointKernAdvance(f, cps[i], cps[i + 1]) * scale;
    }
    return w;
}

// --- 描画用のキャンバス（色は「不透明度を掛けた形」で持ち、上に重ねていく） -------------

struct Color {
    float r, g, b, a;
};

const Color kWhite{0.96f, 0.96f, 0.96f, 1.0f};
const Color kDim{0.96f, 0.96f, 0.96f, 0.55f};
const Color kAccent{0.30f, 0.64f, 1.0f, 1.0f};   // 使うボタンを示す青
const Color kPanel{0.0f, 0.0f, 0.0f, 0.55f};

struct Canvas {
    int w, h;
    std::vector<float> px;  // r, g, b, a（r, g, b には a を掛けておく）
    Canvas(int width, int height) : w(width), h(height), px(static_cast<size_t>(width) * height * 4, 0.0f) {}

    void blend(int x, int y, const Color& c, float coverage) {
        if (x < 0 || y < 0 || x >= w || y >= h || coverage <= 0) return;
        float a = c.a * std::min(coverage, 1.0f);
        float* p = &px[(static_cast<size_t>(y) * w + x) * 4];
        p[0] = c.r * a + p[0] * (1 - a);
        p[1] = c.g * a + p[1] * (1 - a);
        p[2] = c.b * a + p[2] * (1 - a);
        p[3] = a + p[3] * (1 - a);
    }

    // sdf: 点 (x, y) から形の縁までの距離（内側が負）。範囲 [x0, x1) × [y0, y1) だけを調べる
    void fill(const std::function<float(float, float)>& sdf, float x0, float y0, float x1, float y1, const Color& c) {
        for (int y = std::max(0, static_cast<int>(y0) - 1); y < std::min(h, static_cast<int>(y1) + 2); ++y)
            for (int x = std::max(0, static_cast<int>(x0) - 1); x < std::min(w, static_cast<int>(x1) + 2); ++x)
                blend(x, y, c, 0.5f - sdf(x + 0.5f, y + 0.5f));
    }
    void stroke(const std::function<float(float, float)>& sdf, float width, float x0, float y0, float x1, float y1,
                const Color& c) {
        fill([&](float x, float y) { return std::fabs(sdf(x, y)) - width * 0.5f; }, x0 - width, y0 - width,
             x1 + width, y1 + width, c);
    }

    void text(const std::string& s, float x, float baseline, float pxSize, const Color& c) {
        const stbtt_fontinfo* f = font();
        if (!f) return;
        float scale = stbtt_ScaleForPixelHeight(f, pxSize);
        std::vector<int> cps = decodeUtf8(s);
        for (size_t i = 0; i < cps.size(); ++i) {
            int adv, lsb;
            stbtt_GetCodepointHMetrics(f, cps[i], &adv, &lsb);
            float shift = x - std::floor(x);
            int gx0, gy0, gx1, gy1;
            stbtt_GetCodepointBitmapBoxSubpixel(f, cps[i], scale, scale, shift, 0, &gx0, &gy0, &gx1, &gy1);
            int gw = gx1 - gx0, gh = gy1 - gy0;
            if (gw > 0 && gh > 0) {
                std::vector<unsigned char> glyph(static_cast<size_t>(gw) * gh);
                stbtt_MakeCodepointBitmapSubpixel(f, glyph.data(), gw, gh, gw, scale, scale, shift, 0, cps[i]);
                int ox = static_cast<int>(std::floor(x)) + gx0, oy = static_cast<int>(std::round(baseline)) + gy0;
                for (int r = 0; r < gh; ++r)
                    for (int col = 0; col < gw; ++col)
                        blend(ox + col, oy + r, c, glyph[static_cast<size_t>(r) * gw + col] / 255.0f);
            }
            x += adv * scale;
            if (i + 1 < cps.size()) x += stbtt_GetCodepointKernAdvance(f, cps[i], cps[i + 1]) * scale;
        }
    }
    // 中央ぞろえの文字（cy は文字の縦の中心）
    void textCentered(const std::string& s, float cx, float cy, float pxSize, const Color& c) {
        text(s, cx - textWidth(s, pxSize) / 2, cy + pxSize * 0.36f, pxSize, c);
    }
};

// --- 形の距離関数 ---------------------------------------------------------------

float sdRoundBox(float x, float y, float cx, float cy, float hw, float hh, float r) {
    float qx = std::fabs(x - cx) - (hw - r), qy = std::fabs(y - cy) - (hh - r);
    float ox = std::max(qx, 0.0f), oy = std::max(qy, 0.0f);
    return std::sqrt(ox * ox + oy * oy) + std::min(std::max(qx, qy), 0.0f) - r;
}

float sdSegment(float x, float y, float ax, float ay, float bx, float by) {
    float px = x - ax, py = y - ay, dx = bx - ax, dy = by - ay;
    float t = std::clamp((px * dx + py * dy) / (dx * dx + dy * dy), 0.0f, 1.0f);
    float ex = px - dx * t, ey = py - dy * t;
    return std::sqrt(ex * ex + ey * ey);
}

// 三角形（頂点は時計回り・反時計回りどちらでもよい）
float sdTriangle(float x, float y, const float v[6]) {
    float area = (v[2] - v[0]) * (v[5] - v[1]) - (v[3] - v[1]) * (v[4] - v[0]);
    float s = area > 0 ? 1.0f : -1.0f;
    float d = -1e9f;
    for (int i = 0; i < 3; ++i) {
        float ax = v[i * 2], ay = v[i * 2 + 1], bx = v[(i * 2 + 2) % 6], by = v[(i * 2 + 3) % 6];
        float ex = bx - ax, ey = by - ay;
        float len = std::sqrt(ex * ex + ey * ey);
        // 辺の外向きの法線方向の距離
        d = std::max(d, s * ((x - ax) * ey - (y - ay) * ex) / len);
    }
    return d;
}

// --- 絵の部品 ---------------------------------------------------------------------

// 矢印（a → b、先に三角の頭）
void arrow(Canvas& c, float ax, float ay, float bx, float by, float s, const Color& col) {
    float dx = bx - ax, dy = by - ay, len = std::sqrt(dx * dx + dy * dy);
    float ux = dx / len, uy = dy / len;
    float head = 3.6f * s;
    float lx = bx - ux * head, ly = by - uy * head;
    c.stroke([&](float x, float y) { return sdSegment(x, y, ax, ay, lx, ly); }, 1.4f * s,
             std::min(ax, lx), std::min(ay, ly), std::max(ax, lx), std::max(ay, ly), col);
    float v[6] = {bx, by, lx - uy * head * 0.75f, ly + ux * head * 0.75f, lx + uy * head * 0.75f, ly - ux * head * 0.75f};
    c.fill([&](float x, float y) { return sdTriangle(x, y, v); }, bx - head * 2, by - head * 2, bx + head * 2,
           by + head * 2, col);
}

enum class MouseButton { Left, Wheel };

// マウスの絵（cx, cy = 中心）。使うボタンを青く塗る
void mouse(Canvas& c, float cx, float cy, float s, MouseButton active) {
    const float hw = 9 * s, hh = 13 * s, r = 8.5f * s;
    const float top = cy - hh, split = top + 10.5f * s;
    auto body = [&](float x, float y) { return sdRoundBox(x, y, cx, cy, hw, hh, r); };
    if (active == MouseButton::Left) {
        // 左ボタン = 本体の左上の部分
        c.fill([&](float x, float y) { return std::max({body(x, y), x - cx, y - split}); }, cx - hw, top, cx, split, kAccent);
    }
    c.stroke(body, 1.5f * s, cx - hw, cy - hh, cx + hw, cy + hh, kWhite);
    c.stroke([&](float x, float y) { return sdSegment(x, y, cx - hw, split, cx + hw, split); }, 1.2f * s, cx - hw, split,
             cx + hw, split, kWhite);
    c.stroke([&](float x, float y) { return sdSegment(x, y, cx, top, cx, split); }, 1.2f * s, cx, top, cx, split, kWhite);
    // ホイール
    float wy = top + 5.6f * s;
    auto wheel = [&](float x, float y) { return sdRoundBox(x, y, cx, wy, 2.2f * s, 3.6f * s, 2.2f * s); };
    c.fill(wheel, cx - 3 * s, wy - 4 * s, cx + 3 * s, wy + 4 * s, active == MouseButton::Wheel ? kAccent : kPanel);
    c.stroke(wheel, 1.1f * s, cx - 3 * s, wy - 4 * s, cx + 3 * s, wy + 4 * s, active == MouseButton::Wheel ? kAccent : kWhite);
}

// キーの絵。幅は文字に合わせる。戻り値 = 幅
float keycapWidth(const std::string& label, float s) { return std::max(17.0f * s, textWidth(label, 10.5f * s) + 9 * s); }
void keycap(Canvas& c, float x, float cy, const std::string& label, float s) {
    float w = keycapWidth(label, s), h = 17 * s;
    auto box = [&](float px, float py) { return sdRoundBox(px, py, x + w / 2, cy, w / 2, h / 2, 3.5f * s); };
    c.fill(box, x, cy - h / 2, x + w, cy + h / 2, Color{1, 1, 1, 0.08f});
    c.stroke(box, 1.2f * s, x, cy - h / 2, x + w, cy + h / 2, kWhite);
    c.textCentered(label, x + w / 2, cy, 10.5f * s, kWhite);
}
float keysWidth(const std::vector<std::string>& keys, float s) {
    float w = 0;
    for (const auto& k : keys) w += keycapWidth(k, s) + 3 * s;
    return w - 3 * s;
}
void keys(Canvas& c, float x, float cy, const std::vector<std::string>& labels, float s) {
    for (const auto& k : labels) {
        keycap(c, x, cy, k, s);
        x += keycapWidth(k, s) + 3 * s;
    }
}

// 回転の矢印（円弧＋矢じり）
void orbitArrow(Canvas& c, float cx, float cy, float rad, float s) {
    const float a0 = -2.6f, a1 = 1.4f;  // 円弧の範囲（ラジアン、画面の下向きが正）
    auto arc = [&](float x, float y) {
        float dx = x - cx, dy = y - cy;
        float ang = std::atan2(dy, dx);
        if (ang >= a0 && ang <= a1) return std::fabs(std::sqrt(dx * dx + dy * dy) - rad);
        float e0x = cx + std::cos(a0) * rad, e0y = cy + std::sin(a0) * rad;
        float e1x = cx + std::cos(a1) * rad, e1y = cy + std::sin(a1) * rad;
        return std::min(std::hypot(x - e0x, y - e0y), std::hypot(x - e1x, y - e1y));
    };
    c.fill([&](float x, float y) { return arc(x, y) - 0.7f * s; }, cx - rad - 2 * s, cy - rad - 2 * s, cx + rad + 2 * s,
           cy + rad + 2 * s, kWhite);
    // 矢じり（円弧の終わりで接線の向き）
    float ex = cx + std::cos(a1) * rad, ey = cy + std::sin(a1) * rad;
    float tx = -std::sin(a1), ty = std::cos(a1);  // 角度が増える向きの接線
    float head = 3.6f * s;
    float v[6] = {ex + tx * head, ey + ty * head, ex - ty * head * 0.8f, ey + tx * head * 0.8f,
                  ex + ty * head * 0.8f, ey - tx * head * 0.8f};
    c.fill([&](float x, float y) { return sdTriangle(x, y, v); }, ex - head * 2, ey - head * 2, ex + head * 2,
           ey + head * 2, kWhite);
}

// 上下左右の矢印（移動・見回す）
void crossArrows(Canvas& c, float cx, float cy, float s) {
    float r = 7 * s;
    arrow(c, cx, cy, cx + r, cy, s, kWhite);
    arrow(c, cx, cy, cx - r, cy, s, kWhite);
    arrow(c, cx, cy, cx, cy + r, s, kWhite);
    arrow(c, cx, cy, cx, cy - r, s, kWhite);
}

// 案内の項目: 絵（幅 iconW）と、その下の短い言葉
struct Item {
    float iconW;
    std::function<void(Canvas&, float cx, float cy)> draw;
    std::string label;
};

HelpImage finish(const Canvas& c) {
    HelpImage img;
    img.width = static_cast<uint32_t>(c.w);
    img.height = static_cast<uint32_t>(c.h);
    img.rgba.resize(static_cast<size_t>(c.w) * c.h * 4);
    for (size_t i = 0; i < static_cast<size_t>(c.w) * c.h; ++i) {
        float a = c.px[i * 4 + 3];
        for (int k = 0; k < 3; ++k) {
            float v = a > 1e-4f ? c.px[i * 4 + k] / a : 0.0f;  // 重ねる側は「不透明度を掛けていない色」を使う
            img.rgba[i * 4 + k] = static_cast<uint8_t>(std::round(std::clamp(v, 0.0f, 1.0f) * 255));
        }
        img.rgba[i * 4 + 3] = static_cast<uint8_t>(std::round(std::clamp(a, 0.0f, 1.0f) * 255));
    }
    return img;
}

}  // namespace

HelpImage buildHelpImage(bool expanded, float s) {
    if (!expanded) {
        // 閉じているとき: 「?」の丸い記号だけ
        int d = static_cast<int>(std::ceil(28 * s));
        Canvas c(d, d);
        float r = d / 2.0f;
        auto circle = [&](float x, float y) { return std::hypot(x - r, y - r) - (r - 1.5f * s); };
        c.fill(circle, 0, 0, static_cast<float>(d), static_cast<float>(d), kPanel);
        c.stroke(circle, 1.3f * s, 0, 0, static_cast<float>(d), static_cast<float>(d), kDim);
        c.textCentered("?", r, r, 16 * s, kWhite);
        HelpImage img = finish(c);
        img.toggleX = 0;
        img.toggleY = 0;
        img.toggleW = d;
        img.toggleH = d;
        return img;
    }

    const float mouseW = 18 * s, glyphW = 18 * s, gap = 4 * s;
    std::vector<Item> mouseRow = {
        {mouseW + gap + glyphW,
         [&](Canvas& c, float cx, float cy) {
             float mx = cx - (gap + glyphW) / 2, gx = cx + (mouseW + gap) / 2;
             mouse(c, mx, cy, s, MouseButton::Wheel);
             arrow(c, gx, cy, gx, cy - 9 * s, s, kWhite);
             arrow(c, gx, cy, gx, cy + 9 * s, s, kWhite);
         },
         tr("HelpZoom")},
        {mouseW + gap + glyphW,
         [&](Canvas& c, float cx, float cy) {
             mouse(c, cx - (gap + glyphW) / 2, cy, s, MouseButton::Wheel);
             crossArrows(c, cx + (mouseW + gap) / 2, cy, s);
         },
         tr("HelpPan")},
        {keycapWidth("Shift", s) + gap + mouseW + gap + glyphW,
         [&](Canvas& c, float cx, float cy) {
             float total = keycapWidth("Shift", s) + gap + mouseW + gap + glyphW;
             float x = cx - total / 2;
             keycap(c, x, cy, "Shift", s);
             x += keycapWidth("Shift", s) + gap;
             mouse(c, x + mouseW / 2, cy, s, MouseButton::Wheel);
             orbitArrow(c, x + mouseW + gap + glyphW / 2, cy, 6.5f * s, s);
         },
         tr("HelpOrbit")},
        {mouseW + gap + glyphW,
         [&](Canvas& c, float cx, float cy) {
             mouse(c, cx - (gap + glyphW) / 2, cy, s, MouseButton::Wheel);
             c.textCentered("×2", cx + (mouseW + gap) / 2, cy, 11.5f * s, kWhite);
         },
         tr("HelpFit")},
        {mouseW + gap + glyphW,
         [&](Canvas& c, float cx, float cy) {
             mouse(c, cx - (gap + glyphW) / 2, cy, s, MouseButton::Left);
             crossArrows(c, cx + (mouseW + gap) / 2, cy, s);
         },
         tr("HelpLook")},
    };
    std::vector<std::vector<std::string>> keyIcons = {{"W", "A", "S", "D"}, {"Q", "E"}, {"+", "−"}, {"R"}};
    std::vector<Item> keyRow;
    const char* keyLabels[] = {"HelpWalk", "HelpUpDown", "HelpBright", "HelpReset"};
    for (size_t i = 0; i < keyIcons.size(); ++i) {
        std::vector<std::string> k = keyIcons[i];
        keyRow.push_back({keysWidth(k, s), [=](Canvas& c, float cx, float cy) { keys(c, cx - keysWidth(k, s) / 2, cy, k, s); },
                          tr(keyLabels[i])});
    }

    // 大きさを決める（マウスの項目とキーの項目を 1 段に並べ、間を細い縦線で区切る）
    const float pad = 10 * s, labelPx = 11.5f * s, iconH = 28 * s, labelH = 15 * s, colGap = 12 * s;
    const float sepGap = 14 * s, closeSize = 18 * s;
    auto columnWidth = [&](const Item& it) { return std::max(it.iconW, textWidth(it.label, labelPx)); };
    auto rowWidth = [&](const std::vector<Item>& row) {
        float w = 0;
        for (const auto& it : row) w += columnWidth(it) + colGap;
        return w - colGap;
    };
    float mouseW2 = rowWidth(mouseRow), keyW = rowWidth(keyRow);
    float contentW = mouseW2 + sepGap * 2 + keyW;
    int width = static_cast<int>(std::ceil(pad * 2 + contentW + sepGap + closeSize));
    int height = static_cast<int>(std::ceil(pad * 2 + iconH + labelH));
    Canvas c(width, height);

    // 背景（角の丸い半透明の黒）
    auto panel = [&](float x, float y) { return sdRoundBox(x, y, width / 2.0f, height / 2.0f, width / 2.0f, height / 2.0f, 8 * s); };
    c.fill(panel, 0, 0, static_cast<float>(width), static_cast<float>(height), kPanel);

    auto drawRow = [&](const std::vector<Item>& row, float left) {
        float x = left;
        for (const auto& it : row) {
            float cw = columnWidth(it);
            float cx = x + cw / 2;
            it.draw(c, cx, pad + iconH / 2);
            c.textCentered(it.label, cx, pad + iconH + labelH / 2, labelPx, kWhite);
            x += cw + colGap;
        }
    };
    drawRow(mouseRow, pad);
    // マウスの項目とキーの項目の間に細い縦線
    float sepX = pad + mouseW2 + sepGap;
    c.stroke([&](float x, float y) { return sdSegment(x, y, sepX, pad + 2 * s, sepX, height - pad - 2 * s); }, 1.0f * s,
             sepX, pad, sepX, height - pad, Color{1, 1, 1, 0.18f});
    drawRow(keyRow, sepX + sepGap);

    // 閉じる記号「×」（右端の上）
    float bx = width - pad - closeSize / 2, by = pad + closeSize / 2;
    auto circle = [&](float x, float y) { return std::hypot(x - bx, y - by) - closeSize / 2; };
    c.stroke(circle, 1.2f * s, bx - closeSize, by - closeSize, bx + closeSize, by + closeSize, kDim);
    float k = closeSize * 0.2f;
    c.stroke([&](float x, float y) { return sdSegment(x, y, bx - k, by - k, bx + k, by + k); }, 1.4f * s, bx - k, by - k,
             bx + k, by + k, kWhite);
    c.stroke([&](float x, float y) { return sdSegment(x, y, bx - k, by + k, bx + k, by - k); }, 1.4f * s, bx - k, by - k,
             bx + k, by + k, kWhite);

    HelpImage img = finish(c);
    img.toggleX = static_cast<int>(bx - closeSize / 2 - 4 * s);
    img.toggleY = static_cast<int>(by - closeSize / 2 - 4 * s);
    img.toggleW = static_cast<int>(closeSize + 8 * s);
    img.toggleH = static_cast<int>(closeSize + 8 * s);
    return img;
}

namespace {

// ボタンの絵: カメラ（画像を保存）
void cameraIcon(Canvas& c, float cx, float cy, float s, const Color& col) {
    auto body = [&](float x, float y) { return sdRoundBox(x, y, cx, cy + 1.5f * s, 10 * s, 7 * s, 2.5f * s); };
    c.stroke(body, 1.5f * s, cx - 11 * s, cy - 6 * s, cx + 11 * s, cy + 9 * s, col);
    auto top = [&](float x, float y) { return sdRoundBox(x, y, cx - 3 * s, cy - 6.5f * s, 4 * s, 1.8f * s, 1 * s); };
    c.fill(top, cx - 8 * s, cy - 9 * s, cx + 2 * s, cy - 4 * s, col);
    auto lens = [&](float x, float y) { return std::hypot(x - cx, y - (cy + 1.5f * s)) - 3.8f * s; };
    c.stroke(lens, 1.5f * s, cx - 5 * s, cy - 3 * s, cx + 5 * s, cy + 6 * s, col);
}

// ボタンの絵: 地球儀（360°）
void globeIcon(Canvas& c, float cx, float cy, float s, const Color& col) {
    float r = 8.5f * s;
    auto outer = [&](float x, float y) { return std::hypot(x - cx, y - cy) - r; };
    c.stroke(outer, 1.5f * s, cx - r, cy - r, cx + r, cy + r, col);
    // 縦の楕円（経線）と横の線（赤道）
    auto meridian = [&](float x, float y) {
        float dx = (x - cx) / (r * 0.45f), dy = (y - cy) / r;
        return (std::sqrt(dx * dx + dy * dy) - 1.0f) * r * 0.45f;
    };
    c.stroke(meridian, 1.2f * s, cx - r, cy - r, cx + r, cy + r, col);
    c.stroke([&](float x, float y) { return sdSegment(x, y, cx - r, cy, cx + r, cy); }, 1.2f * s, cx - r, cy, cx + r, cy, col);
}

// ボタンの絵: 画面と矢印（ビューアを書き出す）
void viewerIcon(Canvas& c, float cx, float cy, float s, const Color& col) {
    auto screen = [&](float x, float y) { return sdRoundBox(x, y, cx, cy - 1.5f * s, 10 * s, 7 * s, 1.5f * s); };
    c.stroke(screen, 1.5f * s, cx - 11 * s, cy - 9 * s, cx + 11 * s, cy + 6 * s, col);
    c.stroke([&](float x, float y) { return sdSegment(x, y, cx - 4 * s, cy + 8.5f * s, cx + 4 * s, cy + 8.5f * s); }, 1.5f * s,
             cx - 4 * s, cy + 8 * s, cx + 4 * s, cy + 9 * s, col);
    // 再生の三角（中で動かせる）
    float v[6] = {cx - 2.5f * s, cy - 5 * s, cx - 2.5f * s, cy + 2 * s, cx + 3.5f * s, cy - 1.5f * s};
    c.fill([&](float x, float y) { return sdTriangle(x, y, v); }, cx - 4 * s, cy - 6 * s, cx + 5 * s, cy + 3 * s, col);
}

// ボタンの絵: 2 本の回る矢印（同期）
void syncIcon(Canvas& c, float cx, float cy, float s, const Color& col) {
    float r = 7.5f * s;
    for (int k = 0; k < 2; ++k) {
        float a0 = k == 0 ? -2.8f : 0.34f, a1 = a0 + 2.3f;  // 上半分・下半分の円弧
        auto arc = [&](float x, float y) {
            float ang = std::atan2(y - cy, x - cx);
            float d = ang - a0;
            while (d < 0) d += 6.2831853f;
            if (d <= a1 - a0) return std::fabs(std::hypot(x - cx, y - cy) - r);
            float e0x = cx + std::cos(a0) * r, e0y = cy + std::sin(a0) * r;
            float e1x = cx + std::cos(a1) * r, e1y = cy + std::sin(a1) * r;
            return std::min(std::hypot(x - e0x, y - e0y), std::hypot(x - e1x, y - e1y));
        };
        c.fill([&](float x, float y) { return arc(x, y) - 0.75f * s; }, cx - r - 2 * s, cy - r - 2 * s, cx + r + 2 * s,
               cy + r + 2 * s, col);
        float ex = cx + std::cos(a1) * r, ey = cy + std::sin(a1) * r;
        float tx = -std::sin(a1), ty = std::cos(a1), head = 3.4f * s;
        float v[6] = {ex + tx * head, ey + ty * head, ex - ty * head * 0.8f, ey + tx * head * 0.8f,
                      ex + ty * head * 0.8f, ey - tx * head * 0.8f};
        c.fill([&](float x, float y) { return sdTriangle(x, y, v); }, ex - head * 2, ey - head * 2, ex + head * 2,
               ey + head * 2, col);
    }
}

}  // namespace

ToolbarImage buildToolbarImage(float s, int hover, int sync, int minWidth) {
    std::vector<const char*> labels = {"ExportImage", "ExportPanorama", "ExportViewer"};
    if (sync >= 0) labels.push_back(sync > 0 ? "SyncOn" : "SyncOff");
    // 書き出しのボタンには、保存されるファイルの種類を小さな札で出す（どれが exe・html か分かるように）
    const char* kinds[] = {"PNG", "HTML", "EXE", nullptr};
    const int count = static_cast<int>(labels.size());
    const float pad = 6 * s, iconH = 24 * s, labelPx = 11 * s, labelH = 14 * s, gap = 2 * s;
    const float badgePx = 10 * s, badgeH = 15 * s;
    std::vector<float> widths;
    float total = pad;
    for (int i = 0; i < count; ++i) {
        float w = std::max(40 * s, textWidth(tr(labels[i]), labelPx) + 14 * s);
        if (kinds[i]) w = std::max(w, textWidth(kinds[i], badgePx) + 22 * s);
        widths.push_back(w);
        total += w + gap;
    }
    total += pad - gap;
    if (minWidth > total) {
        // 下の「表示」の設定と幅をそろえるため、余った幅をボタンに等しく配る
        float extra = (minWidth - total) / count;
        for (float& w : widths) w += extra;
        total = static_cast<float>(minWidth);
    }
    int width = static_cast<int>(std::ceil(total));
    int height = static_cast<int>(std::ceil(pad * 2 + iconH + labelH + badgeH + 6 * s));
    Canvas c(width, height);
    auto panel = [&](float x, float y) { return sdRoundBox(x, y, width / 2.0f, height / 2.0f, width / 2.0f, height / 2.0f, 8 * s); };
    c.fill(panel, 0, 0, static_cast<float>(width), static_cast<float>(height), kPanel);

    ToolbarImage tb;
    float x = pad;
    for (int i = 0; i < count; ++i) {
        float w = widths[i];
        if (i == kButtonSync) {
            // 同期のボタンは、ほかのボタン（書き出し）と細い線で区切る
            float sx = x - gap / 2 - 1 * s;
            c.stroke([&](float px, float py) { return sdSegment(px, py, sx, pad + 4 * s, sx, height - pad - 4 * s); }, 1.0f * s,
                     sx, pad, sx, height - pad, Color{1, 1, 1, 0.18f});
            x += 3 * s;
        }
        float bx0 = x, by0 = pad, bx1 = x + w, by1 = height - pad;
        if (i == hover) {
            auto hl = [&](float px, float py) {
                return sdRoundBox(px, py, (bx0 + bx1) / 2, (by0 + by1) / 2, w / 2, (by1 - by0) / 2, 5 * s);
            };
            c.fill(hl, bx0, by0, bx1, by1, Color{1, 1, 1, 0.16f});
        }
        // 同期中は青、止めているときは薄い色
        const Color& col = i == hover || (i == kButtonSync && sync > 0) ? kAccent : i == kButtonSync ? kDim : kWhite;
        float cx = x + w / 2, cy = pad + 2 * s + iconH / 2;
        if (i == kButtonImage) cameraIcon(c, cx, cy, s, col);
        else if (i == kButtonPanorama) globeIcon(c, cx, cy, s, col);
        else if (i == kButtonViewer) viewerIcon(c, cx, cy, s, col);
        else syncIcon(c, cx, cy, s, col);
        c.textCentered(tr(labels[i]), cx, pad + 2 * s + iconH + labelH / 2 + 1 * s, labelPx, i == kButtonSync && sync <= 0 ? kDim : kWhite);
        if (kinds[i]) {
            // ファイルの種類の札（角の丸い枠の中に PNG / HTML / EXE）
            float bw = textWidth(kinds[i], badgePx) + 10 * s, by = pad + 2 * s + iconH + labelH + 2 * s + badgeH / 2;
            auto badge = [&](float px, float py) { return sdRoundBox(px, py, cx, by, bw / 2, badgeH / 2 - 1 * s, 3 * s); };
            c.fill(badge, cx - bw / 2, by - badgeH / 2, cx + bw / 2, by + badgeH / 2, i == hover ? kAccent : Color{1, 1, 1, 0.18f});
            c.textCentered(kinds[i], cx, by, badgePx, kWhite);
        }
        tb.buttons.push_back({static_cast<int>(bx0), static_cast<int>(by0), static_cast<int>(w), static_cast<int>(by1 - by0)});
        x += w + gap;
    }
    tb.image = finish(c);
    return tb;
}

ToolbarImage buildLookPanelImage(float s, int hover, const LookState& look, int minWidth) {
    const float pad = 8 * s, rowH = 26 * s, labelPx = 12 * s, box = 14 * s, swatchW = 34 * s, gap = 10 * s;
    // 1 行 = チェックボックス（＋色の見本を行の右端に）。色の見本は、その色が使われるときだけはっきり出す
    struct Row {
        const char* label;
        bool checked;
        int checkId;
        int colorId;  // 色の見本が無い行は -1
        uint32_t color;
        bool colorActive;
    };
    const Row rows[] = {
        {tr("LookSky"), look.sky, kLookSky, kLookBackground, look.background, !look.sky},
        {tr("LookGround"), look.ground, kLookGround, kLookGroundColor, look.groundColor, look.ground},
        {tr("LookMono"), look.mono, kLookMono, kLookMonoColor, look.monoColor, look.mono},
        {tr("LookTransparent"), look.transparentPng, kLookTransparent, -1, 0, false},
    };
    const int rowCount = 4;
    // 時間帯: 1 行目に「時間帯」と今の時刻（Revit の設定のままなら、その旨）、2 行目につまみ
    std::string timeText;
    if (look.timeHour < 0) {
        timeText = tr("LookTimeRevit");
    } else {
        int minutes = static_cast<int>(std::lround(look.timeHour * 60.0f));
        char buf[16];
        snprintf(buf, sizeof(buf), "%d:%02d", minutes / 60, minutes % 60);
        timeText = buf;
    }
    float labelW = textWidth(tr("LookTime"), labelPx) + 12 * s + textWidth(tr("LookTimeRevit"), labelPx);
    for (const Row& r : rows) labelW = std::max(labelW, textWidth(r.label, labelPx));
    float colorW = swatchW + 8 * s;
    int width = static_cast<int>(std::ceil(std::max<float>(pad * 2 + box + 6 * s + labelW + 8 * s + gap + colorW, minWidth)));
    float checkW = width - pad * 2 - gap - colorW;
    int height = static_cast<int>(std::ceil(pad * 2 + rowH * (rowCount + 2) + 4 * s));
    Canvas c(width, height);
    auto panel = [&](float x, float y) { return sdRoundBox(x, y, width / 2.0f, height / 2.0f, width / 2.0f, height / 2.0f, 8 * s); };
    c.fill(panel, 0, 0, static_cast<float>(width), static_cast<float>(height), kPanel);

    ToolbarImage tb;
    tb.buttons.resize(kLookCount);
    auto toColor = [](uint32_t bgr) {
        return Color{(bgr & 255) / 255.0f, ((bgr >> 8) & 255) / 255.0f, ((bgr >> 16) & 255) / 255.0f, 1.0f};
    };
    auto highlight = [&](float x0, float y0, float w, float h) {
        c.fill([&](float px, float py) { return sdRoundBox(px, py, x0 + w / 2, y0 + h / 2, w / 2, h / 2 - 1 * s, 5 * s); },
               x0, y0, x0 + w, y0 + h, Color{1, 1, 1, 0.16f});
    };
    for (int row = 0; row < rowCount; ++row) {
        const Row& r = rows[row];
        float y0 = pad + rowH * row, cy = y0 + rowH / 2;

        // チェックボックスと言葉（言葉の上を押しても切り替わる）
        float x0 = pad;
        if (hover == r.checkId) highlight(x0, y0, checkW, rowH);
        float bx = x0 + 4 * s + box / 2;
        auto boxSd = [&](float px, float py) { return sdRoundBox(px, py, bx, cy, box / 2, box / 2, 3 * s); };
        if (r.checked) {
            c.fill(boxSd, bx - box, cy - box, bx + box, cy + box, kAccent);
            c.stroke([&](float px, float py) {
                return std::min(sdSegment(px, py, bx - box * 0.28f, cy, bx - box * 0.06f, cy + box * 0.24f),
                                sdSegment(px, py, bx - box * 0.06f, cy + box * 0.24f, bx + box * 0.3f, cy - box * 0.22f));
            }, 2.0f * s, bx - box, cy - box, bx + box, cy + box, kWhite);
        } else {
            c.stroke(boxSd, 1.5f * s, bx - box, cy - box, bx + box, cy + box, kWhite);
        }
        float tx = bx + box / 2 + 6 * s;
        c.textCentered(r.label, tx + textWidth(r.label, labelPx) / 2, cy, labelPx, kWhite);
        tb.buttons[r.checkId] = {static_cast<int>(x0), static_cast<int>(y0), static_cast<int>(checkW), static_cast<int>(rowH)};
        if (r.colorId < 0) continue;

        // 色の見本（押すと色を選ぶ画面）。何の色かは、カーソルを乗せたときの説明で出す
        float cx0 = width - pad - colorW;
        if (hover == r.colorId) highlight(cx0, y0, colorW, rowH);
        float sx = cx0 + colorW / 2;
        Color sw = toColor(r.color);
        if (!r.colorActive) sw.a = 0.35f;
        auto swSd = [&](float px, float py) { return sdRoundBox(px, py, sx, cy, swatchW / 2, box / 2, 3 * s); };
        c.fill(swSd, sx - swatchW, cy - box, sx + swatchW, cy + box, sw);
        c.stroke(swSd, 1.0f * s, sx - swatchW, cy - box, sx + swatchW, cy + box, Color{1, 1, 1, r.colorActive ? 0.8f : 0.35f});
        tb.buttons[r.colorId] = {static_cast<int>(cx0), static_cast<int>(y0), static_cast<int>(colorW), static_cast<int>(rowH)};
    }

    // --- 時間帯 ---------------------------------------------------------------
    float ty0 = pad + rowH * rowCount + 4 * s;
    c.stroke([&](float px, float py) { return sdSegment(px, py, pad + 4 * s, ty0 - 2 * s, width - pad - 4 * s, ty0 - 2 * s); },
             1.0f * s, pad, ty0 - 3 * s, static_cast<float>(width) - pad, ty0 - 1 * s, Color{1, 1, 1, 0.18f});
    // 1 行目: 「時間帯」と時刻。押すと Revit の設定に戻す
    float lineW = width - pad * 2;
    if (hover == kLookTimeReset) highlight(pad, ty0, lineW, rowH);
    float tcy = ty0 + rowH / 2;
    c.textCentered(tr("LookTime"), pad + 4 * s + textWidth(tr("LookTime"), labelPx) / 2, tcy, labelPx, kWhite);
    float valueW = textWidth(timeText, labelPx);
    c.textCentered(timeText, width - pad - 4 * s - valueW / 2, tcy, labelPx, look.timeHour < 0 ? kDim : kAccent);
    tb.buttons[kLookTimeReset] = {static_cast<int>(pad), static_cast<int>(ty0), static_cast<int>(lineW), static_cast<int>(rowH)};
    // 2 行目: つまみ。目盛りは 6・9・12・15・18 時
    float sy0 = ty0 + rowH, scy = sy0 + rowH / 2;
    float trackX0 = pad + 10 * s, trackX1 = width - pad - 10 * s;
    auto xOfHour = [&](float h) { return trackX0 + (trackX1 - trackX0) * (h - kTimeFirstHour) / (kTimeLastHour - kTimeFirstHour); };
    for (int h = 6; h <= 18; h += 3) {
        float x = xOfHour(static_cast<float>(h));
        c.stroke([&](float px, float py) { return sdSegment(px, py, x, scy + 5 * s, x, scy + 8 * s); }, 1.0f * s, x - 2 * s,
                 scy + 4 * s, x + 2 * s, scy + 9 * s, Color{1, 1, 1, 0.45f});
    }
    auto track = [&](float px, float py) { return sdRoundBox(px, py, (trackX0 + trackX1) / 2, scy, (trackX1 - trackX0) / 2, 2 * s, 2 * s); };
    c.fill(track, trackX0 - 2 * s, scy - 3 * s, trackX1 + 2 * s, scy + 3 * s, Color{1, 1, 1, 0.35f});
    float hour = look.timeHour < 0 ? look.revitHour : look.timeHour;
    float kx = xOfHour(std::min(std::max(hour, kTimeFirstHour), kTimeLastHour));
    float kr = (hover == kLookTime ? 7.0f : 6.0f) * s;
    c.fill([&](float px, float py) { return std::sqrt((px - kx) * (px - kx) + (py - scy) * (py - scy)) - kr; }, kx - kr - 1, scy - kr - 1,
           kx + kr + 1, scy + kr + 1, look.timeHour < 0 ? Color{0.85f, 0.85f, 0.85f, 1.0f} : kAccent);
    // つまみの当たり判定は、目盛りの端から端まで（行の高さいっぱい）
    tb.buttons[kLookTime] = {static_cast<int>(trackX0), static_cast<int>(sy0), static_cast<int>(trackX1 - trackX0), static_cast<int>(rowH)};
    tb.image = finish(c);
    return tb;
}

HelpImage buildStatusImage(const std::vector<std::string>& lines, float progress, float s) {
    const float pad = 10 * s, linePx = 13 * s, lineH = 19 * s, barH = 6 * s, barGap = 8 * s;
    float textW = 0;
    for (const auto& l : lines) textW = std::max(textW, textWidth(l, linePx));
    bool bar = progress >= 0;
    int width = static_cast<int>(std::ceil(std::max(textW, bar ? 260 * s : 0.0f) + pad * 2 + 8 * s));
    int height = static_cast<int>(std::ceil(pad * 2 + lineH * lines.size() + (bar ? barGap + barH : 0)));
    Canvas c(width, height);
    auto panel = [&](float x, float y) { return sdRoundBox(x, y, width / 2.0f, height / 2.0f, width / 2.0f, height / 2.0f, 8 * s); };
    c.fill(panel, 0, 0, static_cast<float>(width), static_cast<float>(height), Color{0, 0, 0, 0.7f});
    float y = pad;
    for (const auto& l : lines) {
        c.textCentered(l, width / 2.0f, y + lineH / 2, linePx, kWhite);
        y += lineH;
    }
    if (bar) {
        float x0 = pad + 4 * s, x1 = width - pad - 4 * s, by = y + barGap + barH / 2;
        auto track = [&](float px, float py) { return sdRoundBox(px, py, (x0 + x1) / 2, by, (x1 - x0) / 2, barH / 2, barH / 2); };
        c.fill(track, x0, by - barH, x1, by + barH, Color{1, 1, 1, 0.18f});
        float xe = x0 + (x1 - x0) * std::clamp(progress, 0.0f, 1.0f);
        if (xe > x0 + barH * 0.5f) {
            auto fillBar = [&](float px, float py) { return sdRoundBox(px, py, (x0 + xe) / 2, by, (xe - x0) / 2, barH / 2, barH / 2); };
            c.fill(fillBar, x0, by - barH, xe, by + barH, kAccent);
        }
    }
    return finish(c);
}

}  // namespace t28

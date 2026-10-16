#include "text_overlay.h"

#include <algorithm>
#include <cmath>
#include <cstdlib>

#include "i18n.h"
#include "log.h"

#define STB_TRUETYPE_IMPLEMENTATION
#include <stb_truetype.h>

namespace t28 {

namespace {

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

}  // namespace

bool renderTextPanel(const std::vector<std::string>& lines, float pixelHeight, TextImage& out) {
    static std::vector<unsigned char> fontData;  // フォントは一度だけ読む
    static stbtt_fontinfo font;
    static bool fontReady = false;
    if (!fontReady) {
        for (const std::string& path : fontCandidates()) {
            if (!loadFile(path, fontData)) continue;
            int offset = stbtt_GetFontOffsetForIndex(fontData.data(), 0);
            if (offset >= 0 && stbtt_InitFont(&font, fontData.data(), offset)) {
                logInfo("Help font: %s", path.c_str());
                fontReady = true;
                break;
            }
        }
        if (!fontReady) {
            logError("No font found for the help text");
            return false;
        }
    }

    const float scale = stbtt_ScaleForPixelHeight(&font, pixelHeight);
    int ascent, descent, lineGap;
    stbtt_GetFontVMetrics(&font, &ascent, &descent, &lineGap);
    const int lineHeight = static_cast<int>(std::ceil((ascent - descent + lineGap) * scale));
    const int baseline = static_cast<int>(std::round(ascent * scale));
    const int pad = static_cast<int>(std::round(pixelHeight * 0.6f));

    std::vector<std::vector<int>> text;
    int maxWidth = 0;
    for (const std::string& line : lines) {
        text.push_back(decodeUtf8(line));
        float w = 0;
        const auto& cps = text.back();
        for (size_t i = 0; i < cps.size(); ++i) {
            int adv, lsb;
            stbtt_GetCodepointHMetrics(&font, cps[i], &adv, &lsb);
            w += adv * scale;
            if (i + 1 < cps.size()) w += stbtt_GetCodepointKernAdvance(&font, cps[i], cps[i + 1]) * scale;
        }
        maxWidth = std::max(maxWidth, static_cast<int>(std::ceil(w)));
    }

    const int width = maxWidth + pad * 2;
    const int height = lineHeight * static_cast<int>(lines.size()) + pad * 2;
    std::vector<float> coverage(static_cast<size_t>(width) * height, 0.0f);

    for (size_t li = 0; li < text.size(); ++li) {
        float x = static_cast<float>(pad);
        int y = pad + static_cast<int>(li) * lineHeight + baseline;
        const auto& cps = text[li];
        for (size_t i = 0; i < cps.size(); ++i) {
            int adv, lsb;
            stbtt_GetCodepointHMetrics(&font, cps[i], &adv, &lsb);
            int x0, y0, x1, y1;
            float shift = x - std::floor(x);
            stbtt_GetCodepointBitmapBoxSubpixel(&font, cps[i], scale, scale, shift, 0, &x0, &y0, &x1, &y1);
            int gw = x1 - x0, gh = y1 - y0;
            if (gw > 0 && gh > 0) {
                std::vector<unsigned char> glyph(static_cast<size_t>(gw) * gh);
                stbtt_MakeCodepointBitmapSubpixel(&font, glyph.data(), gw, gh, gw, scale, scale, shift, 0, cps[i]);
                int gx = static_cast<int>(std::floor(x)) + x0, gy = y + y0;
                for (int r = 0; r < gh; ++r) {
                    for (int c = 0; c < gw; ++c) {
                        int px = gx + c, py = gy + r;
                        if (px < 0 || py < 0 || px >= width || py >= height) continue;
                        float& cov = coverage[static_cast<size_t>(py) * width + px];
                        cov = std::max(cov, glyph[static_cast<size_t>(r) * gw + c] / 255.0f);
                    }
                }
            }
            x += adv * scale;
            if (i + 1 < cps.size()) x += stbtt_GetCodepointKernAdvance(&font, cps[i], cps[i + 1]) * scale;
        }
    }

    // 黒い半透明の背景の上に白い文字。重ねる側（overlay.comp）は mix(元の色, rgb, a) で合成するので、
    // 「背景 → 文字」の順に重ねた結果になるよう rgb と a を求める
    const float bgAlpha = 0.55f, textAlpha = 0.95f;
    const float radius = pixelHeight * 0.5f;
    out.width = static_cast<uint32_t>(width);
    out.height = static_cast<uint32_t>(height);
    out.rgba.assign(static_cast<size_t>(width) * height * 4, 0);
    for (int py = 0; py < height; ++py) {
        for (int px = 0; px < width; ++px) {
            // 角を丸くする
            float dx = std::max(0.0f, std::max(radius - px - 0.5f, px + 0.5f - (width - radius)));
            float dy = std::max(0.0f, std::max(radius - py - 0.5f, py + 0.5f - (height - radius)));
            float inside = std::clamp(radius - std::sqrt(dx * dx + dy * dy) + 0.5f, 0.0f, 1.0f);
            float bg = bgAlpha * inside;
            float t = coverage[static_cast<size_t>(py) * width + px] * textAlpha;
            float a = 1.0f - (1.0f - bg) * (1.0f - t);
            float rgb = a > 0 ? t / a : 0.0f;
            uint8_t* p = &out.rgba[(static_cast<size_t>(py) * width + px) * 4];
            p[0] = p[1] = p[2] = static_cast<uint8_t>(std::round(rgb * 255));
            p[3] = static_cast<uint8_t>(std::round(a * 255));
        }
    }
    return true;
}

}  // namespace t28

#include "i18n.h"

#include <cstring>

namespace t28 {

static Lang g_lang = Lang::JP;

struct Entry {
    const char* key;
    const char* jp;
    const char* us;
    const char* cn;
};

static const Entry kStrings[] = {
    {"NoVulkan",
     "Vulkan（GPU を動かす仕組み）を使えません。GPU のドライバーを最新にしてください。",
     "Vulkan is not available. Please update your GPU driver.",
     "无法使用 Vulkan。请将显卡驱动更新到最新版本。"},
    {"NoGpu",
     "使える GPU が見つかりません。GPU のドライバーを最新にしてください。\n"
     "（リモートデスクトップや仮想マシンでは GPU を使えないことがあります）",
     "No usable GPU was found. Please update your GPU driver.\n"
     "(The GPU may not be available over Remote Desktop or in a virtual machine.)",
     "未找到可用的 GPU。请将显卡驱动更新到最新版本。\n"
     "（通过远程桌面或在虚拟机中可能无法使用 GPU）"},
    {"SceneLoadFailed",
     "モデルのデータを読み込めませんでした。",
     "Could not load the model data.",
     "无法读取模型数据。"},
    {"WindowTitle",
     "Tools28 リアルタイムレンダリング",
     "Tools28 Real-time Rendering",
     "Tools28 实时渲染"},
    // 操作の案内（ウィンドウの左下に小さく出す）。絵の下に付ける短い言葉
    {"HelpZoom", "ズーム", "Zoom", "缩放"},
    {"HelpPan", "移動", "Pan", "平移"},
    {"HelpOrbit", "回転", "Orbit", "动态观察"},
    {"HelpFit", "全体表示", "Zoom to fit", "缩放匹配"},
    {"HelpLook", "見回す", "Look around", "环视"},
    {"HelpWalk", "歩く", "Walk", "行走"},
    {"HelpUpDown", "下がる・上がる", "Down / up", "下降・上升"},
    {"HelpBright", "明るさ", "Brightness", "亮度"},
    {"HelpReset", "最初の視点", "Reset view", "初始视点"},
};

void setLanguage(const char* code) {
    if (!code) return;
    if (strcmp(code, "US") == 0 || strcmp(code, "EN") == 0) g_lang = Lang::US;
    else if (strcmp(code, "CN") == 0) g_lang = Lang::CN;
    else g_lang = Lang::JP;
}

Lang currentLanguage() { return g_lang; }

const char* tr(const char* key) {
    for (const Entry& e : kStrings) {
        if (strcmp(e.key, key) == 0) {
            switch (g_lang) {
                case Lang::US: return e.us;
                case Lang::CN: return e.cn;
                default: return e.jp;
            }
        }
    }
    return key;
}

}  // namespace t28

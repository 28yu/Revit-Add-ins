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
    // 操作の案内（ウィンドウの左下に小さく出す）。Revit の 3D ビューと同じ操作を先に書く
    {"Help1",
     "ホイール: ズーム　　ホイールドラッグ: 移動　　Shift＋ホイールドラッグ: 回転",
     "Wheel: zoom    Wheel drag: pan    Shift + wheel drag: orbit",
     "滚轮: 缩放　　按住滚轮拖动: 平移　　Shift＋按住滚轮拖动: 动态观察"},
    {"Help2",
     "ホイールをダブルクリック: 全体表示　　左ドラッグ: 見回す　　W A S D・Q E: 歩く・上下",
     "Double-click wheel: zoom to fit    Left drag: look around    W A S D / Q E: walk / up, down",
     "双击滚轮: 缩放匹配　　左键拖动: 环视　　W A S D・Q E: 行走・上下"},
    {"Help3",
     "＋ －: 明るさ　　R: 最初の視点　　H: この案内を隠す",
     "+ / -: brightness    R: reset view    H: hide this help",
     "＋ －: 亮度　　R: 初始视点　　H: 隐藏此提示"},
    {"HelpHidden",
     "H: 操作の案内",
     "H: show controls",
     "H: 操作提示"},
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

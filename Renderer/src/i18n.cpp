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
    {"NoRayTracingGpu",
     "レイトレーシングに対応した GPU が見つかりません。\n"
     "NVIDIA RTX 20 シリーズ以降・AMD RX 6000 シリーズ以降・Intel Arc が必要です。\n"
     "対応 GPU がある場合は、GPU のドライバーを最新にしてください。",
     "No GPU with ray tracing support was found.\n"
     "An NVIDIA RTX 20 series or later, AMD RX 6000 series or later, or Intel Arc GPU is required.\n"
     "If you have one, please update its driver.",
     "未找到支持光线追踪的 GPU。\n"
     "需要 NVIDIA RTX 20 系列及以上、AMD RX 6000 系列及以上或 Intel Arc。\n"
     "如已安装此类 GPU，请将驱动更新到最新版本。"},
    {"SceneLoadFailed",
     "モデルのデータを読み込めませんでした。",
     "Could not load the model data.",
     "无法读取模型数据。"},
    {"WindowTitle",
     "Tools28 リアルタイムレンダリング",
     "Tools28 Real-time Rendering",
     "Tools28 实时渲染"},
    {"Controls",
     "左ドラッグ: 見回す / W A S D: 移動 / Q E: 上下 / ホイール: 移動の速さ / + -: 明るさ / F: 自動露出",
     "Left drag: look / W A S D: move / Q E: down, up / Wheel: speed / + -: brightness / F: auto exposure",
     "左键拖动: 环视 / W A S D: 移动 / Q E: 下、上 / 滚轮: 移动速度 / + -: 亮度 / F: 自动曝光"},
};

void setLanguage(const char* code) {
    if (!code) return;
    if (strcmp(code, "US") == 0 || strcmp(code, "EN") == 0) g_lang = Lang::US;
    else if (strcmp(code, "CN") == 0) g_lang = Lang::CN;
    else g_lang = Lang::JP;
}

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

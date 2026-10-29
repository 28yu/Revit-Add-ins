// ウィンドウで表示して、マウスとキーボードで歩き回るモード。
#include <algorithm>
#include <chrono>
#include <functional>
#include <cmath>
#include <cstdio>
#include <filesystem>
#include <fstream>
#include <limits>
#include <string>
#include <thread>
#include <atomic>
#include <vector>

#include "app.h"
#include "export.h"
#include "i18n.h"
#include "log.h"
#include "help_overlay.h"

#define GLFW_INCLUDE_NONE
#include <GLFW/glfw3.h>
#ifdef _WIN32
#define GLFW_EXPOSE_NATIVE_WIN32
#include <GLFW/glfw3native.h>
#endif

namespace t28 {

namespace {

// 画面に重ねる絵の置き場所（Renderer の overlay の番号）
const uint32_t kHelpSlot = 0, kToolbarSlot = 1, kStatusSlot = 2, kLookSlot = 3;
// 「表示」の設定のボタンの番号（ツールバーのボタンと区別するため 100 から）
const int kLookIdBase = 100;

struct Swapchain {
    VkSwapchainKHR handle = VK_NULL_HANDLE;
    VkFormat format = VK_FORMAT_UNDEFINED;
    VkExtent2D extent{};
    std::vector<VkImage> images;
    std::vector<VkSemaphore> renderDone;  // 画像ごとに1つ（表示が終わるまで使い回さない）
};

void createSwapchain(VkContext& ctx, GLFWwindow* window, Swapchain& sc) {
    VkSurfaceCapabilitiesKHR caps;
    vkGetPhysicalDeviceSurfaceCapabilitiesKHR(ctx.physical, ctx.surface, &caps);

    uint32_t formatCount = 0;
    vkGetPhysicalDeviceSurfaceFormatsKHR(ctx.physical, ctx.surface, &formatCount, nullptr);
    std::vector<VkSurfaceFormatKHR> formats(formatCount);
    vkGetPhysicalDeviceSurfaceFormatsKHR(ctx.physical, ctx.surface, &formatCount, formats.data());
    // 色の補正はシェーダーで済ませているので、補正なし（UNORM）の形式を選ぶ
    VkSurfaceFormatKHR chosen = formats[0];
    for (const auto& f : formats)
        if (f.format == VK_FORMAT_B8G8R8A8_UNORM || f.format == VK_FORMAT_R8G8B8A8_UNORM) { chosen = f; break; }

    int fbw = 0, fbh = 0;
    glfwGetFramebufferSize(window, &fbw, &fbh);
    VkExtent2D extent = caps.currentExtent;
    if (extent.width == UINT32_MAX) extent = {static_cast<uint32_t>(fbw), static_cast<uint32_t>(fbh)};

    uint32_t imageCount = std::max(caps.minImageCount + 1, 2u);
    if (caps.maxImageCount > 0) imageCount = std::min(imageCount, caps.maxImageCount);

    VkSwapchainCreateInfoKHR ci{VK_STRUCTURE_TYPE_SWAPCHAIN_CREATE_INFO_KHR};
    ci.surface = ctx.surface;
    ci.minImageCount = imageCount;
    ci.imageFormat = chosen.format;
    ci.imageColorSpace = chosen.colorSpace;
    ci.imageExtent = extent;
    ci.imageArrayLayers = 1;
    ci.imageUsage = VK_IMAGE_USAGE_TRANSFER_DST_BIT;
    ci.imageSharingMode = VK_SHARING_MODE_EXCLUSIVE;
    ci.preTransform = caps.currentTransform;
    ci.compositeAlpha = VK_COMPOSITE_ALPHA_OPAQUE_BIT_KHR;
    ci.presentMode = VK_PRESENT_MODE_FIFO_KHR;  // 画面の更新に合わせる（ちらつき防止）
    ci.clipped = VK_TRUE;
    ci.oldSwapchain = sc.handle;
    VkSwapchainKHR newHandle;
    vkCheck(vkCreateSwapchainKHR(ctx.device, &ci, nullptr, &newHandle), "vkCreateSwapchainKHR");
    if (sc.handle) vkDestroySwapchainKHR(ctx.device, sc.handle, nullptr);
    for (VkSemaphore s : sc.renderDone) vkDestroySemaphore(ctx.device, s, nullptr);

    sc.handle = newHandle;
    sc.format = chosen.format;
    sc.extent = extent;
    uint32_t count = 0;
    vkGetSwapchainImagesKHR(ctx.device, sc.handle, &count, nullptr);
    sc.images.resize(count);
    vkGetSwapchainImagesKHR(ctx.device, sc.handle, &count, sc.images.data());
    sc.renderDone.resize(count);
    VkSemaphoreCreateInfo sci{VK_STRUCTURE_TYPE_SEMAPHORE_CREATE_INFO};
    for (auto& s : sc.renderDone) vkCheck(vkCreateSemaphore(ctx.device, &sci, nullptr, &s), "vkCreateSemaphore");
    logInfo("Swapchain %ux%u, %u images, format %d", extent.width, extent.height, count, static_cast<int>(chosen.format));
}

// 点 p を、点 center を通る軸 axis（単位ベクトル）のまわりに angle（ラジアン）回す
Vec3 rotateAround(Vec3 p, Vec3 center, Vec3 axis, float angle) {
    Vec3 v{p.x - center.x, p.y - center.y, p.z - center.z};
    float c = std::cos(angle), s = std::sin(angle);
    float d = axis.x * v.x + axis.y * v.y + axis.z * v.z;
    Vec3 cr{axis.y * v.z - axis.z * v.y, axis.z * v.x - axis.x * v.z, axis.x * v.y - axis.y * v.x};
    Vec3 r{v.x * c + cr.x * s + axis.x * d * (1 - c), v.y * c + cr.y * s + axis.y * d * (1 - c),
           v.z * c + cr.z * s + axis.z * d * (1 - c)};
    return {center.x + r.x, center.y + r.y, center.z + r.z};
}

Vec3 add(Vec3 a, Vec3 b, float k = 1.0f) { return {a.x + b.x * k, a.y + b.y * k, a.z + b.z * k}; }
float length(Vec3 a) { return std::sqrt(a.x * a.x + a.y * a.y + a.z * a.z); }

// マウスの操作。Revit の 3D ビューと同じ割り当てにしている
//   ホイール: ズーム / ホイールドラッグ: パン / Shift + ホイールドラッグ: オービット / ホイールのダブルクリック: 全体表示
//   左ドラッグ: 見回す（その場で向きを変える）
enum class Drag { None, Look, Pan, Orbit };

struct Navigator {
    GLFWwindow* window = nullptr;
    CameraController* cam = nullptr;
    CameraController initial;
    const Scene* scene = nullptr;
    const Bvh* bvh = nullptr;

    Drag drag = Drag::None;
    double lastX = 0, lastY = 0;
    Vec3 pivot;           // オービットの中心（ドラッグを始めた位置にある物の表面）
    float panDepth = 1;   // パンの基準の奥行き（ドラッグを始めた位置にある物までの距離）
    double lastMiddleClick = -1;

    float evOffset = 0;
    bool autoExposure = true;
    bool showHelp = true;
    bool changed = false;     // カメラや設定が変わった
    bool helpChanged = false;
    // 案内の切り替えの記号の位置（画面の画素）。クリックで開く / 閉じる
    int toggleX = 0, toggleY = 0, toggleW = 0, toggleH = 0;

    // 右上のツールバー（書き出しのボタン）。位置は画面の画素
    std::vector<ToolbarImage::Rect> buttons;
    std::vector<int> buttonIds;  // buttons と同じ並び。ツールバーは kButton*、「表示」の設定は kLookIdBase + kLook*
    int hoverButton = -1;
    // 時間帯のつまみ: ドラッグ中か、まだ反映していない位置（画面の画素）があるか、離したか
    bool sliderDrag = false, sliderPending = false, sliderReleased = false;
    double sliderX = 0;
    bool toolbarChanged = false;
    int clickedButton = -1;   // 押されたボタン（描画の繰り返しの中で処理する）
    bool exporting = false;   // 書き出し中（カメラの操作を止め、Esc で中止）
    bool cancelExport = false;

    // ウィンドウの座標 → 画面の画素（高 DPI の画面では違う）
    void toFramebuffer(double mx, double my, double& fx, double& fy) const {
        int ww = 1, wh = 1, fw = 1, fh = 1;
        glfwGetWindowSize(window, &ww, &wh);
        glfwGetFramebufferSize(window, &fw, &fh);
        fx = mx * fw / std::max(1, ww);
        fy = my * fh / std::max(1, wh);
    }
    int buttonAt(double mx, double my) const {
        double fx, fy;
        toFramebuffer(mx, my, fx, fy);
        for (size_t i = 0; i < buttons.size(); ++i) {
            const auto& b = buttons[i];
            if (fx >= b.x && fx < b.x + b.w && fy >= b.y && fy < b.y + b.h)
                return i < buttonIds.size() ? buttonIds[i] : static_cast<int>(i);
        }
        return -1;
    }

    Vec3 sceneCenter() const {
        Vec3 mn = bvh->boundsMin(), mx = bvh->boundsMax();
        return {(mn.x + mx.x) / 2, (mn.y + mx.y) / 2, (mn.z + mx.z) / 2};
    }
    float sceneRadius() const { return std::max(1.0f, length(add(bvh->boundsMax(), bvh->boundsMin(), -1.0f)) / 2); }

    // カーソルの位置を通る光線と、その先にある物までの距離（何も無ければモデルの中心までの距離）
    Vec3 cursorRay(double mx, double my, float* distance) const {
        int ww = 1, wh = 1;
        glfwGetWindowSize(window, &ww, &wh);
        float ndcX = static_cast<float>(mx / std::max(1, ww) * 2 - 1);
        float ndcY = static_cast<float>(1 - my / std::max(1, wh) * 2);
        Vec3 dir = cam->rayThrough(ndcX, ndcY, static_cast<float>(ww) / std::max(1, wh));
        float d = bvh->intersect(*scene, cam->eye, dir);
        if (d <= 0) d = std::max(0.5f, length(add(sceneCenter(), cam->eye, -1.0f)));
        *distance = d;
        return dir;
    }

    void zoomToFit() {
        // 今の向きのまま、モデル全体が画面に入る位置まで下がる
        float r = sceneRadius();
        int ww = 1, wh = 1;
        glfwGetWindowSize(window, &ww, &wh);
        float halfFov = cam->fovY * 0.5f * 3.14159265f / 180.0f;
        float halfFovX = std::atan(std::tan(halfFov) * ww / std::max(1, wh));
        float dist = r / std::sin(std::min(halfFov, halfFovX));
        Vec3 f = cam->forward();
        cam->eye = add(sceneCenter(), f, -dist);
        changed = true;
    }

    void onScroll(double dy) {
        if (exporting) return;
        double mx, my;
        glfwGetCursorPos(window, &mx, &my);
        float d;
        Vec3 dir = cursorRay(mx, my, &d);
        // 1 目盛りで、カーソルの先の物まで 2 割近づく（離れるときは同じ割合で戻る）
        float step = static_cast<float>(dy) > 0 ? d * 0.2f : -d * 0.25f;
        cam->eye = add(cam->eye, dir, step * static_cast<float>(std::fabs(dy)));
        changed = true;
    }

    void onMouseButton(int button, int action, int mods) {
        double mx, my;
        glfwGetCursorPos(window, &mx, &my);
        lastX = mx;
        lastY = my;
        if (action == GLFW_RELEASE) {
            drag = Drag::None;
            if (sliderDrag) {
                sliderDrag = false;
                sliderReleased = true;
                sliderPending = true;
            }
            return;
        }
        if (exporting) return;
        if (button == GLFW_MOUSE_BUTTON_LEFT) {
            int b = buttonAt(mx, my);
            if (b == kLookIdBase + kLookTime) {
                // 時間帯のつまみ: 押した位置に動かし、そのままドラッグで動かせる
                double fy;
                toFramebuffer(mx, my, sliderX, fy);
                sliderDrag = true;
                sliderPending = true;
                drag = Drag::None;
                return;
            }
            if (b >= 0) {
                clickedButton = b;
                drag = Drag::None;
                return;
            }
            // 案内の記号（× / ?）の上なら、案内を開く / 閉じる（画面の画素とウィンドウの座標の比をそろえて比べる）
            double fx, fy;
            toFramebuffer(mx, my, fx, fy);
            if (fx >= toggleX && fx < toggleX + toggleW && fy >= toggleY && fy < toggleY + toggleH) {
                showHelp = !showHelp;
                helpChanged = true;
                drag = Drag::None;
                return;
            }
            drag = Drag::Look;
        } else if (button == GLFW_MOUSE_BUTTON_MIDDLE) {
            double now = glfwGetTime();
            if (lastMiddleClick >= 0 && now - lastMiddleClick < 0.35) {
                zoomToFit();
                lastMiddleClick = -1;
                drag = Drag::None;
                return;
            }
            lastMiddleClick = now;
            float d;
            Vec3 dir = cursorRay(mx, my, &d);
            pivot = add(cam->eye, dir, d);
            Vec3 f = cam->forward();
            panDepth = std::max(0.1f, d * (dir.x * f.x + dir.y * f.y + dir.z * f.z));
            drag = (mods & GLFW_MOD_SHIFT) ? Drag::Orbit : Drag::Pan;
        }
    }

    void onCursor(double mx, double my) {
        float dx = static_cast<float>(mx - lastX), dy = static_cast<float>(my - lastY);
        lastX = mx;
        lastY = my;
        if (sliderDrag) {
            double fy;
            toFramebuffer(mx, my, sliderX, fy);
            sliderPending = true;
            return;
        }
        int hover = drag == Drag::None && !exporting ? buttonAt(mx, my) : -1;
        if (hover != hoverButton) {
            hoverButton = hover;
            toolbarChanged = true;
        }
        if (exporting) return;
        if (drag == Drag::None || (dx == 0 && dy == 0)) return;
        int ww = 1, wh = 1;
        glfwGetWindowSize(window, &ww, &wh);

        if (drag == Drag::Look) {
            float sens = 0.0035f * cam->fovY / 60.0f;
            cam->yaw += dx * sens;
            cam->pitch = std::clamp(cam->pitch - dy * sens, -1.55f, 1.55f);
        } else if (drag == Drag::Pan) {
            // カーソルの下の物が、カーソルに付いて動くように平行移動する
            float k = panDepth * 2.0f * std::tan(cam->fovY * 0.5f * 3.14159265f / 180.0f) / std::max(1, wh);
            cam->eye = add(add(cam->eye, cam->right(), -dx * k), cam->up(), dy * k);
        } else if (drag == Drag::Orbit) {
            // 中心のまわりにカメラを回す（左右: 上下の軸まわり、上下: カメラの横の軸まわり）
            const float sens = 0.005f;
            float yawAngle = -dx * sens;
            Vec3 zAxis{0, 0, 1};
            cam->eye = rotateAround(cam->eye, pivot, zAxis, yawAngle);
            cam->yaw -= yawAngle;

            float pitchAngle = -dy * sens;
            Vec3 right = cam->right();
            Vec3 f = rotateAround(cam->forward(), {0, 0, 0}, right, pitchAngle);
            if (std::fabs(f.z) < 0.995f) {  // 真上・真下を越えて裏返らないようにする
                cam->eye = rotateAround(cam->eye, pivot, right, pitchAngle);
                cam->setForward(f);
            }
        }
        changed = true;
    }

    void onKey(int key) {
        if (exporting) {
            if (key == GLFW_KEY_ESCAPE) cancelExport = true;
            return;
        }
        switch (key) {
            case GLFW_KEY_ESCAPE: glfwSetWindowShouldClose(window, GLFW_TRUE); break;
            case GLFW_KEY_P: clickedButton = kButtonImage; break;
            case GLFW_KEY_EQUAL: case GLFW_KEY_KP_ADD: evOffset += 0.5f; changed = true; break;
            case GLFW_KEY_MINUS: case GLFW_KEY_KP_SUBTRACT: evOffset -= 0.5f; changed = true; break;
            case GLFW_KEY_F: autoExposure = !autoExposure; changed = true; break;
            case GLFW_KEY_R: *cam = initial; changed = true; break;
            case GLFW_KEY_H: showHelp = !showHelp; helpChanged = true; break;
            default: break;
        }
    }
};

Navigator* nav(GLFWwindow* w) { return static_cast<Navigator*>(glfwGetWindowUserPointer(w)); }

}  // namespace

int runWindow(const Options& opt, Scene& scene) {
    if (!glfwInit()) fatal("glfwInit failed");
    glfwWindowHint(GLFW_CLIENT_API, GLFW_NO_API);
    std::string windowTitle = opt.title.empty() ? tr("WindowTitle") : opt.title + " - " + tr("WindowTitle");
    GLFWwindow* window = glfwCreateWindow(opt.width, opt.height, windowTitle.c_str(), nullptr, nullptr);
    if (!window) fatal("glfwCreateWindow failed");

    uint32_t extCount = 0;
    const char** exts = glfwGetRequiredInstanceExtensions(&extCount);
    if (!exts) fatal("%s", tr("NoVulkan"));
    std::vector<const char*> instanceExts(exts, exts + extCount);

    VkContext ctx;
    ctx.init(instanceExts, [&](VkInstance inst) {
        VkSurfaceKHR surface;
        vkCheck(glfwCreateWindowSurface(inst, window, nullptr, &surface), "glfwCreateWindowSurface");
        return surface;
    }, opt.gpu, !opt.forceSoftware);
    if (opt.gpuMemoryMb > 0) ctx.deviceLocalBytes = static_cast<VkDeviceSize>(opt.gpuMemoryMb) << 20;  // 動作確認用

    Bvh bvh;
    uint32_t glassStart = prepareScene(scene, bvh, opt.instancing);
    // 止まっている間の仕上げ（Open Image Denoise。インストール先の oidn フォルダにあれば使う。ビューアの exe には無い）
    OidnDenoiser oidn;
    if (opt.finalDenoise) oidn.load();
    Renderer renderer;
    renderer.skyWidth_ = static_cast<uint32_t>(opt.skySize);
    renderer.init(ctx, scene, glassStart, bvh);
    renderer.setLocalToneMapping(opt.localToneMapping);
    renderer.setSkySampling(opt.skySampling);
    renderer.setFinalDenoiser(&oidn);

    // RT コアの無い GPU では、半分の大きさで描いて拡大する（速さを保つため）
    float renderScale = opt.renderScale > 0 ? opt.renderScale : (ctx.hardwareRt ? 1.0f : 0.5f);
    // 描く大きさの上限。重いときは自動で下げ、軽くなったらここまで戻す（adaptQuality）
    const float maxRenderScale = renderScale;
    // 最初は小さめに描き、軽ければすぐに戻す（重いモデルで、最初の 1 枚に時間がかかりすぎて GPU がリセットされないように）
    if (opt.renderScale <= 0) renderScale = std::min(renderScale, 0.5f);
    logInfo("Render scale %.2f (max %.2f)", renderScale, maxRenderScale);

    // GPU のメモリに収まる描画の大きさか（足りなければ縮める倍率を返す。1 = そのまま）。
    // 描画用の画像は大きさに比例して増える（1920×1080 で約 400 MB、4K で約 1.6 GB）ので、メモリの少ない GPU で落ちないようにする
    auto memoryFit = [&](uint32_t w, uint32_t h) {
        if (ctx.deviceLocalBytes == 0) return 1.0f;
        const VkDeviceSize others = ctx.allocatedDeviceBytes > renderer.imageBytesNow() ? ctx.allocatedDeviceBytes - renderer.imageBytesNow() : 0;
        const VkDeviceSize cap = ctx.deviceLocalBytes * 85 / 100;
        const VkDeviceSize budget = cap > others ? cap - others : 0;
        const VkDeviceSize need = Renderer::imageBytes(w, h);
        if (need <= budget) return 1.0f;
        return static_cast<float>(std::sqrt(static_cast<double>(budget) / static_cast<double>(need))) * 0.95f;
    };

    Swapchain sc;
    Image display;  // 画面と同じ大きさの画像（描画結果を拡大して写し、案内を重ねてから画面に出す）
    VkImageLayout displayLayout = VK_IMAGE_LAYOUT_UNDEFINED;
    VkExtent2D rs{};
    auto recreate = [&]() {
        vkDeviceWaitIdle(ctx.device);
        createSwapchain(ctx, window, sc);
        rs = {std::max(16u, static_cast<uint32_t>(sc.extent.width * renderScale)),
              std::max(16u, static_cast<uint32_t>(sc.extent.height * renderScale))};
        if (float fit = memoryFit(rs.width, rs.height); fit < 1.0f) {
            renderScale = std::max(0.1f, renderScale * fit);
            rs = {std::max(16u, static_cast<uint32_t>(sc.extent.width * renderScale)),
                  std::max(16u, static_cast<uint32_t>(sc.extent.height * renderScale))};
            logInfo("Render scale lowered to %.2f to fit GPU memory", renderScale);
        }
        renderer.resize(rs.width, rs.height);
        ctx.destroyImage(display);
        display = ctx.createImage(sc.extent.width, sc.extent.height, VK_FORMAT_R8G8B8A8_UNORM,
                                  VK_IMAGE_USAGE_STORAGE_BIT | VK_IMAGE_USAGE_TRANSFER_SRC_BIT | VK_IMAGE_USAGE_TRANSFER_DST_BIT);
        displayLayout = VK_IMAGE_LAYOUT_UNDEFINED;
        renderer.setOverlayTarget(display);
    };
    recreate();

    // --- 重いモデルでの自動調整 ---------------------------------------------
    // Windows は GPU の 1 回の処理が約 2 秒を超えると GPU をリセットし、描画プログラムが落ちる。
    // 1 枚に時間がかかるようになったら描く大きさ（と光の跳ね返りの回数）を下げ、軽くなったら少しずつ戻す
    const float kMinRenderScale = 0.25f;
    uint32_t activeBounces = opt.bounces;
    int slowFrames = 0;
    double fastSeconds = 0;
    bool loweredNotified = false;
    bool everSlow = false;  // 一度でも重かったら、戻すのはゆっくりにする
    std::function<void(const std::vector<std::string>&, double)> notify;  // 後で toast を入れる
    auto applyScale = [&](float scale) {
        renderScale = scale;
        rs = {std::max(16u, static_cast<uint32_t>(sc.extent.width * renderScale)),
              std::max(16u, static_cast<uint32_t>(sc.extent.height * renderScale))};
        renderer.resize(rs.width, rs.height);
    };
    // 重い 1 枚は光の計算を横長の帯に分けて GPU に送る（帯 1 本を短く保ち、Windows の GPU リセットを防ぐ）
    BandGovernor ptBands;
    if (opt.bands > 0) {
        ptBands.bands = static_cast<uint32_t>(opt.bands);
        renderer.setPathTraceBands(ptBands.bands);
    }
    auto updateBands = [&](double gpuTime) {
        if (opt.bands > 0 || !ptBands.update(gpuTime)) return;
        renderer.setPathTraceBands(ptBands.bands);
        logInfo("Frame took %.0f ms: path trace in %u bands", gpuTime * 1000.0, ptBands.bands);
    };
    auto adaptQuality = [&](double gpuTime, float dt) {
        auto lower = [&](float factor) {
            if (renderScale > kMinRenderScale + 1e-3f) {
                applyScale(std::max(kMinRenderScale, renderScale * factor));
            } else if (activeBounces > 1) {
                --activeBounces;
            } else {
                return;
            }
            logInfo("Frame took %.0f ms: render scale %.2f, bounces %u (GPU memory %.0f MB)", gpuTime * 1000.0, renderScale,
                    activeBounces, ctx.allocatedDeviceBytes / 1048576.0);
            if (!loweredNotified && notify) {
                notify({tr("QualityLowered")}, 4.0);
                loweredNotified = true;
            }
            slowFrames = 0;
            fastSeconds = 0;
            everSlow = true;
        };
        if (gpuTime > 1.0) {
            lower(0.5f);  // 危ない: すぐに大きく下げる
        } else if (gpuTime > 0.25) {
            if (++slowFrames >= 3) lower(0.75f);
        } else {
            slowFrames = 0;
        }
        // 続けて軽ければ、1 段ずつ戻す（先に跳ね返りの回数、次に大きさ）。
        // 起動直後で一度も重くなっていなければ 0.5 秒ごと、重くなったことがあれば 5 秒ごと
        fastSeconds = gpuTime < 0.05 ? fastSeconds + dt : 0;
        if (fastSeconds > (everSlow ? 5.0 : 0.5)) {
            fastSeconds = 0;
            if (activeBounces < opt.bounces) {
                ++activeBounces;
                logInfo("Restored bounces %u", activeBounces);
            } else if (renderScale < maxRenderScale - 1e-3f) {
                float next = std::min(maxRenderScale, renderScale / 0.75f);
                // 戻した大きさが GPU のメモリに収まるときだけ戻す
                if (memoryFit(static_cast<uint32_t>(sc.extent.width * next), static_cast<uint32_t>(sc.extent.height * next)) >= 1.0f) {
                    applyScale(next);
                    logInfo("Restored render scale %.2f", renderScale);
                }
            }
        }
    };

    CameraController cam;
    cam.setFrom(scene.camera);
    Navigator nv;
    nv.window = window;
    nv.cam = &cam;
    nv.initial = cam;
    nv.scene = &scene;
    nv.bvh = &bvh;
    nv.evOffset = opt.exposureEv;
    nv.autoExposure = opt.autoExposure;
    glfwSetWindowUserPointer(window, &nv);
    glfwSetScrollCallback(window, [](GLFWwindow* w, double, double dy) { nav(w)->onScroll(dy); });
    glfwSetMouseButtonCallback(window, [](GLFWwindow* w, int b, int a, int m) { nav(w)->onMouseButton(b, a, m); });
    glfwSetCursorPosCallback(window, [](GLFWwindow* w, double x, double y) { nav(w)->onCursor(x, y); });
    glfwSetKeyCallback(window, [](GLFWwindow* w, int key, int, int action, int) {
        if (action == GLFW_PRESS) nav(w)->onKey(key);
    });

    float contentScale = 1.0f;
    glfwGetWindowContentScale(window, &contentScale, nullptr);
    contentScale = std::clamp(contentScale, 1.0f, 3.0f);
    HelpImage help;
    auto updateHelp = [&]() {
        help = buildHelpImage(nv.showHelp, contentScale);
        // ウィンドウが狭くて入りきらないときは、案内を小さくする
        int available = static_cast<int>(sc.extent.width) - static_cast<int>(24 * contentScale);
        if (static_cast<int>(help.width) > available && available > 100)
            help = buildHelpImage(nv.showHelp, contentScale * available / static_cast<float>(help.width));
        renderer.setOverlay(kHelpSlot, help.rgba, help.width, help.height);
    };
    updateHelp();

    AutoExposure exposure;

    // --- ライブ同期 -----------------------------------------------------------
    // Revit 側（Tools28）は、モデルが変わるとシーンファイルを書き直す（視点は伝えない。描画ウィンドウと Revit は別々に動かす）。
    // ここでは 0.15 秒ごとにファイルの更新時刻を調べて読み直す。同期を止めると「.pause」を作り、Revit 側はそれを見て書き出しを休む
    namespace fs = std::filesystem;
    const bool liveSync = opt.liveSync && !opt.scenePath.empty();
    const std::string pausePath = opt.scenePath + ".pause";
    // 更新時刻（ファイルが無ければ kNoFile。時刻の数値は負になることもあるので、負かどうかでは判断しない）
    const long long kNoFile = std::numeric_limits<long long>::min();
    auto stampOf = [&](const std::string& path) -> long long {
        std::error_code ec;
        auto t = fs::last_write_time(fs::u8path(path), ec);
        return ec ? kNoFile : static_cast<long long>(t.time_since_epoch().count());
    };
    bool syncOn = true;
    long long sceneStamp = kNoFile;
    if (liveSync) {
        std::error_code ec;
        fs::remove(fs::u8path(pausePath), ec);
        sceneStamp = stampOf(opt.scenePath);
    }
    double syncTimer = 0;

    // --- 「表示」の設定（空・雲／背景色、単色表示／色）。前回の設定をログと同じフォルダの look.ini に残す ------
    LookState look;
    std::string lookPath;
    if (!opt.logPath.empty()) {
        size_t slash = opt.logPath.find_last_of("/\\");
        lookPath = (slash == std::string::npos ? std::string() : opt.logPath.substr(0, slash + 1)) + "look.ini";
        if (FILE* f = openFile(lookPath, "r")) {
            char line[128];
            while (fgets(line, sizeof(line), f)) {
                unsigned v = 0;
                float hourValue = 0;
                if (sscanf(line, "sky=%u", &v) == 1) look.sky = v != 0;
                else if (sscanf(line, "background=%x", &v) == 1) look.background = v & 0xFFFFFF;
                else if (sscanf(line, "mono=%u", &v) == 1) look.mono = v != 0;
                else if (sscanf(line, "monoColor=%x", &v) == 1) look.monoColor = v & 0xFFFFFF;
                else if (sscanf(line, "ground=%u", &v) == 1) look.ground = v != 0;
                else if (sscanf(line, "transparentPng=%u", &v) == 1) look.transparentPng = v != 0;
                else if (sscanf(line, "groundColor=%x", &v) == 1) look.groundColor = v & 0xFFFFFF;
                else if (sscanf(line, "timeHour=%f", &hourValue) == 1) look.timeHour = hourValue;
            }
            fclose(f);
        }
    }
    auto saveLook = [&]() {
        if (lookPath.empty()) return;
        if (FILE* f = openFile(lookPath, "w")) {
            fprintf(f, "sky=%d\nbackground=%06X\nground=%d\ngroundColor=%06X\nmono=%d\nmonoColor=%06X\ntransparentPng=%d\ntimeHour=%.2f\n",
                    look.sky ? 1 : 0, look.background, look.ground ? 1 : 0, look.groundColor, look.mono ? 1 : 0, look.monoColor,
                    look.transparentPng ? 1 : 0, look.timeHour);
            fclose(f);
        }
    };
    // 3D ビューの「レンダリング設定」を最初の見た目にする（look.ini より優先。そのあとウィンドウで変えればそちらが優先）。
    // 時間帯も、開くたびに Revit の「太陽の設定」から始める
    auto applyRevitLook = [&](const RevitLook& rl) {
        look.sun = rl.sun;
        look.lights = rl.artificial;
        if (rl.background == 0) {
            look.sky = true;
        } else {
            look.sky = false;
            if (rl.background == 1) look.background = rl.backgroundColor;
            else look.transparentPng = true;  // 背景「透明」: PNG の書き出しで背景を透明にする
        }
        if ((rl.groundColor >> 24) != 0) look.groundColor = rl.groundColor & 0xFFFFFF;
        renderer.setCloudAmount(rl.cloudAmount);
    };
    RevitLook lastRevitLook = scene.revitLook;
    applyRevitLook(lastRevitLook);
    look.timeHour = -1.0f;
    logInfo("Revit rendering settings: sun=%d artificial=%d background=%d color=%06X ground=%08X clouds=%.2f",
            lastRevitLook.sun ? 1 : 0, lastRevitLook.artificial ? 1 : 0, lastRevitLook.background, lastRevitLook.backgroundColor,
            lastRevitLook.groundColor, lastRevitLook.cloudAmount);
    auto applyLook = [&]() {
        renderer.setLook(look.sky ? 0u : (0xFF000000u | look.background), look.mono ? (0xFF000000u | look.monoColor) : 0u,
                         !look.ground, 0xFF000000u | look.groundColor);
        renderer.setLighting(look.sun, look.lights);
    };
    applyLook();
    // 時間帯: つまみで選んだ時刻の太陽にする（負なら Revit の「太陽の設定」のまま）
    auto applySun = [&]() {
        look.revitHour = hourOfSun(scene, scene.sunDir);
        if (look.timeHour >= 0) {
            look.timeHour = std::min(std::max(look.timeHour, kTimeFirstHour), kTimeLastHour);
            renderer.setSunDirection(sunDirectionAt(scene, look.timeHour));
        } else {
            renderer.setSunDirection(scene.sunDir);
        }
    };
    applySun();
    double lastSunUpdate = 0;

    ToolbarImage toolbar, lookPanel;
    int timeTrackX = 0, timeTrackW = 0;  // 時間帯のつまみの範囲（画面の画素）
    auto updateToolbar = [&]() {
        int hover = nv.hoverButton;
        // 書き出しのボタンと「表示」の設定は、広いほうに幅をそろえる
        int tbHover = hover < kLookIdBase ? hover : -1, lookHover = hover >= kLookIdBase ? hover - kLookIdBase : -1;
        int sync = liveSync ? (syncOn ? 1 : 0) : -1;
        int w = static_cast<int>(std::max(buildToolbarImage(contentScale, -1, sync, 0, look.transparentPng).image.width,
                                          buildLookPanelImage(contentScale, -1, look).image.width));
        toolbar = buildToolbarImage(contentScale, tbHover, sync, w, look.transparentPng);
        renderer.setOverlay(kToolbarSlot, toolbar.image.rgba, toolbar.image.width, toolbar.image.height);
        lookPanel = buildLookPanelImage(contentScale, lookHover, look, w);
        renderer.setOverlay(kLookSlot, lookPanel.image.rgba, lookPanel.image.width, lookPanel.image.height);
    };
    updateToolbar();

    // 上の中央のお知らせ（ボタンの説明・書き出しの進み具合・保存しました）
    HelpImage status;
    std::vector<std::string> statusLines;
    float statusProgress = -1;
    double toastUntil = 0;  // この時刻まで「保存しました」などを出す
    auto setStatus = [&](const std::vector<std::string>& lines, float progress) {
        if (lines == statusLines && progress == statusProgress) return;
        statusLines = lines;
        statusProgress = progress;
        status = lines.empty() ? HelpImage{} : buildStatusImage(lines, progress, contentScale);
        renderer.setOverlay(kStatusSlot, status.rgba, status.width, status.height);
    };
    auto toast = [&](const std::vector<std::string>& lines, double seconds = 6.0) {
        setStatus(lines, -1);
        toastUntil = glfwGetTime() + seconds;
    };
    notify = [&](const std::vector<std::string>& lines, double seconds) { toast(lines, seconds); };

    // --- 書き出し -------------------------------------------------------------
    // 画像・360°は、ウィンドウの描画を一時的に書き出す大きさに切り替え、決めた回数だけ足し合わせてから保存する
    // （進み具合を画面に出しながら。Esc で中止）。ビューアはシーンを exe に付け足すだけ
    struct ExportJob {
        ExportKind kind = ExportKind::Image;
        std::string path;
        ViewParams view;
        uint32_t width = 0, height = 0;
        int samples = 0, done = 0, lastPercent = -1;
        bool transparent = false;  // 背景を透明にした PNG
        bool finishing = false;    // 足し合わせ終わり、Open Image Denoise で仕上げている
    } job;
    VkPhysicalDeviceMemoryProperties memProps;
    vkGetPhysicalDeviceMemoryProperties(ctx.physical, &memProps);
    VkDeviceSize gpuMemory = 0;
    for (uint32_t i = 0; i < memProps.memoryHeapCount; ++i)
        if (memProps.memoryHeaps[i].flags & VK_MEMORY_HEAP_DEVICE_LOCAL_BIT)
            gpuMemory = std::max(gpuMemory, memProps.memoryHeaps[i].size);
    void* nativeWindow = nullptr;
#ifdef _WIN32
    nativeWindow = glfwGetWin32Window(window);
#endif
    const char* tipKeys[kButtonCount] = {"TipImage", "TipPanorama", "TipViewer", "TipSync"};
    // ビューアの書き出しは、シーンを詰める（圧縮する）のに時間がかかるので別のスレッドで行う（その間ウィンドウは動かせる）。
    // 書き出し中はライブ同期でシーンを入れ替えない（別のスレッドが読んでいるため）
    std::thread viewerThread;
    std::atomic<int> viewerState{0};  // 0 = していない, 1 = 書き出し中, 2 = 終わった
    bool viewerOk = false;
    std::string viewerPath, viewerError;

    auto startExport = [&](int button) {
        ExportKind kind = button == kButtonImage ? ExportKind::Image
                        : button == kButtonPanorama ? ExportKind::Panorama : ExportKind::Viewer;
        if (viewerState != 0) return;  // ビューアの書き出しが終わるまで待つ
        // 画像は JPEG（背景を透明にするときだけ PNG）
        std::string path = chooseSavePath(nativeWindow, kind, defaultExportName(opt.viewName, kind, look.transparentPng));
        if (path.empty()) return;
        if (kind == ExportKind::Viewer) {
            // 今の視点を最初の視点にして、シーンを詰めて exe に付け足す
            Camera start = scene.camera;
            start.eye = cam.eye;
            start.forward = cam.forward();
            start.up = cam.up();
            start.fovY = cam.fovY;
            viewerPath = path;
            viewerState = 1;
            logInfo("Viewer export start: %s", path.c_str());
            viewerThread = std::thread([&, start, path]() {
                auto t = std::chrono::steady_clock::now();
                viewerOk = writeViewerExe(path, scene, start, opt.title, viewerError);
                logInfo("Viewer export took %.1f s", std::chrono::duration<double>(std::chrono::steady_clock::now() - t).count());
                viewerState = 2;
            });
            return;
        }
        job = ExportJob{};
        job.kind = kind;
        job.path = path;
        job.view = cam.view();
        job.view.maxBounces = opt.bounces;
        job.view.exposure = nv.autoExposure ? exposure.current : std::exp2(nv.evOffset);
        job.view.samplesPerFrame = ctx.hardwareRt ? 4 : 1;
        if (kind == ExportKind::Panorama) {
            // パノラマは水平に構える（見上げ・見下ろしは見る人が操作する）
            CameraController flat = cam;
            flat.pitch = 0;
            ViewParams fv = flat.view();
            job.view.forward = fv.forward;
            job.view.right = fv.right;
            job.view.up = fv.up;
            job.view.panorama = true;
            // 大きいほどきれいだが GPU のメモリを多く使う（4096×2048 で約 1.7 GB）
            bool large = ctx.hardwareRt && gpuMemory >= (5ull << 30);
            job.width = large ? 4096 : 2048;
            job.height = job.width / 2;
            // 最後に Open Image Denoise で仕上げるときは、足し合わせる回数を半分にする（ノイズは仕上げで取れるので、書き出しが速くなる）
            job.samples = (ctx.hardwareRt ? 256 : 64) / (renderer.finalDenoiserAvailable() ? 2 : 1);
        } else {
            // 画面と同じ縦横比で、長い辺が 1920 画素以上（最大 3840）
            uint32_t longSide = std::max(sc.extent.width, sc.extent.height);
            float k = std::clamp(1920.0f / longSide, 1.0f, 3840.0f / longSide);
            job.width = static_cast<uint32_t>(sc.extent.width * k);
            job.height = static_cast<uint32_t>(sc.extent.height * k);
            job.samples = (ctx.hardwareRt ? 512 : 128) / (renderer.finalDenoiserAvailable() ? 2 : 1);
        }
        if (opt.exportSamples > 0) job.samples = opt.exportSamples;
        // GPU のメモリに収まらない大きさなら、縦横比を保って小さくする（足りないまま書き出すと落ちる）
        if (float fit = memoryFit(job.width, job.height); fit < 1.0f) {
            uint32_t ow = job.width, oh = job.height;
            job.width = std::max(256u, static_cast<uint32_t>(job.width * fit) / 2 * 2);
            job.height = kind == ExportKind::Panorama ? job.width / 2 : std::max(144u, static_cast<uint32_t>(job.height * fit));
            logInfo("Export size reduced to fit GPU memory: %ux%u -> %ux%u (GPU memory %.0f MB of %.0f MB)", ow, oh, job.width,
                    job.height, ctx.allocatedDeviceBytes / 1048576.0, ctx.deviceLocalBytes / 1048576.0);
            toast({tr("ExportReduced"), std::to_string(job.width) + " x " + std::to_string(job.height)});
        }
        // PNG の書き出しで「背景を透明に」なら、空（と消した地面）の所を透明にして描く
        job.transparent = kind == ExportKind::Image && look.transparentPng;
        renderer.setTransparentOutput(job.transparent);
        // 書き出す大きさから、光の計算を帯に分ける数を決めておく（4K なら 4 本。かかった時間でさらに直す）
        ptBands = BandGovernor{};
        ptBands.bands = opt.bands > 0 ? static_cast<uint32_t>(opt.bands) : BandGovernor::initialFor(job.width, job.height);
        renderer.setPathTraceBands(ptBands.bands);
        logInfo("Export start: %s %ux%u, %d samples%s, %u bands", path.c_str(), job.width, job.height, job.samples,
                job.transparent ? ", transparent background" : "", ptBands.bands);
        renderer.resize(job.width, job.height);
        // 途中では仕上げない（最後に 1 回だけ。書き出す画像は最後の結果で仕上げる）
        renderer.setAutoFinalDenoise(false);
        nv.exporting = true;
        nv.cancelExport = false;
        nv.hoverButton = -1;
        nv.drag = Drag::None;
    };
    auto finishExport = [&](bool cancelled) {
        if (!cancelled) {
            std::vector<uint8_t> rgba = readImageRgba(ctx, renderer.output(), job.transparent);
            bool ok = job.kind == ExportKind::Image ? writeImage(job.path, rgba, job.width, job.height)
                                                    : writePanoramaHtml(job.path, rgba, job.width, job.height, opt.title);
            logInfo("Export %s: %s", ok ? "done" : "FAILED", job.path.c_str());
            if (!ok) toast({std::string(tr("ExportFailed")) + fileNameOf(job.path)});
            else if (job.kind == ExportKind::Panorama)
                toast({std::string(tr("Saved")) + fileNameOf(job.path), tr("SavedPanoramaNote")});
            else toast({std::string(tr("Saved")) + fileNameOf(job.path)});
        } else {
            logInfo("Export cancelled");
            toast({tr("ExportCancelled")});
        }
        nv.exporting = false;
        renderer.setAutoFinalDenoise(true);
        renderer.setTransparentOutput(false);
        ptBands = BandGovernor{};
        if (opt.bands > 0) ptBands.bands = static_cast<uint32_t>(opt.bands);
        renderer.setPathTraceBands(ptBands.bands);
        nv.cancelExport = false;
        renderer.resize(rs.width, rs.height);
        nv.toolbarChanged = true;
    };

    VkCommandBufferAllocateInfo ai{VK_STRUCTURE_TYPE_COMMAND_BUFFER_ALLOCATE_INFO};
    ai.commandPool = ctx.commandPool;
    ai.level = VK_COMMAND_BUFFER_LEVEL_PRIMARY;
    ai.commandBufferCount = 1;
    VkCommandBuffer cmd;
    vkCheck(vkAllocateCommandBuffers(ctx.device, &ai, &cmd), "vkAllocateCommandBuffers");
    VkSemaphoreCreateInfo sci{VK_STRUCTURE_TYPE_SEMAPHORE_CREATE_INFO};
    VkSemaphore imageReady;
    vkCheck(vkCreateSemaphore(ctx.device, &sci, nullptr, &imageReady), "vkCreateSemaphore");
    VkFenceCreateInfo fci{VK_STRUCTURE_TYPE_FENCE_CREATE_INFO};
    VkFence fence;
    vkCheck(vkCreateFence(ctx.device, &fci, nullptr, &fence), "vkCreateFence");

    bool needResize = false;
    // 1 フレームで計算する回数。GPU の速さに合わせて 1〜4 回の間で自動で調整する（目安 30 fps）
    uint32_t samplesPerFrame = 1;
    double frameTimeAvg = 0.0, sppTimer = 0.0;
    auto last = std::chrono::steady_clock::now();
    double statsTimer = 0;
    int statsFrames = 0;

    while (!glfwWindowShouldClose(window)) {
        glfwPollEvents();
        auto now = std::chrono::steady_clock::now();
        float dt = std::min(0.1f, std::chrono::duration<float>(now - last).count());
        last = now;

        // --- キーボードで歩く（W A S D / 矢印: 水平に移動、Q E: 上下、Shift: 速く） ----------
        Vec3 f = cam.forward(), r = cam.right();
        Vec3 flat{f.x, f.y, 0};
        float flatLen = std::sqrt(flat.x * flat.x + flat.y * flat.y);
        if (flatLen > 1e-4f) { flat.x /= flatLen; flat.y /= flatLen; }
        float step = 3.0f * dt * (glfwGetKey(window, GLFW_KEY_LEFT_SHIFT) == GLFW_PRESS ? 3.0f : 1.0f);
        auto key = [&](int k) { return glfwGetKey(window, k) == GLFW_PRESS; };
        Vec3 move{0, 0, 0};
        if (key(GLFW_KEY_W) || key(GLFW_KEY_UP)) move = add(move, flat);
        if (key(GLFW_KEY_S) || key(GLFW_KEY_DOWN)) move = add(move, flat, -1);
        if (key(GLFW_KEY_D) || key(GLFW_KEY_RIGHT)) move = add(move, r);
        if (key(GLFW_KEY_A) || key(GLFW_KEY_LEFT)) move = add(move, r, -1);
        if (key(GLFW_KEY_E)) move.z += 1;
        if (key(GLFW_KEY_Q)) move.z -= 1;
        if ((move.x != 0 || move.y != 0 || move.z != 0) && !nv.exporting) {
            cam.eye = add(cam.eye, move, step);
            nv.changed = true;
        }
        // カメラが動いたかどうかは Renderer が前のフレームと比べて判断する（動いても履歴は捨てずに重ね直す）
        nv.changed = false;
        if (nv.helpChanged) {
            updateHelp();
            nv.helpChanged = false;
        }
        if (nv.clickedButton == kButtonSync) {
            nv.clickedButton = -1;
            syncOn = !syncOn;
            std::error_code ec;
            if (syncOn) fs::remove(fs::u8path(pausePath), ec);
            else std::ofstream(fs::u8path(pausePath)) << "paused\n";
            logInfo("Live sync %s", syncOn ? "on" : "off");
            toast({tr(syncOn ? "SyncResumed" : "SyncPaused")}, 3.0);
            syncTimer = 1.0;  // すぐに調べ直す
            updateToolbar();
        } else if (nv.clickedButton >= kLookIdBase) {
            int b = nv.clickedButton - kLookIdBase;
            nv.clickedButton = -1;
            if (b == kLookSky) look.sky = !look.sky;
            else if (b == kLookMono) look.mono = !look.mono;
            else if (b == kLookGround) look.ground = !look.ground;
            else if (b == kLookTransparent) look.transparentPng = !look.transparentPng;
            else if (b == kLookSun) look.sun = !look.sun;
            else if (b == kLookLights) look.lights = !look.lights;
            else if (b == kLookTimeReset) {
                look.timeHour = -1;  // Revit の「太陽の設定」に戻す
                applySun();
            } else {
                uint32_t& color = b == kLookBackground ? look.background : b == kLookGroundColor ? look.groundColor : look.monoColor;
                if (chooseColor(nativeWindow, color)) {
                    // 色を選んだら、その色が使われるように切り替えも入れる
                    if (b == kLookBackground) look.sky = false;
                    else if (b == kLookGroundColor) look.ground = true;
                    else look.mono = true;
                }
                last = std::chrono::steady_clock::now();  // 色を選ぶ画面で止まっていた時間は数えない
            }
            logInfo("Look: sky=%d background=%06X ground=%d groundColor=%06X sun=%d lights=%d mono=%d color=%06X transparentPng=%d time=%.2f",
                    look.sky ? 1 : 0, look.background, look.ground ? 1 : 0, look.groundColor, look.sun ? 1 : 0, look.lights ? 1 : 0,
                    look.mono ? 1 : 0, look.monoColor, look.transparentPng ? 1 : 0, look.timeHour);
            applyLook();
            saveLook();
            updateToolbar();
        } else if (nv.clickedButton >= 0) {
            int b = nv.clickedButton;
            nv.clickedButton = -1;
            startExport(b);
            last = std::chrono::steady_clock::now();  // 保存先を選ぶ画面で止まっていた時間は数えない
        }

        // Revit からの変更を受け取る（書き出し中・同期を止めている間は受け取らない）
        syncTimer += dt;
        if (viewerState == 2) {
            viewerThread.join();
            viewerState = 0;
            if (viewerOk) {
                logInfo("Viewer exported: %s", viewerPath.c_str());
                toast({std::string(tr("Saved")) + fileNameOf(viewerPath), tr("SavedViewerNote")});
            } else {
                logError("Viewer export failed: %s", viewerError.c_str());
                toast({std::string(tr("ExportFailed")) + fileNameOf(viewerPath)});
            }
        }
        if (liveSync && syncOn && !nv.exporting && viewerState == 0 && syncTimer >= 0.15) {
            syncTimer = 0;
            long long st = stampOf(opt.scenePath);
            if (st != sceneStamp && st != kNoFile) {
                Scene next;
                std::string error;
                // Revit が書いている途中で読めなかったときは、次の機会にもう一度読む
                if (loadScene(opt.scenePath, next, error)) {
                    Bvh nextBvh;
                    uint32_t nextGlass = prepareScene(next, nextBvh, opt.instancing);
                    renderer.reloadScene(next, nextGlass, nextBvh);
                    scene.sunDir = next.sunDir;  // 時間帯を Revit のままにしているときは新しい太陽を使う
                    scene.latitude = next.latitude;
                    scene.trueNorth = next.trueNorth;
                    scene.dayOfYear = next.dayOfYear;
                    applySun();
                    // Revit でレンダリング設定を変えたときは、その項目を新しい設定に合わせる
                    if (next.revitLook != lastRevitLook) {
                        lastRevitLook = next.revitLook;
                        applyRevitLook(lastRevitLook);
                        applyLook();
                        nv.toolbarChanged = true;
                        logInfo("Live sync: Revit rendering settings changed");
                    }
                    scene = std::move(next);
                    bvh = std::move(nextBvh);
                    sceneStamp = st;
                    logInfo("Live sync: scene reloaded (%zu triangles)", scene.triangleCount());
                    toast({tr("SyncUpdated")}, 2.5);
                } else {
                    logInfo("Live sync: could not read the scene yet (%s)", error.c_str());
                }
            }
        }
        // 時間帯のつまみ: 空を描き直すのに時間がかかるので、ドラッグ中は 0.15 秒に 1 回だけ反映する
        if (nv.sliderPending && (nv.sliderReleased || glfwGetTime() - lastSunUpdate > 0.15) && timeTrackW > 0) {
            float f = static_cast<float>((nv.sliderX - timeTrackX) / timeTrackW);
            float hour = kTimeFirstHour + std::min(std::max(f, 0.0f), 1.0f) * (kTimeLastHour - kTimeFirstHour);
            hour = std::round(hour * 4.0f) / 4.0f;  // 15 分刻み
            if (hour != look.timeHour) {
                look.timeHour = hour;
                applySun();
                updateToolbar();
            }
            if (nv.sliderReleased) {
                logInfo("Time of day: %.2f h", look.timeHour);
                saveLook();
            }
            nv.sliderPending = nv.sliderReleased = false;
            lastSunUpdate = glfwGetTime();
        }
        if (nv.toolbarChanged) {
            updateToolbar();
            nv.toolbarChanged = false;
        }
        // お知らせ: 書き出し中は進み具合、ボタンの上ではその説明、保存した後はしばらく結果
        if (viewerState == 1) {
            setStatus({tr("ExportingViewer")}, -1);
        } else if (nv.exporting && job.finishing) {
            setStatus({tr("ExportFinishing")}, 1.0f);
        } else if (nv.exporting) {
            int percent = job.samples > 0 ? job.done * 100 / job.samples : 0;
            if (percent != job.lastPercent) {
                job.lastPercent = percent;
                const char* key = job.kind == ExportKind::Image ? "ExportingImage" : "ExportingPanorama";
                setStatus({std::string(tr(key)) + "  " + std::to_string(percent) + "%", tr("ExportCancelHint")},
                          percent / 100.0f);
            }
        } else if (glfwGetTime() < toastUntil) {
            // 「保存しました」などはボタンの説明より優先して、しばらく出したままにする
        } else if (nv.hoverButton >= kLookIdBase) {
            static const char* lookTips[kLookCount] = {"TipLookSky",    "TipLookBackground",  "TipLookMono",
                                                       "TipLookMonoColor", "TipLookGround", "TipLookTransparent",
                                                       "TipLookGroundColor", "TipLookTimeReset", "TipLookTime",
                                                       "TipLookSun", "TipLookLights"};
            setStatus({tr(lookTips[nv.hoverButton - kLookIdBase])}, -1);
        } else if (nv.hoverButton >= 0) {
            setStatus({tr(tipKeys[nv.hoverButton])}, -1);
        } else {
            setStatus({}, -1);
        }

        // --- 描画 ---------------------------------------------------------
        int fbw = 0, fbh = 0;
        glfwGetFramebufferSize(window, &fbw, &fbh);
        if (fbw == 0 || fbh == 0) {  // 最小化中
            glfwWaitEvents();
            continue;
        }
        if (needResize || static_cast<uint32_t>(fbw) != sc.extent.width || static_cast<uint32_t>(fbh) != sc.extent.height) {
            recreate();
            if (nv.exporting) renderer.resize(job.width, job.height);  // 書き出し中は書き出す大きさのまま
            updateHelp();
            needResize = false;
        }

        uint32_t imageIndex = 0;
        VkResult acq = vkAcquireNextImageKHR(ctx.device, sc.handle, UINT64_MAX, imageReady, VK_NULL_HANDLE, &imageIndex);
        if (acq == VK_ERROR_OUT_OF_DATE_KHR) { needResize = true; continue; }
        if (acq != VK_SUCCESS && acq != VK_SUBOPTIMAL_KHR) vkCheck(acq, "vkAcquireNextImageKHR");

        // 動いている間は 4 画素に 1 つだけ光を追う（速い GPU で 1 フレームに 2 回以上計算できているときは使わない）
        renderer.setSparseWhileMoving(opt.sparse == 1 || (opt.sparse < 0 && samplesPerFrame == 1));
        ViewParams view = cam.view();
        view.maxBounces = activeBounces;
        view.samplesPerFrame = samplesPerFrame;
        view.exposure = nv.autoExposure ? exposure.current : std::exp2(nv.evOffset);
        if (nv.exporting) view = job.view;
        const uint32_t srcW = nv.exporting ? job.width : rs.width, srcH = nv.exporting ? job.height : rs.height;

        VkCommandBufferBeginInfo bi{VK_STRUCTURE_TYPE_COMMAND_BUFFER_BEGIN_INFO};
        bi.flags = VK_COMMAND_BUFFER_USAGE_ONE_TIME_SUBMIT_BIT;
        vkBeginCommandBuffer(cmd, &bi);
        renderer.record(cmd, view);

        // 1. 描いた画像を、画面の大きさの画像に拡大して写す
        transitionImage(cmd, renderer.output().image, VK_IMAGE_LAYOUT_GENERAL, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL,
                        VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_ACCESS_SHADER_WRITE_BIT,
                        VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_READ_BIT);
        transitionImage(cmd, display.image, displayLayout, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL,
                        VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_READ_BIT,
                        VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_WRITE_BIT);
        VkImageBlit blit{};
        blit.srcSubresource = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 0, 1};
        blit.srcOffsets[1] = {static_cast<int32_t>(srcW), static_cast<int32_t>(srcH), 1};
        blit.dstSubresource = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 0, 1};
        blit.dstOffsets[1] = {static_cast<int32_t>(sc.extent.width), static_cast<int32_t>(sc.extent.height), 1};
        if (nv.exporting) {
            // 書き出し中は、書き出す画像の全体が入るように縮めて真ん中に出す（周りは暗い色）
            VkClearColorValue dark{{0.08f, 0.08f, 0.08f, 1.0f}};
            VkImageSubresourceRange range{VK_IMAGE_ASPECT_COLOR_BIT, 0, 1, 0, 1};
            vkCmdClearColorImage(cmd, display.image, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, &dark, 1, &range);
            VkMemoryBarrier mb{VK_STRUCTURE_TYPE_MEMORY_BARRIER};
            mb.srcAccessMask = VK_ACCESS_TRANSFER_WRITE_BIT;
            mb.dstAccessMask = VK_ACCESS_TRANSFER_WRITE_BIT;
            vkCmdPipelineBarrier(cmd, VK_PIPELINE_STAGE_TRANSFER_BIT, VK_PIPELINE_STAGE_TRANSFER_BIT, 0, 1, &mb, 0, nullptr, 0, nullptr);
            float k = std::min(static_cast<float>(sc.extent.width) / srcW, static_cast<float>(sc.extent.height) / srcH);
            int32_t dw = std::max(1, static_cast<int32_t>(srcW * k)), dh = std::max(1, static_cast<int32_t>(srcH * k));
            int32_t ox = (static_cast<int32_t>(sc.extent.width) - dw) / 2, oy = (static_cast<int32_t>(sc.extent.height) - dh) / 2;
            blit.dstOffsets[0] = {ox, oy, 0};
            blit.dstOffsets[1] = {ox + dw, oy + dh, 1};
        }
        vkCmdBlitImage(cmd, renderer.output().image, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL, display.image,
                       VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, 1, &blit, VK_FILTER_LINEAR);

        // 2. 操作の案内を左下に重ねる
        transitionImage(cmd, display.image, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, VK_IMAGE_LAYOUT_GENERAL,
                        VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_WRITE_BIT, VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT,
                        VK_ACCESS_SHADER_READ_BIT | VK_ACCESS_SHADER_WRITE_BIT);
        int margin = static_cast<int>(12 * contentScale);
        int helpX = margin, helpY = static_cast<int>(sc.extent.height) - static_cast<int>(help.height) - margin;
        int barX = static_cast<int>(sc.extent.width) - static_cast<int>(toolbar.image.width) - margin, barY = margin;
        nv.buttons.clear();
        nv.buttonIds.clear();
        // 「表示」の設定はツールバー（書き出しのボタン）のすぐ下。幅はツールバーとそろえてある
        int lookX = barX;
        int lookY = barY + static_cast<int>(toolbar.image.height) + margin / 2;
        // 上の中央のお知らせは、ツールバーと「表示」の設定のうち下に長いほうの下に出す
        const int panelsBottom = std::max(barY + static_cast<int>(toolbar.image.height), lookY + static_cast<int>(lookPanel.image.height));
        if (!nv.exporting) {
            renderer.recordOverlay(cmd, kHelpSlot, helpX, helpY, nv.showHelp ? 0.95f : 0.8f);
            renderer.recordOverlay(cmd, kToolbarSlot, barX, barY, 0.95f);
            renderer.recordOverlay(cmd, kLookSlot, lookX, lookY, 0.95f);
            for (size_t i = 0; i < toolbar.buttons.size(); ++i) {
                const auto& b = toolbar.buttons[i];
                nv.buttons.push_back({barX + b.x, barY + b.y, b.w, b.h});
                nv.buttonIds.push_back(static_cast<int>(i));
            }
            for (size_t i = 0; i < lookPanel.buttons.size(); ++i) {
                const auto& b = lookPanel.buttons[i];
                nv.buttons.push_back({lookX + b.x, lookY + b.y, b.w, b.h});
                nv.buttonIds.push_back(kLookIdBase + static_cast<int>(i));
            }
            timeTrackX = lookX + lookPanel.buttons[kLookTime].x;
            timeTrackW = lookPanel.buttons[kLookTime].w;
        }
        if (status.width > 0) {
            int statusX = (static_cast<int>(sc.extent.width) - static_cast<int>(status.width)) / 2;
            int statusY = nv.exporting ? margin : panelsBottom + margin / 2;
            renderer.recordOverlay(cmd, kStatusSlot, statusX, statusY, 1.0f);
        }
        nv.toggleX = helpX + help.toggleX;
        nv.toggleY = helpY + help.toggleY;
        nv.toggleW = nv.exporting ? 0 : help.toggleW;
        nv.toggleH = help.toggleH;

        // 3. 画面に写す（形式の違い BGRA / RGBA も blit が吸収する）
        VkImage swapImage = sc.images[imageIndex];
        transitionImage(cmd, display.image, VK_IMAGE_LAYOUT_GENERAL, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL,
                        VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_ACCESS_SHADER_WRITE_BIT,
                        VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_READ_BIT);
        displayLayout = VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL;
        transitionImage(cmd, swapImage, VK_IMAGE_LAYOUT_UNDEFINED, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL,
                        VK_PIPELINE_STAGE_TRANSFER_BIT, 0, VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_WRITE_BIT);
        VkImageBlit present{};
        present.srcSubresource = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 0, 1};
        present.srcOffsets[1] = {static_cast<int32_t>(sc.extent.width), static_cast<int32_t>(sc.extent.height), 1};
        present.dstSubresource = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 0, 1};
        present.dstOffsets[1] = present.srcOffsets[1];
        vkCmdBlitImage(cmd, display.image, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL, swapImage,
                       VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, 1, &present, VK_FILTER_NEAREST);
        transitionImage(cmd, swapImage, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, VK_IMAGE_LAYOUT_PRESENT_SRC_KHR,
                        VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_WRITE_BIT,
                        VK_PIPELINE_STAGE_BOTTOM_OF_PIPE_BIT, 0);
        transitionImage(cmd, renderer.output().image, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL, VK_IMAGE_LAYOUT_GENERAL,
                        VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_READ_BIT,
                        VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_ACCESS_SHADER_WRITE_BIT);
        vkEndCommandBuffer(cmd);

        VkPipelineStageFlags waitStage = VK_PIPELINE_STAGE_TRANSFER_BIT;
        VkSubmitInfo si{VK_STRUCTURE_TYPE_SUBMIT_INFO};
        si.waitSemaphoreCount = 1;
        si.pWaitSemaphores = &imageReady;
        si.pWaitDstStageMask = &waitStage;
        si.commandBufferCount = 1;
        si.pCommandBuffers = &cmd;
        si.signalSemaphoreCount = 1;
        si.pSignalSemaphores = &sc.renderDone[imageIndex];
        const auto submitTime = std::chrono::steady_clock::now();
        vkCheck(vkQueueSubmit(ctx.queue, 1, &si, fence), "vkQueueSubmit");

        VkPresentInfoKHR pi{VK_STRUCTURE_TYPE_PRESENT_INFO_KHR};
        pi.waitSemaphoreCount = 1;
        pi.pWaitSemaphores = &sc.renderDone[imageIndex];
        pi.swapchainCount = 1;
        pi.pSwapchains = &sc.handle;
        pi.pImageIndices = &imageIndex;
        VkResult pres = vkQueuePresentKHR(ctx.queue, &pi);
        if (pres == VK_ERROR_OUT_OF_DATE_KHR || pres == VK_SUBOPTIMAL_KHR) needResize = true;
        else if (pres != VK_SUCCESS) vkCheck(pres, "vkQueuePresentKHR");

        // GPU がリセットされたとき（VK_ERROR_DEVICE_LOST）は、ここで分かる言葉で知らせて終わる
        vkCheck(vkWaitForFences(ctx.device, 1, &fence, VK_TRUE, UINT64_MAX), "vkWaitForFences");
        vkResetFences(ctx.device, 1, &fence);
        const double gpuTime = std::chrono::duration<double>(std::chrono::steady_clock::now() - submitTime).count();
        if (nv.exporting) {
            // 書き出し中は明るさを変えない（書き出しを始めたときの明るさのまま）
            renderer.finishFrame();
            updateBands(gpuTime);
            job.done += static_cast<int>(job.view.samplesPerFrame);
            if (nv.cancelExport) {
                renderer.waitFinalDenoise();
                finishExport(true);
            } else if (job.done >= job.samples) {
                // 足し合わせ終わったら、Open Image Denoise で仕上げた絵を 1 枚描いてから保存する
                //   1 枚目: 仕上げを頼む → 2 枚目: 結果を CPU に写して計算（終わるまで待つ）→ 3 枚目: 仕上げた絵を描く → 保存
                if (renderer.finalDenoiseShown() || !renderer.finalDenoiserAvailable()) {
                    finishExport(false);
                } else if (!job.finishing) {
                    job.finishing = true;
                    renderer.requestFinalDenoise();
                } else {
                    renderer.waitFinalDenoise();
                }
            }
            continue;
        }
        // 動いている間はゆっくり、止まった直後に素早く合わせると目が疲れにくい
        exposure.update(renderer.finishFrame(), nv.evOffset, 0.9f);
        adaptQuality(gpuTime, dt);
        updateBands(gpuTime);

        if (opt.exitAfterFrames > 0 && static_cast<int>(renderer.accumulatedSamples()) >= opt.exitAfterFrames) {
            logInfo("Exit after %d frames (exposure x%.2f)", opt.exitAfterFrames, exposure.current);
            glfwSetWindowShouldClose(window, GLFW_TRUE);
        }

        // 4 画素に 1 つだけ光を追ったフレームは軽いので、全部の画素で追ったときの時間に直して数える
        // （そうしないと「軽い → 回数を増やす → 4 画素に 1 つをやめる → 重い → 回数を減らす」をくり返す）
        frameTimeAvg = frameTimeAvg * 0.9 + dt * (renderer.lastFrameSparse() ? 2.5 : 1.0) * 0.1;
        sppTimer += dt;
        if (sppTimer >= 0.5) {
            sppTimer = 0;
            if (frameTimeAvg < 0.020 && samplesPerFrame < 4) samplesPerFrame++;
            else if (frameTimeAvg > 0.040 && samplesPerFrame > 1) samplesPerFrame--;
        }

        // 速さは画面には出さず、ログに残す（不具合の調査用）
        statsTimer += dt;
        statsFrames++;
        if (statsTimer >= 10.0) {
            logInfo("%.1f fps, %u samples/frame, %u frames since the camera stopped", statsFrames / statsTimer,
                    samplesPerFrame, renderer.accumulatedSamples());
            statsTimer = 0;
            statsFrames = 0;
        }
    }

    if (viewerThread.joinable()) viewerThread.join();
    vkDeviceWaitIdle(ctx.device);
    vkDestroyFence(ctx.device, fence, nullptr);
    vkDestroySemaphore(ctx.device, imageReady, nullptr);
    for (VkSemaphore s : sc.renderDone) vkDestroySemaphore(ctx.device, s, nullptr);
    vkDestroySwapchainKHR(ctx.device, sc.handle, nullptr);
    ctx.destroyImage(display);
    renderer.destroy();
    ctx.destroy();
    glfwDestroyWindow(window);
    glfwTerminate();
    logInfo("Tools28Render exit");
    return 0;
}

}  // namespace t28

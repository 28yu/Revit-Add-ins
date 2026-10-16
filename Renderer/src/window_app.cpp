// ウィンドウで表示して、マウスとキーボードで歩き回るモード。
#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <string>
#include <vector>

#include "app.h"
#include "i18n.h"
#include "log.h"
#include "help_overlay.h"

#define GLFW_INCLUDE_NONE
#include <GLFW/glfw3.h>

namespace t28 {

namespace {

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
            return;
        }
        if (button == GLFW_MOUSE_BUTTON_LEFT) {
            // 案内の記号（× / ?）の上なら、案内を開く / 閉じる（画面の画素とウィンドウの座標の比をそろえて比べる）
            int ww = 1, wh = 1, fw = 1, fh = 1;
            glfwGetWindowSize(window, &ww, &wh);
            glfwGetFramebufferSize(window, &fw, &fh);
            double fx = mx * fw / std::max(1, ww), fy = my * fh / std::max(1, wh);
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
        switch (key) {
            case GLFW_KEY_ESCAPE: glfwSetWindowShouldClose(window, GLFW_TRUE); break;
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
    GLFWwindow* window = glfwCreateWindow(opt.width, opt.height, tr("WindowTitle"), nullptr, nullptr);
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

    Bvh bvh;
    uint32_t glassStart = prepareScene(scene, bvh);
    Renderer renderer;
    renderer.init(ctx, scene, glassStart, bvh);

    // RT コアの無い GPU では、半分の大きさで描いて拡大する（速さを保つため）
    float renderScale = opt.renderScale > 0 ? opt.renderScale : (ctx.hardwareRt ? 1.0f : 0.5f);
    logInfo("Render scale %.2f", renderScale);

    Swapchain sc;
    Image display;  // 画面と同じ大きさの画像（描画結果を拡大して写し、案内を重ねてから画面に出す）
    VkImageLayout displayLayout = VK_IMAGE_LAYOUT_UNDEFINED;
    VkExtent2D rs{};
    auto recreate = [&]() {
        vkDeviceWaitIdle(ctx.device);
        createSwapchain(ctx, window, sc);
        rs = {std::max(16u, static_cast<uint32_t>(sc.extent.width * renderScale)),
              std::max(16u, static_cast<uint32_t>(sc.extent.height * renderScale))};
        renderer.resize(rs.width, rs.height);
        ctx.destroyImage(display);
        display = ctx.createImage(sc.extent.width, sc.extent.height, VK_FORMAT_R8G8B8A8_UNORM,
                                  VK_IMAGE_USAGE_STORAGE_BIT | VK_IMAGE_USAGE_TRANSFER_SRC_BIT | VK_IMAGE_USAGE_TRANSFER_DST_BIT);
        displayLayout = VK_IMAGE_LAYOUT_UNDEFINED;
        renderer.setOverlayTarget(display);
    };
    recreate();

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
        renderer.setOverlay(help.rgba, help.width, help.height);
    };
    updateHelp();

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

    AutoExposure exposure;
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
        if (move.x != 0 || move.y != 0 || move.z != 0) {
            cam.eye = add(cam.eye, move, step);
            nv.changed = true;
        }
        // カメラが動いたかどうかは Renderer が前のフレームと比べて判断する（動いても履歴は捨てずに重ね直す）
        nv.changed = false;
        if (nv.helpChanged) {
            updateHelp();
            nv.helpChanged = false;
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
            updateHelp();
            needResize = false;
        }

        uint32_t imageIndex = 0;
        VkResult acq = vkAcquireNextImageKHR(ctx.device, sc.handle, UINT64_MAX, imageReady, VK_NULL_HANDLE, &imageIndex);
        if (acq == VK_ERROR_OUT_OF_DATE_KHR) { needResize = true; continue; }
        if (acq != VK_SUCCESS && acq != VK_SUBOPTIMAL_KHR) vkCheck(acq, "vkAcquireNextImageKHR");

        ViewParams view = cam.view();
        view.maxBounces = opt.bounces;
        view.samplesPerFrame = samplesPerFrame;
        view.exposure = nv.autoExposure ? exposure.current : std::exp2(nv.evOffset);

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
        blit.srcOffsets[1] = {static_cast<int32_t>(rs.width), static_cast<int32_t>(rs.height), 1};
        blit.dstSubresource = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 0, 1};
        blit.dstOffsets[1] = {static_cast<int32_t>(sc.extent.width), static_cast<int32_t>(sc.extent.height), 1};
        vkCmdBlitImage(cmd, renderer.output().image, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL, display.image,
                       VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, 1, &blit, VK_FILTER_LINEAR);

        // 2. 操作の案内を左下に重ねる
        transitionImage(cmd, display.image, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, VK_IMAGE_LAYOUT_GENERAL,
                        VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_WRITE_BIT, VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT,
                        VK_ACCESS_SHADER_READ_BIT | VK_ACCESS_SHADER_WRITE_BIT);
        int margin = static_cast<int>(12 * contentScale);
        int helpX = margin, helpY = static_cast<int>(sc.extent.height) - static_cast<int>(help.height) - margin;
        renderer.recordOverlay(cmd, helpX, helpY, nv.showHelp ? 0.95f : 0.8f);
        nv.toggleX = helpX + help.toggleX;
        nv.toggleY = helpY + help.toggleY;
        nv.toggleW = help.toggleW;
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

        vkWaitForFences(ctx.device, 1, &fence, VK_TRUE, UINT64_MAX);
        vkResetFences(ctx.device, 1, &fence);
        // 動いている間はゆっくり、止まった直後に素早く合わせると目が疲れにくい
        exposure.update(renderer.finishFrame(), nv.evOffset, 0.9f);

        if (opt.exitAfterFrames > 0 && static_cast<int>(renderer.accumulatedSamples()) >= opt.exitAfterFrames) {
            logInfo("Exit after %d frames (exposure x%.2f)", opt.exitAfterFrames, exposure.current);
            glfwSetWindowShouldClose(window, GLFW_TRUE);
        }

        frameTimeAvg = frameTimeAvg * 0.9 + dt * 0.1;
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

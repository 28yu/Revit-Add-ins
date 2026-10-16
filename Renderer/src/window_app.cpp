// ウィンドウで表示して、マウスとキーボードで歩き回るモード。
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <string>
#include <vector>

#include "app.h"
#include "log.h"
#include "i18n.h"

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

struct InputState {
    bool dragging = false;
    double lastX = 0, lastY = 0;
    float speed = 3.0f;  // 歩く速さ（m/秒）
    float evOffset = 0;
    bool autoExposure = true;
    bool changed = false;  // カメラや設定が変わった（画質の足し合わせをやり直す）
    CameraController* cam = nullptr;
    CameraController initial;
};

void onScroll(GLFWwindow* w, double, double dy) {
    auto* in = static_cast<InputState*>(glfwGetWindowUserPointer(w));
    in->speed = std::clamp(in->speed * (dy > 0 ? 1.25f : 0.8f), 0.2f, 100.0f);
}

void onKey(GLFWwindow* w, int key, int, int action, int) {
    if (action != GLFW_PRESS) return;
    auto* in = static_cast<InputState*>(glfwGetWindowUserPointer(w));
    switch (key) {
        case GLFW_KEY_ESCAPE: glfwSetWindowShouldClose(w, GLFW_TRUE); break;
        case GLFW_KEY_EQUAL: case GLFW_KEY_KP_ADD: in->evOffset += 0.5f; in->changed = true; break;
        case GLFW_KEY_MINUS: case GLFW_KEY_KP_SUBTRACT: in->evOffset -= 0.5f; in->changed = true; break;
        case GLFW_KEY_F: in->autoExposure = !in->autoExposure; in->changed = true; break;
        case GLFW_KEY_R: *in->cam = in->initial; in->changed = true; break;
        default: break;
    }
}

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
    }, opt.gpu);

    Renderer renderer;
    renderer.init(ctx, scene);

    Swapchain sc;
    createSwapchain(ctx, window, sc);
    auto renderSize = [&](VkExtent2D e) {
        return VkExtent2D{std::max(16u, static_cast<uint32_t>(e.width * opt.renderScale)),
                          std::max(16u, static_cast<uint32_t>(e.height * opt.renderScale))};
    };
    VkExtent2D rs = renderSize(sc.extent);
    renderer.resize(rs.width, rs.height);

    CameraController cam;
    cam.setFrom(scene.camera);
    InputState input;
    input.cam = &cam;
    input.initial = cam;
    input.evOffset = opt.exposureEv;
    input.autoExposure = opt.autoExposure;
    glfwSetWindowUserPointer(window, &input);
    glfwSetScrollCallback(window, onScroll);
    glfwSetKeyCallback(window, onKey);
    logInfo("%s", tr("Controls"));

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
    auto last = std::chrono::steady_clock::now();
    double titleTimer = 0;
    int framesSinceTitle = 0;

    while (!glfwWindowShouldClose(window)) {
        glfwPollEvents();
        auto now = std::chrono::steady_clock::now();
        float dt = std::min(0.1f, std::chrono::duration<float>(now - last).count());
        last = now;

        // --- 操作 ---------------------------------------------------------
        double mx, my;
        glfwGetCursorPos(window, &mx, &my);
        bool leftDown = glfwGetMouseButton(window, GLFW_MOUSE_BUTTON_LEFT) == GLFW_PRESS;
        if (leftDown && input.dragging) {
            float sens = 0.0035f * cam.fovY / 60.0f;
            cam.yaw += static_cast<float>(mx - input.lastX) * sens;
            cam.pitch -= static_cast<float>(my - input.lastY) * sens;
            cam.pitch = std::clamp(cam.pitch, -1.55f, 1.55f);
            if (mx != input.lastX || my != input.lastY) input.changed = true;
        }
        input.dragging = leftDown;
        input.lastX = mx;
        input.lastY = my;

        Vec3 f = cam.forward(), r = cam.right();
        Vec3 flat{f.x, f.y, 0};
        float flatLen = std::sqrt(flat.x * flat.x + flat.y * flat.y);
        if (flatLen > 1e-4f) { flat.x /= flatLen; flat.y /= flatLen; }
        float step = input.speed * dt * (glfwGetKey(window, GLFW_KEY_LEFT_SHIFT) == GLFW_PRESS ? 3.0f : 1.0f);
        auto key = [&](int k) { return glfwGetKey(window, k) == GLFW_PRESS; };
        Vec3 move{0, 0, 0};
        if (key(GLFW_KEY_W) || key(GLFW_KEY_UP)) { move.x += flat.x; move.y += flat.y; }
        if (key(GLFW_KEY_S) || key(GLFW_KEY_DOWN)) { move.x -= flat.x; move.y -= flat.y; }
        if (key(GLFW_KEY_D) || key(GLFW_KEY_RIGHT)) { move.x += r.x; move.y += r.y; }
        if (key(GLFW_KEY_A) || key(GLFW_KEY_LEFT)) { move.x -= r.x; move.y -= r.y; }
        if (key(GLFW_KEY_E)) move.z += 1;
        if (key(GLFW_KEY_Q)) move.z -= 1;
        if (move.x != 0 || move.y != 0 || move.z != 0) {
            cam.eye.x += move.x * step;
            cam.eye.y += move.y * step;
            cam.eye.z += move.z * step;
            input.changed = true;
        }
        if (input.changed) {
            renderer.resetAccumulation();
            input.changed = false;
        }

        // --- 描画 ---------------------------------------------------------
        int fbw = 0, fbh = 0;
        glfwGetFramebufferSize(window, &fbw, &fbh);
        if (fbw == 0 || fbh == 0) {  // 最小化中
            glfwWaitEvents();
            continue;
        }
        if (needResize || static_cast<uint32_t>(fbw) != sc.extent.width || static_cast<uint32_t>(fbh) != sc.extent.height) {
            vkDeviceWaitIdle(ctx.device);
            createSwapchain(ctx, window, sc);
            rs = renderSize(sc.extent);
            renderer.resize(rs.width, rs.height);
            needResize = false;
        }

        uint32_t imageIndex = 0;
        VkResult acq = vkAcquireNextImageKHR(ctx.device, sc.handle, UINT64_MAX, imageReady, VK_NULL_HANDLE, &imageIndex);
        if (acq == VK_ERROR_OUT_OF_DATE_KHR) { needResize = true; continue; }
        if (acq != VK_SUCCESS && acq != VK_SUBOPTIMAL_KHR) vkCheck(acq, "vkAcquireNextImageKHR");

        ViewParams view = cam.view();
        view.maxBounces = opt.bounces;
        view.exposure = input.autoExposure ? exposure.current : std::exp2(input.evOffset);

        VkCommandBufferBeginInfo bi{VK_STRUCTURE_TYPE_COMMAND_BUFFER_BEGIN_INFO};
        bi.flags = VK_COMMAND_BUFFER_USAGE_ONE_TIME_SUBMIT_BIT;
        vkBeginCommandBuffer(cmd, &bi);
        renderer.record(cmd, view);

        // 描いた画像を画面の大きさに拡大・縮小して写す
        VkImage swapImage = sc.images[imageIndex];
        transitionImage(cmd, renderer.output().image, VK_IMAGE_LAYOUT_GENERAL, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL,
                        VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_ACCESS_SHADER_WRITE_BIT,
                        VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_READ_BIT);
        transitionImage(cmd, swapImage, VK_IMAGE_LAYOUT_UNDEFINED, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL,
                        VK_PIPELINE_STAGE_TRANSFER_BIT, 0, VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_WRITE_BIT);
        VkImageBlit blit{};
        blit.srcSubresource = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 0, 1};
        blit.srcOffsets[1] = {static_cast<int32_t>(rs.width), static_cast<int32_t>(rs.height), 1};
        blit.dstSubresource = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 0, 1};
        blit.dstOffsets[1] = {static_cast<int32_t>(sc.extent.width), static_cast<int32_t>(sc.extent.height), 1};
        vkCmdBlitImage(cmd, renderer.output().image, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL, swapImage,
                       VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, 1, &blit, VK_FILTER_LINEAR);
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
        exposure.update(renderer.finishFrame(), input.evOffset, 0.9f);

        if (opt.exitAfterFrames > 0 && static_cast<int>(renderer.accumulatedSamples()) >= opt.exitAfterFrames) {
            logInfo("Exit after %d frames (exposure x%.2f)", opt.exitAfterFrames, exposure.current);
            glfwSetWindowShouldClose(window, GLFW_TRUE);
        }

        // --- タイトルに状態を出す（1 秒ごと） -------------------------------
        titleTimer += dt;
        framesSinceTitle++;
        if (titleTimer >= 1.0) {
            char title[512];
            snprintf(title, sizeof(title), "%s  |  %s  |  %.0f fps  |  %u samples  |  %s%+.1f EV",
                     tr("WindowTitle"), ctx.gpuName.c_str(), framesSinceTitle / titleTimer,
                     renderer.accumulatedSamples(), input.autoExposure ? "Auto " : "", input.evOffset);
            glfwSetWindowTitle(window, title);
            titleTimer = 0;
            framesSinceTitle = 0;
        }
    }

    vkDeviceWaitIdle(ctx.device);
    vkDestroyFence(ctx.device, fence, nullptr);
    vkDestroySemaphore(ctx.device, imageReady, nullptr);
    for (VkSemaphore s : sc.renderDone) vkDestroySemaphore(ctx.device, s, nullptr);
    vkDestroySwapchainKHR(ctx.device, sc.handle, nullptr);
    renderer.destroy();
    ctx.destroy();
    glfwDestroyWindow(window);
    glfwTerminate();
    logInfo("Tools28Render exit");
    return 0;
}

}  // namespace t28

// 画像だけ作るモード（ウィンドウを出さない）。動作確認や 360°パノラマの書き出しに使う。
#include <chrono>
#include <cstring>

#include "app.h"
#include "log.h"

#define STB_IMAGE_WRITE_IMPLEMENTATION
#define STBIW_WINDOWS_UTF8
#include <stb_image_write.h>

namespace t28 {

int runOffscreen(const Options& opt, Scene& scene) {
    VkContext ctx;
    ctx.init({}, nullptr, opt.gpu);
    Renderer renderer;
    renderer.init(ctx, scene);
    uint32_t w = static_cast<uint32_t>(opt.width), h = static_cast<uint32_t>(opt.height);
    renderer.resize(w, h);

    CameraController cam;
    cam.setFrom(scene.camera);
    ViewParams view = cam.view();
    view.panorama = opt.panorama;
    view.maxBounces = opt.bounces;
    if (opt.panorama) {
        // パノラマは水平に構える（見上げ・見下ろしは見る人が操作する）
        cam.pitch = 0;
        ViewParams flat = cam.view();
        view.forward = flat.forward;
        view.right = flat.right;
        view.up = flat.up;
    }

    AutoExposure exposure;
    if (!opt.autoExposure) exposure.current = std::exp2(opt.exposureEv);

    VkCommandBufferAllocateInfo ai{VK_STRUCTURE_TYPE_COMMAND_BUFFER_ALLOCATE_INFO};
    ai.commandPool = ctx.commandPool;
    ai.level = VK_COMMAND_BUFFER_LEVEL_PRIMARY;
    ai.commandBufferCount = 1;
    VkCommandBuffer cmd;
    vkCheck(vkAllocateCommandBuffers(ctx.device, &ai, &cmd), "vkAllocateCommandBuffers");
    VkFenceCreateInfo fci{VK_STRUCTURE_TYPE_FENCE_CREATE_INFO};
    VkFence fence;
    vkCheck(vkCreateFence(ctx.device, &fci, nullptr, &fence), "vkCreateFence");

    auto start = std::chrono::steady_clock::now();
    for (int s = 0; s < opt.samples; ++s) {
        view.exposure = exposure.current;
        VkCommandBufferBeginInfo bi{VK_STRUCTURE_TYPE_COMMAND_BUFFER_BEGIN_INFO};
        bi.flags = VK_COMMAND_BUFFER_USAGE_ONE_TIME_SUBMIT_BIT;
        vkBeginCommandBuffer(cmd, &bi);
        renderer.record(cmd, view);
        vkEndCommandBuffer(cmd);
        VkSubmitInfo si{VK_STRUCTURE_TYPE_SUBMIT_INFO};
        si.commandBufferCount = 1;
        si.pCommandBuffers = &cmd;
        vkCheck(vkQueueSubmit(ctx.queue, 1, &si, fence), "vkQueueSubmit");
        vkWaitForFences(ctx.device, 1, &fence, VK_TRUE, UINT64_MAX);
        vkResetFences(ctx.device, 1, &fence);
        if (opt.autoExposure) exposure.update(renderer.finishFrame(), opt.exposureEv, 0.0f);
        if ((s + 1) % 16 == 0 || s + 1 == opt.samples) logInfo("sample %d / %d", s + 1, opt.samples);
    }
    double seconds = std::chrono::duration<double>(std::chrono::steady_clock::now() - start).count();
    logInfo("Rendered %d samples in %.1f s (exposure x%.2f)", opt.samples, seconds, exposure.current);

    // GPU の画像を CPU 側に写して PNG に保存する
    Buffer readback = ctx.createBuffer(static_cast<VkDeviceSize>(w) * h * 4, VK_BUFFER_USAGE_TRANSFER_DST_BIT,
                                       VK_MEMORY_PROPERTY_HOST_VISIBLE_BIT | VK_MEMORY_PROPERTY_HOST_COHERENT_BIT);
    VkCommandBuffer copy = ctx.beginOneShot();
    transitionImage(copy, renderer.output().image, VK_IMAGE_LAYOUT_GENERAL, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL,
                    VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_ACCESS_SHADER_WRITE_BIT, VK_PIPELINE_STAGE_TRANSFER_BIT,
                    VK_ACCESS_TRANSFER_READ_BIT);
    VkBufferImageCopy region{};
    region.imageSubresource = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 0, 1};
    region.imageExtent = {w, h, 1};
    vkCmdCopyImageToBuffer(copy, renderer.output().image, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL, readback.buffer, 1, &region);
    ctx.endOneShot(copy);

    int ok = stbi_write_png(opt.outputPath.c_str(), static_cast<int>(w), static_cast<int>(h), 4, readback.mapped,
                            static_cast<int>(w) * 4);
    ctx.destroyBuffer(readback);
    vkDestroyFence(ctx.device, fence, nullptr);
    renderer.destroy();
    ctx.destroy();
    if (!ok) {
        logError("Failed to write %s", opt.outputPath.c_str());
        return 1;
    }
    logInfo("Wrote %s", opt.outputPath.c_str());
    return 0;
}

}  // namespace t28

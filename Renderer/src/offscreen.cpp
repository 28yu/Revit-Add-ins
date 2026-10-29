// 画像だけ作るモード（ウィンドウを出さない）。動作確認や 360°パノラマの書き出しに使う。
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <cstring>

#include "app.h"
#include "log.h"

#define STB_IMAGE_WRITE_IMPLEMENTATION
#define STBIW_WINDOWS_UTF8
#include <stb_image_write.h>

namespace t28 {

std::vector<uint8_t> readImageRgba(VkContext& ctx, const Image& image, bool keepAlpha) {
    const uint32_t w = image.width, h = image.height;
    Buffer readback = ctx.createBuffer(static_cast<VkDeviceSize>(w) * h * 4, VK_BUFFER_USAGE_TRANSFER_DST_BIT,
                                       VK_MEMORY_PROPERTY_HOST_VISIBLE_BIT | VK_MEMORY_PROPERTY_HOST_COHERENT_BIT);
    VkCommandBuffer copy = ctx.beginOneShot();
    transitionImage(copy, image.image, VK_IMAGE_LAYOUT_GENERAL, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL,
                    VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_ACCESS_SHADER_WRITE_BIT, VK_PIPELINE_STAGE_TRANSFER_BIT,
                    VK_ACCESS_TRANSFER_READ_BIT);
    VkBufferImageCopy region{};
    region.imageSubresource = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 0, 1};
    region.imageExtent = {w, h, 1};
    vkCmdCopyImageToBuffer(copy, image.image, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL, readback.buffer, 1, &region);
    transitionImage(copy, image.image, VK_IMAGE_LAYOUT_TRANSFER_SRC_OPTIMAL, VK_IMAGE_LAYOUT_GENERAL,
                    VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_READ_BIT, VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT,
                    VK_ACCESS_SHADER_WRITE_BIT);
    ctx.endOneShot(copy);
    std::vector<uint8_t> rgba(static_cast<const uint8_t*>(readback.mapped),
                              static_cast<const uint8_t*>(readback.mapped) + static_cast<size_t>(w) * h * 4);
    ctx.destroyBuffer(readback);
    // 透明の書き出し以外は、透明にならないよう不透明度をすべて 255 にする
    if (!keepAlpha)
        for (size_t i = 3; i < rgba.size(); i += 4) rgba[i] = 255;
    return rgba;
}

int runOffscreen(const Options& opt, Scene& scene) {
    using Clock = std::chrono::steady_clock;
    auto secondsSince = [](Clock::time_point t) { return std::chrono::duration<double>(Clock::now() - t).count(); };
    VkContext ctx;
    ctx.init({}, nullptr, opt.gpu, !opt.forceSoftware);
    if (opt.gpuMemoryMb > 0) ctx.deviceLocalBytes = static_cast<VkDeviceSize>(opt.gpuMemoryMb) << 20;  // 動作確認用
    // 性能測定: シーンファイルの大きさ（ビューアの exe に入る量）と、保存にかかる時間
    size_t sceneBytes = 0;
    double saveSeconds = 0;
    if (opt.bench || !opt.saveScenePath.empty()) {
        auto t = Clock::now();
        std::vector<uint8_t> bytes;
        saveScene(scene, bytes);
        sceneBytes = bytes.size();
        if (!opt.saveScenePath.empty()) {
            FILE* f = openFile(opt.saveScenePath, "wb");
            if (f) {
                fwrite(bytes.data(), 1, bytes.size(), f);
                fclose(f);
            } else {
                logError("Failed to write %s", opt.saveScenePath.c_str());
            }
        }
        saveSeconds = secondsSince(t);
    }
    auto prepStart = Clock::now();
    Bvh bvh;
    uint32_t glassStart = prepareScene(scene, bvh, opt.instancing);
    double prepareSeconds = secondsSince(prepStart);
    auto initStart = Clock::now();
    Renderer renderer;
    renderer.skyWidth_ = static_cast<uint32_t>(opt.skySize);
    renderer.init(ctx, scene, glassStart, bvh);
    renderer.setLocalToneMapping(opt.localToneMapping);
    renderer.setSkySampling(opt.skySampling);
    renderer.setSpatialDenoise(opt.spatialDenoise);
    double initSeconds = secondsSince(initStart);
    // 3D ビューの「レンダリング設定」（照明・雲の量）
    renderer.setLighting(scene.revitLook.sun, scene.revitLook.artificial);
    renderer.setCloudAmount(scene.revitLook.cloudAmount);
    renderer.setLook(opt.lookBackground, opt.lookMono, opt.hideGround);
    renderer.setTransparentOutput(opt.transparent);
    if (opt.timeHour >= 0) renderer.setSunDirection(sunDirectionAt(scene, opt.timeHour));
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

    // 残像の確認: 横に歩きながら向きを変え続け、最後のフレームを保存する（--motion-test フレーム数）
    auto step = [&]() {
        Vec3 r = cam.right();
        cam.eye = {cam.eye.x + r.x * 0.04f, cam.eye.y + r.y * 0.04f, cam.eye.z};
        cam.yaw -= 0.006f;
    };
    if (opt.motionReference > 0) {
        for (int i = 0; i < opt.motionReference; ++i) step();
        ViewParams v = cam.view();
        v.maxBounces = opt.bounces;
        view = v;
    }
    // 重い 1 枚は横長の帯に分けて描く（最初は画像の大きさから決め、かかった時間で直す。--bands で決め打ちもできる）
    BandGovernor bands;
    bands.bands = opt.bands > 0 ? static_cast<uint32_t>(opt.bands) : BandGovernor::initialFor(w, h);
    renderer.setPathTraceBands(bands.bands);
    // 1 フレームを描いて終わるまで待つ（かかった秒数を返す）
    auto renderOne = [&](const ViewParams& v) {
        auto t = Clock::now();
        VkCommandBufferBeginInfo bi{VK_STRUCTURE_TYPE_COMMAND_BUFFER_BEGIN_INFO};
        bi.flags = VK_COMMAND_BUFFER_USAGE_ONE_TIME_SUBMIT_BIT;
        vkBeginCommandBuffer(cmd, &bi);
        renderer.record(cmd, v);
        vkEndCommandBuffer(cmd);
        VkSubmitInfo si{VK_STRUCTURE_TYPE_SUBMIT_INFO};
        si.commandBufferCount = 1;
        si.pCommandBuffers = &cmd;
        vkCheck(vkQueueSubmit(ctx.queue, 1, &si, fence), "vkQueueSubmit");
        vkCheck(vkWaitForFences(ctx.device, 1, &fence, VK_TRUE, UINT64_MAX), "vkWaitForFences");
        vkResetFences(ctx.device, 1, &fence);
        if (opt.autoExposure) exposure.update(renderer.finishFrame(), opt.exposureEv, 0.0f);
        double seconds = secondsSince(t);
        if (opt.bands <= 0 && bands.update(seconds)) {
            renderer.setPathTraceBands(bands.bands);
            logInfo("Frame took %.0f ms: path trace in %u bands", seconds * 1000.0, bands.bands);
        }
        return seconds;
    };
    struct FrameStats {
        double avg = 0, max = 0;
        void add(double s, int n) { avg += s / n; max = std::max(max, s); }
    };

    if (opt.bench) {
        // 性能測定: 止まっている間 benchFrames フレーム（その絵を出力の画像にする）→ 動いている間 benchFrames フレーム
        FrameStats still, moving;
        for (int i = 0; i < opt.benchFrames; ++i) {
            view.exposure = exposure.current;
            still.add(renderOne(view), opt.benchFrames);
        }
        std::vector<uint8_t> rgba = readImageRgba(ctx, renderer.output(), opt.transparent);
        // 白とびの割合（RGB すべてが 250 以上の画素）と、平均の明るさ（0〜255）
        size_t blown = 0;
        double lumSum = 0;
        for (size_t i = 0; i + 3 < rgba.size(); i += 4) {
            if (rgba[i] >= 250 && rgba[i + 1] >= 250 && rgba[i + 2] >= 250) ++blown;
            lumSum += 0.2126 * rgba[i] + 0.7152 * rgba[i + 1] + 0.0722 * rgba[i + 2];
        }
        size_t pixels = rgba.size() / 4;
        std::vector<uint8_t> png;
        stbi_write_png_to_func([](void* c, void* d, int n) {
            auto* v = static_cast<std::vector<uint8_t>*>(c);
            v->insert(v->end(), static_cast<uint8_t*>(d), static_cast<uint8_t*>(d) + n);
        }, &png, static_cast<int>(w), static_cast<int>(h), 4, rgba.data(), static_cast<int>(w) * 4);
        bool ok = true;
        if (FILE* f = openFile(opt.outputPath, "wb")) {
            fwrite(png.data(), 1, png.size(), f);
            fclose(f);
        } else {
            ok = false;
        }
        for (int i = 0; i < opt.benchFrames; ++i) {
            step();
            ViewParams v = cam.view();
            v.maxBounces = opt.bounces;
            v.exposure = exposure.current;
            moving.add(renderOne(v), opt.benchFrames);
        }
        logInfo("BENCH {\"drawn_triangles\":%zu,\"shapes\":%zu,\"placements\":%zu,\"triangles\":%zu,\"vertices\":%zu,\"textures\":%zu,\"scene_mb\":%.1f,\"save_s\":%.2f,"
                "\"load_s\":%.2f,\"prepare_s\":%.2f,\"init_s\":%.2f,\"gpu_mb\":%.0f,\"gpu_peak_mb\":%.0f,"
                "\"width\":%u,\"height\":%u,\"still_ms\":%.1f,\"still_max_ms\":%.1f,\"moving_ms\":%.1f,"
                "\"moving_max_ms\":%.1f,\"bands\":%u,\"png_kb\":%.0f,\"blown_pct\":%.2f,\"mean_luma\":%.1f,\"hw_rt\":%d}",
                scene.instancedTriangleCount(), scene.meshes.size(), scene.instances.size(),
                scene.triangleCount(), scene.vertices.size(), scene.textures.size(), sceneBytes / 1048576.0, saveSeconds,
                opt.sceneLoadSeconds, prepareSeconds, initSeconds, ctx.allocatedDeviceBytes / 1048576.0,
                ctx.peakDeviceBytes / 1048576.0, w, h, still.avg * 1000.0, still.max * 1000.0, moving.avg * 1000.0,
                moving.max * 1000.0, renderer.pathTraceBands(), png.size() / 1024.0, pixels ? 100.0 * blown / pixels : 0.0,
                pixels ? lumSum / pixels : 0.0, ctx.hardwareRt ? 1 : 0);
        vkDestroyFence(ctx.device, fence, nullptr);
        renderer.destroy();
        ctx.destroy();
        return ok ? 0 : 1;
    }

    const int frames = opt.motionTest > 0 ? opt.motionTest : opt.samples;
    auto start = std::chrono::steady_clock::now();
    for (int s = 0; s < frames; ++s) {
        if (opt.motionTest > 0) {
            step();
            ViewParams moving = cam.view();
            moving.maxBounces = opt.bounces;
            view = moving;
        }
        view.exposure = exposure.current;
        renderOne(view);
        if ((s + 1) % 16 == 0 || s + 1 == frames) logInfo("sample %d / %d", s + 1, frames);
    }
    double seconds = std::chrono::duration<double>(std::chrono::steady_clock::now() - start).count();
    logInfo("Rendered %d samples in %.1f s (exposure x%.2f)", frames, seconds, exposure.current);
    if (opt.motionTest > 0)
        logInfo("Final camera: eye %.4f %.4f %.4f yaw %.5f pitch %.5f", cam.eye.x, cam.eye.y, cam.eye.z, cam.yaw, cam.pitch);

    // GPU の画像を CPU 側に写して PNG に保存する
    std::vector<uint8_t> rgba = readImageRgba(ctx, renderer.output(), opt.transparent);
    int ok = stbi_write_png(opt.outputPath.c_str(), static_cast<int>(w), static_cast<int>(h), 4, rgba.data(),
                            static_cast<int>(w) * 4);
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

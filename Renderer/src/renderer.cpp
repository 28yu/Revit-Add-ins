#include "renderer.h"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstring>

#include "i18n.h"
#include "log.h"

namespace t28 {

// シェーダー（GPU 用プログラム）。shaders/*.comp を glslc で変換したもの
static const uint32_t kPathTraceHwSpv[] =  // RT コアで当たり判定
#include "pathtrace_hw.spv.inc"
    ;
static const uint32_t kPathTraceSwSpv[] =  // BVH を自分でたどる予備の方式
#include "pathtrace_sw.spv.inc"
    ;
static const uint32_t kTemporalSpv[] =     // ノイズ除去 その1（前のフレームを重ねる）
#include "temporal.spv.inc"
    ;
static const uint32_t kAtrousSpv[] =       // ノイズ除去 その2（周りの画素と平均する）
#include "atrous.spv.inc"
    ;
static const uint32_t kCompositeSpv[] =    // 仕上げ（色を掛け戻す・トーンマッピング）
#include "composite.spv.inc"
    ;
static const uint32_t kLocalDownSpv[] =    // ローカルトーンマッピング: 大まかな明るさを 8×8 画素ごとにまとめる（自動露出の集計も）
#include "localexp_down.spv.inc"
    ;
static const uint32_t kLocalBlurSpv[] =    // ローカルトーンマッピング: 大まかな明るさを、明るさの境目を守りながらぼかす
#include "localexp_blur.spv.inc"
    ;
static const uint32_t kOverlaySpv[] =      // 操作の案内を重ねる
#include "overlay.spv.inc"
    ;
static const uint32_t kSkySpv[] =          // 空と雲の画像を前もって描く
#include "sky.spv.inc"
    ;
static const uint32_t kSkyCdfSpv[] =       // 空の明るさの地図（空の光を直接狙うため）
#include "sky_cdf.spv.inc"
    ;

namespace {

const int kAtrousIterations = 5;           // 平均する回数（間隔 1, 2, 4, 8, 16 画素）
const float kMovingHistoryAlpha = 0.2f;    // 動いている間、今回の結果を混ぜる最小の割合（= 最大 5 フレーム分を重ねる。多いと残像になる）
const float kMaxHistory = 65536.0f;        // 止まっている間に重ねる回数の上限

struct PathTracePush {
    float camPos[4];
    float camForward[4];
    float camRight[4];
    float camUp[4];
    float sunDir[4];
    float misc[4];
    uint32_t frame[4];
    uint32_t look[4];  // 背景色・単色（RGBA8。a が 0 以外なら有効）
};
static_assert(sizeof(PathTracePush) <= 128, "push constants must fit in 128 bytes");

struct TemporalPush {
    float prevEye[4];
    float prevForward[4];
    float prevRight[4];
    float prevUp[4];
    float params[4];
};

struct AtrousPush {
    int32_t stepSize;
};

struct CompositePush {
    float exposure;
    uint32_t background;
    uint32_t transparent;
    float pivot;     // 画面の平均の明るさ（log2）。ローカルトーンマッピングはここを中心に縮める
    float strength;  // ローカルトーンマッピングの強さ（0 = 使わない）
    float brightIn;  // 明るい所（上位 5%）が平均から何段上か
    float brightOut; // それを何段上に収めるか
};

struct LocalDownPush {
    uint32_t background;
};

struct LocalBlurPush {
    int32_t dir[2];
};

struct SkyPush {
    float sunDir[4];
    float params[4];  // 空の明るさ, 始める行, 行の数
};

struct OverlayPush {
    int32_t offset[2];
    float opacity;
};

struct ExposureData {
    int32_t logLumSum;
    int32_t logLumCount;
    int32_t histogram[64];  // 8×8 画素ごとの明るさの度数分布（-16〜+16 段、0.5 段刻み）
};

struct GpuMaterial {
    float color[4];
    float params[4];     // 粗さ, 金属, 発光, 反射率
    int32_t tex[4];      // 色の画像の番号, 凹凸の画像の番号（-1 = 無し）, 凹凸の種類（0 = 高さ, 1 = 法線）, 未使用
    float colorMap[2][4];  // 色の画像の貼り方（[0] = xform[0..2] と 凹凸の深さ, [1] = xform[3..5]）
    float bumpMap[2][4];   // 凹凸の画像の貼り方
};

// 照明器具の光源（シェーダーの Light と同じ並び）
struct GpuLight {
    float positionRadius[4];
    float intensityType[4];  // rgb = 明るさ, w = 種類（0 = 全方向, 1 = 半球, 2 = スポット）
    float direction[4];      // xyz = 向き, w = cos（落ち始め）
    float cone[4];           // x = cos（届く範囲）
};

// 置き場所（シェーダーの Instance と同じ並び。160 バイト）
struct GpuInstance {
    float objectToWorld[12];  // 3 行 4 列（行ごと）
    float worldToObject[12];
    float normal[12];         // 向き（法線）を動かす 3 行 3 列（各行の 4 つ目は未使用）
    uint32_t mesh[4];         // x = 形の番号
};

// 形（シェーダーの Mesh と同じ並び）
struct GpuMesh {
    uint32_t firstTri, triCount, opaqueCount;  // 三角形の範囲。前から opaqueCount 個が不透明、残りがガラス
    uint32_t bvhRoot;                          // 予備の方式で使う、この形の BVH の根
};

// テクスチャ 1 枚分の置き場所（シェーダーの TexInfo と同じ並び）
struct GpuTexInfo {
    uint32_t offset;  // texels の何番目から始まるか（一番大きい段）
    uint32_t width, height;
    uint32_t levels;  // 縮小した段の数（ミップマップ。遠くの模様がちらつかないよう、離れたら小さい段を使う）
};

// GPU に送る頂点（24 バイト。shaders/common_head.glsl の vertexPos / vertexNormal / vertexUv と同じ並び）
struct GpuVertex {
    float px, py, pz;
    uint32_t normal;  // 面の向きを八面体に写して 2 つの 16 ビットに詰めたもの（0x8000 = 向き無し）
    float u, v;
};
static_assert(sizeof(GpuVertex) == 24, "GpuVertex must be 24 bytes");

uint32_t packNormal(float x, float y, float z) {
    float len = std::fabs(x) + std::fabs(y) + std::fabs(z);
    if (len < 1e-12f) return 0x8000u;
    x /= len;
    y /= len;
    if (z < 0.0f) {
        float ox = (1.0f - std::fabs(y)) * (x >= 0.0f ? 1.0f : -1.0f);
        float oy = (1.0f - std::fabs(x)) * (y >= 0.0f ? 1.0f : -1.0f);
        x = ox;
        y = oy;
    }
    auto q = [](float f) {
        int v = static_cast<int>(std::lround(std::max(-1.0f, std::min(1.0f, f)) * 32767.0f));
        return static_cast<uint32_t>(static_cast<uint16_t>(static_cast<int16_t>(v)));
    };
    return q(x) | (q(y) << 16);
}

std::vector<GpuVertex> packVertices(const std::vector<Vertex>& vertices) {
    std::vector<GpuVertex> out(vertices.size());
    for (size_t i = 0; i < vertices.size(); ++i) {
        const Vertex& v = vertices[i];
        out[i] = {v.px, v.py, v.pz, packNormal(v.nx, v.ny, v.nz), v.u, v.v};
    }
    return out;
}

float srgbToLinear(uint8_t c) {
    float s = c / 255.0f;
    return s <= 0.04045f ? s / 12.92f : std::pow((s + 0.055f) / 1.055f, 2.4f);
}

uint8_t linearToSrgb(float l) {
    l = std::min(std::max(l, 0.0f), 1.0f);
    float s = l <= 0.0031308f ? l * 12.92f : 1.055f * std::pow(l, 1.0f / 2.4f) - 0.055f;
    return static_cast<uint8_t>(s * 255.0f + 0.5f);
}

// 全テクスチャを 1 つの並びに詰める。段ごとに縦横半分にした画像も後ろに続けて入れる。
// maxSize より大きい画像は、縦横とも maxSize 以下になる段から入れる（大きなモデルで GPU のメモリに収めるため）
void packTextures(const std::vector<Texture>& textures, std::vector<GpuTexInfo>& infos, std::vector<uint32_t>& texels,
                  uint32_t maxSize) {
    for (const Texture& t : textures) {
        GpuTexInfo info{0, 0, 0, 0};
        bool started = false;
        uint32_t w = t.width, h = t.height;
        // 色の画像は明るさに直してから平均する（sRGB のまま平均すると縮小した段が暗くなる）。凹凸の画像はそのまま平均する
        auto toLinear = [&](uint8_t c) { return t.linear ? c / 255.0f : srgbToLinear(c); };
        auto toByte = [&](float v) {
            return t.linear ? static_cast<uint8_t>(std::min(std::max(v, 0.0f), 1.0f) * 255.0f + 0.5f) : linearToSrgb(v);
        };
        std::vector<float> level(static_cast<size_t>(w) * h * 4);
        for (size_t i = 0; i < level.size(); ++i)
            level[i] = (i % 4 == 3) ? t.rgba[i] / 255.0f : toLinear(t.rgba[i]);
        while (true) {
            bool keep = (w <= maxSize && h <= maxSize) || (w == 1 && h == 1);
            if (keep && !started) {
                started = true;
                info.offset = static_cast<uint32_t>(texels.size());
                info.width = w;
                info.height = h;
            }
            for (size_t i = 0; keep && i < static_cast<size_t>(w) * h; ++i) {
                uint32_t r = toByte(level[i * 4]), g = toByte(level[i * 4 + 1]), b = toByte(level[i * 4 + 2]);
                uint32_t a = static_cast<uint32_t>(std::min(std::max(level[i * 4 + 3], 0.0f), 1.0f) * 255.0f + 0.5f);
                texels.push_back(r | (g << 8) | (b << 16) | (a << 24));
            }
            if (keep) ++info.levels;
            if (w == 1 && h == 1) break;
            uint32_t nw = std::max(w / 2, 1u), nh = std::max(h / 2, 1u);
            std::vector<float> next(static_cast<size_t>(nw) * nh * 4, 0.0f);
            for (uint32_t y = 0; y < nh; ++y)
                for (uint32_t x = 0; x < nw; ++x)
                    for (int c = 0; c < 4; ++c) {
                        float sum = 0.0f;
                        int n = 0;
                        for (uint32_t dy = 0; dy < 2; ++dy)
                            for (uint32_t dx = 0; dx < 2; ++dx) {
                                uint32_t sx = std::min(x * 2 + dx, w - 1), sy = std::min(y * 2 + dy, h - 1);
                                sum += level[(static_cast<size_t>(sy) * w + sx) * 4 + c];
                                ++n;
                            }
                        next[(static_cast<size_t>(y) * nw + x) * 4 + c] = sum / n;
                    }
            level.swap(next);
            w = nw;
            h = nh;
        }
        infos.push_back(info);
    }
    // 空のバッファは作れないので、テクスチャが無くても 1 つは入れておく
    if (infos.empty()) infos.push_back({0, 1, 1, 1});
    if (texels.empty()) texels.push_back(0xffffffffu);
}

VkDeviceSize alignUp(VkDeviceSize v, VkDeviceSize a) { return (v + a - 1) / a * a; }

struct BindingSpec {
    uint32_t binding;
    VkDescriptorType type;
};

// シェーダー1つ分の入出力の一覧・パイプラインを作る（setCount 個の入出力の組を用意する）
Renderer::ComputePass createPass(VkDevice device, const std::vector<BindingSpec>& specs, uint32_t setCount,
                                 uint32_t pushSize, const uint32_t* spv, size_t spvSize, const char* name) {
    Renderer::ComputePass pass;
    std::vector<VkDescriptorSetLayoutBinding> bindings(specs.size());
    uint32_t images = 0, buffers = 0, accels = 0;
    for (size_t i = 0; i < specs.size(); ++i) {
        bindings[i].binding = specs[i].binding;
        bindings[i].descriptorType = specs[i].type;
        bindings[i].descriptorCount = 1;
        bindings[i].stageFlags = VK_SHADER_STAGE_COMPUTE_BIT;
        if (specs[i].type == VK_DESCRIPTOR_TYPE_STORAGE_IMAGE) images++;
        else if (specs[i].type == VK_DESCRIPTOR_TYPE_STORAGE_BUFFER) buffers++;
        else accels++;
    }
    VkDescriptorSetLayoutCreateInfo lci{VK_STRUCTURE_TYPE_DESCRIPTOR_SET_LAYOUT_CREATE_INFO};
    lci.bindingCount = static_cast<uint32_t>(bindings.size());
    lci.pBindings = bindings.data();
    vkCheck(vkCreateDescriptorSetLayout(device, &lci, nullptr, &pass.layout), name);

    std::vector<VkDescriptorPoolSize> sizes;
    if (images) sizes.push_back({VK_DESCRIPTOR_TYPE_STORAGE_IMAGE, images * setCount});
    if (buffers) sizes.push_back({VK_DESCRIPTOR_TYPE_STORAGE_BUFFER, buffers * setCount});
    if (accels) sizes.push_back({VK_DESCRIPTOR_TYPE_ACCELERATION_STRUCTURE_KHR, accels * setCount});
    VkDescriptorPoolCreateInfo pci{VK_STRUCTURE_TYPE_DESCRIPTOR_POOL_CREATE_INFO};
    pci.maxSets = setCount;
    pci.poolSizeCount = static_cast<uint32_t>(sizes.size());
    pci.pPoolSizes = sizes.data();
    vkCheck(vkCreateDescriptorPool(device, &pci, nullptr, &pass.pool), name);
    std::vector<VkDescriptorSetLayout> layouts(setCount, pass.layout);
    pass.sets.resize(setCount);
    VkDescriptorSetAllocateInfo ai{VK_STRUCTURE_TYPE_DESCRIPTOR_SET_ALLOCATE_INFO};
    ai.descriptorPool = pass.pool;
    ai.descriptorSetCount = setCount;
    ai.pSetLayouts = layouts.data();
    vkCheck(vkAllocateDescriptorSets(device, &ai, pass.sets.data()), name);

    VkPushConstantRange pcr{VK_SHADER_STAGE_COMPUTE_BIT, 0, pushSize};
    VkPipelineLayoutCreateInfo plci{VK_STRUCTURE_TYPE_PIPELINE_LAYOUT_CREATE_INFO};
    plci.setLayoutCount = 1;
    plci.pSetLayouts = &pass.layout;
    plci.pushConstantRangeCount = 1;
    plci.pPushConstantRanges = &pcr;
    vkCheck(vkCreatePipelineLayout(device, &plci, nullptr, &pass.pipelineLayout), name);

    VkShaderModuleCreateInfo smci{VK_STRUCTURE_TYPE_SHADER_MODULE_CREATE_INFO};
    smci.codeSize = spvSize;
    smci.pCode = spv;
    VkShaderModule module;
    vkCheck(vkCreateShaderModule(device, &smci, nullptr, &module), name);
    VkComputePipelineCreateInfo cpci{VK_STRUCTURE_TYPE_COMPUTE_PIPELINE_CREATE_INFO};
    cpci.stage.sType = VK_STRUCTURE_TYPE_PIPELINE_SHADER_STAGE_CREATE_INFO;
    cpci.stage.stage = VK_SHADER_STAGE_COMPUTE_BIT;
    cpci.stage.module = module;
    cpci.stage.pName = "main";
    cpci.layout = pass.pipelineLayout;
    vkCheck(vkCreateComputePipelines(device, VK_NULL_HANDLE, 1, &cpci, nullptr, &pass.pipeline), name);
    vkDestroyShaderModule(device, module, nullptr);
    return pass;
}

void destroyPass(VkDevice device, Renderer::ComputePass& pass) {
    if (pass.pipeline) vkDestroyPipeline(device, pass.pipeline, nullptr);
    if (pass.pipelineLayout) vkDestroyPipelineLayout(device, pass.pipelineLayout, nullptr);
    if (pass.pool) vkDestroyDescriptorPool(device, pass.pool, nullptr);
    if (pass.layout) vkDestroyDescriptorSetLayout(device, pass.layout, nullptr);
    pass = Renderer::ComputePass{};
}

// 入出力の組に、画像・バッファ・加速構造を結び付ける
struct Bind {
    uint32_t binding;
    VkImageView image = VK_NULL_HANDLE;
    VkBuffer buffer = VK_NULL_HANDLE;
    const VkAccelerationStructureKHR* accel = nullptr;
};

void writeSet(VkDevice device, VkDescriptorSet set, const std::vector<Bind>& binds) {
    std::vector<VkWriteDescriptorSet> writes(binds.size());
    std::vector<VkDescriptorImageInfo> images(binds.size());
    std::vector<VkDescriptorBufferInfo> buffers(binds.size());
    std::vector<VkWriteDescriptorSetAccelerationStructureKHR> accels(binds.size());
    for (size_t i = 0; i < binds.size(); ++i) {
        VkWriteDescriptorSet& w = writes[i];
        w = VkWriteDescriptorSet{VK_STRUCTURE_TYPE_WRITE_DESCRIPTOR_SET};
        w.dstSet = set;
        w.dstBinding = binds[i].binding;
        w.descriptorCount = 1;
        if (binds[i].accel) {
            accels[i] = VkWriteDescriptorSetAccelerationStructureKHR{VK_STRUCTURE_TYPE_WRITE_DESCRIPTOR_SET_ACCELERATION_STRUCTURE_KHR};
            accels[i].accelerationStructureCount = 1;
            accels[i].pAccelerationStructures = binds[i].accel;
            w.descriptorType = VK_DESCRIPTOR_TYPE_ACCELERATION_STRUCTURE_KHR;
            w.pNext = &accels[i];
        } else if (binds[i].image) {
            images[i] = {VK_NULL_HANDLE, binds[i].image, VK_IMAGE_LAYOUT_GENERAL};
            w.descriptorType = VK_DESCRIPTOR_TYPE_STORAGE_IMAGE;
            w.pImageInfo = &images[i];
        } else {
            buffers[i] = {binds[i].buffer, 0, VK_WHOLE_SIZE};
            w.descriptorType = VK_DESCRIPTOR_TYPE_STORAGE_BUFFER;
            w.pBufferInfo = &buffers[i];
        }
    }
    vkUpdateDescriptorSets(device, static_cast<uint32_t>(writes.size()), writes.data(), 0, nullptr);
}

// 前の計算の書き込みが終わってから次の計算が読むようにする
void computeBarrier(VkCommandBuffer cmd) {
    VkMemoryBarrier b{VK_STRUCTURE_TYPE_MEMORY_BARRIER};
    b.srcAccessMask = VK_ACCESS_SHADER_WRITE_BIT;
    b.dstAccessMask = VK_ACCESS_SHADER_READ_BIT | VK_ACCESS_SHADER_WRITE_BIT;
    vkCmdPipelineBarrier(cmd, VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, 0, 1, &b, 0,
                         nullptr, 0, nullptr);
}

bool sameView(const ViewParams& a, const ViewParams& b) {
    auto close = [](Vec3 p, Vec3 q) {
        return std::fabs(p.x - q.x) < 1e-6f && std::fabs(p.y - q.y) < 1e-6f && std::fabs(p.z - q.z) < 1e-6f;
    };
    return close(a.eye, b.eye) && close(a.forward, b.forward) && close(a.up, b.up) && a.fovY == b.fovY &&
           a.panorama == b.panorama && a.maxBounces == b.maxBounces;
}

}  // namespace

void Renderer::init(VkContext& ctx, const Scene& scene, uint32_t glassStart, const Bvh& bvh) {
    ctx_ = &ctx;
    hardwareRt_ = ctx.hardwareRt;
    exposure_ = ctx.createBuffer(sizeof(ExposureData), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT,
                                 VK_MEMORY_PROPERTY_HOST_VISIBLE_BIT | VK_MEMORY_PROPERTY_HOST_COHERENT_BIT);
    memset(exposure_.mapped, 0, sizeof(ExposureData));
    // 空と雲の画像（横 = 方位、縦 = 高度。縦は横の 1/4）
    skyImage_ = ctx.createImage(skyWidth_, skyWidth_ / 4, VK_FORMAT_R16G16B16A16_SFLOAT, VK_IMAGE_USAGE_STORAGE_BIT);
    // 空と地面の明るさの地図: 行の累積 64 + 列の累積 64×64 + マスの割合 64×64（shaders/sky_cdf.comp）
    skyDist_ = ctx.createBuffer((64 + 2 * 64 * 64) * sizeof(float), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT,
                                VK_MEMORY_PROPERTY_DEVICE_LOCAL_BIT);
    uploadScene(scene, glassStart, bvh);
    createPasses();
    computeSky();
}

void Renderer::computeSky() {
    auto start = std::chrono::steady_clock::now();
    writeSet(ctx_->device, skyPass_.sets[0], {{0, skyImage_.view}});
    const uint32_t w = skyImage_.width, h = skyImage_.height;
    // 1 回の計算が長いと Windows が GPU を止めてしまう（2 秒）ので、少しずつ分けて計算する
    const uint32_t rowsPerChunk = std::max(8u, (1u << 18) / w);
    for (uint32_t row = 0; row < h; row += rowsPerChunk) {
        VkCommandBuffer cmd = ctx_->beginOneShot();
        if (row == 0)
            transitionImage(cmd, skyImage_.image, VK_IMAGE_LAYOUT_UNDEFINED, VK_IMAGE_LAYOUT_GENERAL,
                            VK_PIPELINE_STAGE_TOP_OF_PIPE_BIT, 0, VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_ACCESS_SHADER_WRITE_BIT);
        uint32_t rows = std::min(rowsPerChunk, h - row);
        SkyPush sp{{sunDir_.x, sunDir_.y, sunDir_.z, effectiveSun()}, {effectiveSky(), static_cast<float>(row), static_cast<float>(rows), cloudAmount_}};
        dispatch(cmd, skyPass_, 0, &sp, sizeof(sp), w, rows);
        ctx_->endOneShot(cmd);
    }
    // 空と地面の明るさの地図を作り直す（1 つのまとまりで計算する小さな処理）。
    // 地面の明るさはおおよそでよい（狙う割合を決めるだけ）: 地面の色 × (太陽の高さに応じた直射 + 空の光)
    float sunZ = std::max(sunDir_.z, 0.0f);
    float low = std::min(sunZ / 0.3f, 1.0f);
    float groundLum = 0.4f * (effectiveSun() * sunZ * (0.3f + 0.7f * low * low * (3.0f - 2.0f * low)) + effectiveSky() * 0.8f);
    writeSet(ctx_->device, skyCdfPass_.sets[0], {{0, skyImage_.view}, {1, {}, skyDist_.buffer}});
    VkCommandBuffer cmd = ctx_->beginOneShot();
    computeBarrier(cmd);
    dispatch(cmd, skyCdfPass_, 0, &groundLum, sizeof(groundLum), 8, 8);
    ctx_->endOneShot(cmd);
    skySun_ = sunDir_;
    double ms = std::chrono::duration<double, std::milli>(std::chrono::steady_clock::now() - start).count();
    logInfo("Sky computed: %ux%u in %.0f ms", w, h, ms);
}

void Renderer::reloadScene(const Scene& scene, uint32_t glassStart, const Bvh& bvh) {
    vkDeviceWaitIdle(ctx_->device);
    destroyScene();
    uploadScene(scene, glassStart, bvh);
    // 太陽の向きが変わったときだけ、空を描き直す
    if (std::fabs(sunDir_.x - skySun_.x) + std::fabs(sunDir_.y - skySun_.y) + std::fabs(sunDir_.z - skySun_.z) > 1e-5f) computeSky();
    if (output_.view) writeDescriptors();
    resetHistory_ = true;
}

void Renderer::destroyScene() {
    destroyAccel(tlas_);
    for (AccelStruct& b : blas_) destroyAccel(b);
    blas_.clear();
    for (Buffer* b : {&vertices_, &indices_, &triMaterials_, &materials_, &texInfos_, &texels_, &lights_, &instances_, &bvhNodes_,
                      &bvhTris_, &sceneInstances_, &meshes_})
        ctx_->destroyBuffer(*b);
}

void Renderer::uploadScene(const Scene& scene, uint32_t glassStart, const Bvh& bvh) {
    VkContext& ctx = *ctx_;
    glassStart_ = glassStart;
    sunDir_ = scene.sunDir;
    sunIntensity_ = scene.sunIntensity;
    skyIntensity_ = scene.skyIntensity;
    logInfo("Scene: %zu vertices, %zu triangles (%u opaque, %zu glass), %zu materials, %zu shapes, %zu placements "
            "(%zu triangles drawn)",
            scene.vertices.size(), scene.triangleCount(), glassStart_, scene.triangleCount() - glassStart_,
            scene.materials.size(), scene.meshes.size(), scene.instances.size(), scene.instancedTriangleCount());

    // RT コアを使うときだけ、加速構造の材料として使う印を付ける（予備の方式ではその機能自体が無い）
    const VkBufferUsageFlags geomUsage =
        VK_BUFFER_USAGE_STORAGE_BUFFER_BIT |
        (hardwareRt_ ? VK_BUFFER_USAGE_SHADER_DEVICE_ADDRESS_BIT | VK_BUFFER_USAGE_ACCELERATION_STRUCTURE_BUILD_INPUT_READ_ONLY_BIT_KHR : 0);
    // 形状のデータが 1 つの記憶領域の上限を超えると、GPU が範囲外を読んで落ちる。分かる言葉で止める
    const VkDeviceSize vertexBytes = scene.vertices.size() * sizeof(GpuVertex);
    const VkDeviceSize indexBytes = scene.indices.size() * sizeof(uint32_t);
    const VkDeviceSize storageLimit = ctx.properties.limits.maxStorageBufferRange;
    // 予備の方式（RT コアの無い GPU）では、自分でたどる BVH も 1 つの記憶領域に入れる
    const VkDeviceSize bvhBytes = hardwareRt_ ? 0 : bvh.nodes().size() * sizeof(BvhNode);
    const VkDeviceSize instanceBytes = scene.instances.size() * sizeof(GpuInstance);
    if (vertexBytes > storageLimit || indexBytes > storageLimit || bvhBytes > storageLimit || instanceBytes > storageLimit)
        fatal("%s\n(vertices %.0f MB, indices %.0f MB, BVH %.0f MB, limit %.0f MB)", tr("ModelTooLarge"),
              vertexBytes / 1048576.0, indexBytes / 1048576.0, bvhBytes / 1048576.0, storageLimit / 1048576.0);
    {
        std::vector<GpuVertex> packed = packVertices(scene.vertices);
        vertices_ = ctx.createDeviceBuffer(packed.data(), packed.size() * sizeof(GpuVertex), geomUsage);
    }
    indices_ = ctx.createDeviceBuffer(scene.indices.data(), scene.indices.size() * sizeof(uint32_t), geomUsage);
    triMaterials_ = ctx.createDeviceBuffer(scene.triMaterials.data(), scene.triMaterials.size() * sizeof(uint32_t),
                                           VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);

    // 置き場所: 形の座標 → 世界の座標、その逆、向き（法線）の動かし方（逆行列の転置）、形の番号
    std::vector<GpuInstance> insts(scene.instances.size());
    for (size_t i = 0; i < scene.instances.size(); ++i) {
        const Instance& in = scene.instances[i];
        GpuInstance& g = insts[i];
        std::copy(in.m, in.m + 12, g.objectToWorld);
        invertAffine(in.m, g.worldToObject);
        for (int r = 0; r < 3; ++r) {
            for (int c = 0; c < 3; ++c) g.normal[r * 4 + c] = g.worldToObject[c * 4 + r];
            g.normal[r * 4 + 3] = 0.0f;
        }
        g.mesh[0] = in.mesh;
    }
    sceneInstances_ = ctx.createDeviceBuffer(insts.data(), insts.size() * sizeof(GpuInstance), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);
    std::vector<GpuMesh> meshes(scene.meshes.size());
    for (size_t i = 0; i < scene.meshes.size(); ++i) {
        const MeshRange& m = scene.meshes[i];
        meshes[i] = {m.firstTri, m.triCount, m.opaqueCount, i < bvh.meshRoots().size() ? bvh.meshRoots()[i] : 0u};
    }
    meshes_ = ctx.createDeviceBuffer(meshes.data(), meshes.size() * sizeof(GpuMesh), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);

    std::vector<GpuMaterial> mats;
    for (const Material& m : scene.materials)
        mats.push_back({{m.r, m.g, m.b, m.opacity},
                        {m.roughness, m.metallic, m.emissive, m.f0},
                        {m.color.texture, m.bump.texture, static_cast<int32_t>(m.bumpType), 0},
                        {{m.color.xform[0], m.color.xform[1], m.color.xform[2], m.bumpDepth},
                         {m.color.xform[3], m.color.xform[4], m.color.xform[5], 0}},
                        {{m.bump.xform[0], m.bump.xform[1], m.bump.xform[2], 0},
                         {m.bump.xform[3], m.bump.xform[4], m.bump.xform[5], 0}}});
    materials_ = ctx.createDeviceBuffer(mats.data(), mats.size() * sizeof(GpuMaterial), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);

    std::vector<GpuLight> lights;
    for (const Light& l : scene.lights)
        lights.push_back({{l.position.x, l.position.y, l.position.z, l.radius},
                          {l.intensity.x, l.intensity.y, l.intensity.z, static_cast<float>(l.type)},
                          {l.direction.x, l.direction.y, l.direction.z, l.cosInner},
                          {l.cosOuter, 0, 0, 0}});
    lightCount_ = static_cast<uint32_t>(lights.size());
    // 地面の高さ = モデルの一番下（置き場所の位置まで含めた全体の箱の下の面）
    groundZ_ = bvh.nodes().empty() ? 0.0f : bvh.boundsMin().z;
    if (lights.empty()) lights.push_back({});  // 空のバッファは作れないので 1 つ入れておく（数は 0 のまま）
    lights_ = ctx.createDeviceBuffer(lights.data(), lights.size() * sizeof(GpuLight), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);
    logInfo("Lights: %u", lightCount_);

    // 1 つの記憶領域の上限（GPU ごとに違う）と GPU のメモリの 1/4 に収まるまで、テクスチャを縮める
    const VkDeviceSize maxRange = ctx.properties.limits.maxStorageBufferRange;
    // 形がすでに使っている分も考えて、GPU のメモリの 6 割を超えないようにする（加速構造と描画用の画像の分を残す）
    VkDeviceSize texBudget = std::min<VkDeviceSize>(maxRange, ctx.deviceLocalBytes > 0 ? ctx.deviceLocalBytes / 4 : maxRange);
    if (ctx.deviceLocalBytes > 0) {
        const VkDeviceSize cap = ctx.deviceLocalBytes * 6 / 10;
        const VkDeviceSize left = cap > ctx.allocatedDeviceBytes ? cap - ctx.allocatedDeviceBytes : 0;
        texBudget = std::min(texBudget, std::max<VkDeviceSize>(left, 32ull << 20));
    }
    std::vector<GpuTexInfo> texInfos;
    std::vector<uint32_t> texels;
    for (uint32_t maxSize = 16384;; maxSize /= 2) {
        texInfos.clear();
        texels.clear();
        packTextures(scene.textures, texInfos, texels, maxSize);
        if (texels.size() * sizeof(uint32_t) <= texBudget || maxSize <= 64) {
            if (maxSize < 16384) logInfo("Textures reduced to %u px to fit GPU memory", maxSize);
            break;
        }
    }
    texInfos_ = ctx.createDeviceBuffer(texInfos.data(), texInfos.size() * sizeof(GpuTexInfo), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);
    texels_ = ctx.createDeviceBuffer(texels.data(), texels.size() * sizeof(uint32_t), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);
    logInfo("Textures: %zu (%.1f MB)", scene.textures.size(), texels.size() * 4.0 / (1024.0 * 1024.0));

    if (hardwareRt_) {
        buildAccelerationStructures(scene);
    } else {
        bvhNodes_ = ctx.createDeviceBuffer(bvh.nodes().data(), bvh.nodes().size() * sizeof(BvhNode), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);
        bvhTris_ = ctx.createDeviceBuffer(bvh.triIndices().data(), bvh.triIndices().size() * sizeof(uint32_t), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);
        logInfo("BVH uploaded: %zu nodes", bvh.nodes().size());
    }
    logInfo("GPU memory used: %.0f MB of %.0f MB", ctx.allocatedDeviceBytes / 1048576.0, ctx.deviceLocalBytes / 1048576.0);
}

// --- 加速構造（光線の当たり判定を速くするための索引） ----------------------------

Renderer::AccelStruct Renderer::buildBlas(VkDeviceSize indexOffsetBytes, uint32_t triangleCount, uint32_t vertexCount) {
    VkAccelerationStructureGeometryKHR geom{VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_GEOMETRY_KHR};
    geom.geometryType = VK_GEOMETRY_TYPE_TRIANGLES_KHR;
    geom.flags = VK_GEOMETRY_OPAQUE_BIT_KHR;
    auto& tri = geom.geometry.triangles;
    tri.sType = VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_GEOMETRY_TRIANGLES_DATA_KHR;
    tri.vertexFormat = VK_FORMAT_R32G32B32_SFLOAT;
    tri.vertexData.deviceAddress = vertices_.address;
    tri.vertexStride = sizeof(GpuVertex);
    tri.maxVertex = vertexCount - 1;
    tri.indexType = VK_INDEX_TYPE_UINT32;
    tri.indexData.deviceAddress = indices_.address + indexOffsetBytes;

    VkAccelerationStructureBuildGeometryInfoKHR build{VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_BUILD_GEOMETRY_INFO_KHR};
    build.type = VK_ACCELERATION_STRUCTURE_TYPE_BOTTOM_LEVEL_KHR;
    // 作った後に詰め直せるようにする（大きなモデルで GPU のメモリを節約するため）
    build.flags = VK_BUILD_ACCELERATION_STRUCTURE_PREFER_FAST_TRACE_BIT_KHR |
                  VK_BUILD_ACCELERATION_STRUCTURE_ALLOW_COMPACTION_BIT_KHR;
    build.mode = VK_BUILD_ACCELERATION_STRUCTURE_MODE_BUILD_KHR;
    build.geometryCount = 1;
    build.pGeometries = &geom;

    VkAccelerationStructureBuildSizesInfoKHR sizes{VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_BUILD_SIZES_INFO_KHR};
    vkGetAccelerationStructureBuildSizesKHR(ctx_->device, VK_ACCELERATION_STRUCTURE_BUILD_TYPE_DEVICE_KHR,
                                            &build, &triangleCount, &sizes);

    AccelStruct as;
    as.buffer = ctx_->createBuffer(sizes.accelerationStructureSize,
                                   VK_BUFFER_USAGE_ACCELERATION_STRUCTURE_STORAGE_BIT_KHR | VK_BUFFER_USAGE_SHADER_DEVICE_ADDRESS_BIT,
                                   VK_MEMORY_PROPERTY_DEVICE_LOCAL_BIT);
    VkAccelerationStructureCreateInfoKHR ci{VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_CREATE_INFO_KHR};
    ci.buffer = as.buffer.buffer;
    ci.size = sizes.accelerationStructureSize;
    ci.type = VK_ACCELERATION_STRUCTURE_TYPE_BOTTOM_LEVEL_KHR;
    vkCheck(vkCreateAccelerationStructureKHR(ctx_->device, &ci, nullptr, &as.handle), "vkCreateAccelerationStructureKHR (BLAS)");

    // 作業用の領域（決められた境界に合わせる必要がある）
    VkPhysicalDeviceAccelerationStructurePropertiesKHR asProps{VK_STRUCTURE_TYPE_PHYSICAL_DEVICE_ACCELERATION_STRUCTURE_PROPERTIES_KHR};
    VkPhysicalDeviceProperties2 props2{VK_STRUCTURE_TYPE_PHYSICAL_DEVICE_PROPERTIES_2};
    props2.pNext = &asProps;
    vkGetPhysicalDeviceProperties2(ctx_->physical, &props2);
    VkDeviceSize align = asProps.minAccelerationStructureScratchOffsetAlignment;
    Buffer scratch = ctx_->createBuffer(sizes.buildScratchSize + align,
                                        VK_BUFFER_USAGE_STORAGE_BUFFER_BIT | VK_BUFFER_USAGE_SHADER_DEVICE_ADDRESS_BIT,
                                        VK_MEMORY_PROPERTY_DEVICE_LOCAL_BIT);

    build.dstAccelerationStructure = as.handle;
    build.scratchData.deviceAddress = alignUp(scratch.address, align);
    VkAccelerationStructureBuildRangeInfoKHR range{triangleCount, 0, 0, 0};
    const VkAccelerationStructureBuildRangeInfoKHR* ranges = &range;
    VkQueryPoolCreateInfo qpi{VK_STRUCTURE_TYPE_QUERY_POOL_CREATE_INFO};
    qpi.queryType = VK_QUERY_TYPE_ACCELERATION_STRUCTURE_COMPACTED_SIZE_KHR;
    qpi.queryCount = 1;
    VkQueryPool queryPool = VK_NULL_HANDLE;
    vkCheck(vkCreateQueryPool(ctx_->device, &qpi, nullptr, &queryPool), "vkCreateQueryPool");

    VkCommandBuffer cmd = ctx_->beginOneShot();
    vkCmdResetQueryPool(cmd, queryPool, 0, 1);
    vkCmdBuildAccelerationStructuresKHR(cmd, 1, &build, &ranges);
    VkMemoryBarrier mb{VK_STRUCTURE_TYPE_MEMORY_BARRIER};
    mb.srcAccessMask = VK_ACCESS_ACCELERATION_STRUCTURE_WRITE_BIT_KHR;
    mb.dstAccessMask = VK_ACCESS_ACCELERATION_STRUCTURE_READ_BIT_KHR;
    vkCmdPipelineBarrier(cmd, VK_PIPELINE_STAGE_ACCELERATION_STRUCTURE_BUILD_BIT_KHR,
                         VK_PIPELINE_STAGE_ACCELERATION_STRUCTURE_BUILD_BIT_KHR, 0, 1, &mb, 0, nullptr, 0, nullptr);
    vkCmdWriteAccelerationStructuresPropertiesKHR(cmd, 1, &as.handle,
                                                  VK_QUERY_TYPE_ACCELERATION_STRUCTURE_COMPACTED_SIZE_KHR, queryPool, 0);
    ctx_->endOneShot(cmd);
    ctx_->destroyBuffer(scratch);

    // 詰め直した大きさで作り直して写す（元の大きさの半分ほどになることが多い）
    VkDeviceSize compactSize = 0;
    VkResult qr = vkGetQueryPoolResults(ctx_->device, queryPool, 0, 1, sizeof(compactSize), &compactSize, sizeof(compactSize),
                                        VK_QUERY_RESULT_64_BIT | VK_QUERY_RESULT_WAIT_BIT);
    vkDestroyQueryPool(ctx_->device, queryPool, nullptr);
    if (qr == VK_SUCCESS && compactSize > 0 && compactSize < sizes.accelerationStructureSize) {
        AccelStruct compact;
        compact.buffer = ctx_->createBuffer(compactSize,
                                            VK_BUFFER_USAGE_ACCELERATION_STRUCTURE_STORAGE_BIT_KHR | VK_BUFFER_USAGE_SHADER_DEVICE_ADDRESS_BIT,
                                            VK_MEMORY_PROPERTY_DEVICE_LOCAL_BIT);
        VkAccelerationStructureCreateInfoKHR cci{VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_CREATE_INFO_KHR};
        cci.buffer = compact.buffer.buffer;
        cci.size = compactSize;
        cci.type = VK_ACCELERATION_STRUCTURE_TYPE_BOTTOM_LEVEL_KHR;
        vkCheck(vkCreateAccelerationStructureKHR(ctx_->device, &cci, nullptr, &compact.handle), "vkCreateAccelerationStructureKHR (compact)");
        VkCopyAccelerationStructureInfoKHR copy{VK_STRUCTURE_TYPE_COPY_ACCELERATION_STRUCTURE_INFO_KHR};
        copy.src = as.handle;
        copy.dst = compact.handle;
        copy.mode = VK_COPY_ACCELERATION_STRUCTURE_MODE_COMPACT_KHR;
        VkCommandBuffer copyCmd = ctx_->beginOneShot();
        vkCmdCopyAccelerationStructureKHR(copyCmd, &copy);
        ctx_->endOneShot(copyCmd);
        // 形がたくさんあるとログが長くなるので、大きいものだけ書く（合計は buildAccelerationStructures が書く）
        if (triangleCount >= 100000)
            logInfo("BLAS: %u triangles, %.1f MB -> %.1f MB (compacted)", triangleCount,
                    sizes.accelerationStructureSize / 1048576.0, compactSize / 1048576.0);
        destroyAccel(as);
        as = compact;
    } else {
        if (triangleCount >= 100000)
            logInfo("BLAS: %u triangles, %.1f MB", triangleCount, sizes.accelerationStructureSize / 1048576.0);
    }

    VkAccelerationStructureDeviceAddressInfoKHR ai{VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_DEVICE_ADDRESS_INFO_KHR};
    ai.accelerationStructure = as.handle;
    as.address = vkGetAccelerationStructureDeviceAddressKHR(ctx_->device, &ai);
    return as;
}

void Renderer::buildTlas(const std::vector<VkAccelerationStructureInstanceKHR>& instances) {
    instances_ = ctx_->createDeviceBuffer(instances.data(), instances.size() * sizeof(VkAccelerationStructureInstanceKHR),
                                          VK_BUFFER_USAGE_SHADER_DEVICE_ADDRESS_BIT |
                                              VK_BUFFER_USAGE_ACCELERATION_STRUCTURE_BUILD_INPUT_READ_ONLY_BIT_KHR);

    VkAccelerationStructureGeometryKHR geom{VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_GEOMETRY_KHR};
    geom.geometryType = VK_GEOMETRY_TYPE_INSTANCES_KHR;
    geom.geometry.instances.sType = VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_GEOMETRY_INSTANCES_DATA_KHR;
    geom.geometry.instances.data.deviceAddress = instances_.address;

    VkAccelerationStructureBuildGeometryInfoKHR build{VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_BUILD_GEOMETRY_INFO_KHR};
    build.type = VK_ACCELERATION_STRUCTURE_TYPE_TOP_LEVEL_KHR;
    build.flags = VK_BUILD_ACCELERATION_STRUCTURE_PREFER_FAST_TRACE_BIT_KHR;
    build.mode = VK_BUILD_ACCELERATION_STRUCTURE_MODE_BUILD_KHR;
    build.geometryCount = 1;
    build.pGeometries = &geom;

    uint32_t count = static_cast<uint32_t>(instances.size());
    VkAccelerationStructureBuildSizesInfoKHR sizes{VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_BUILD_SIZES_INFO_KHR};
    vkGetAccelerationStructureBuildSizesKHR(ctx_->device, VK_ACCELERATION_STRUCTURE_BUILD_TYPE_DEVICE_KHR, &build, &count, &sizes);

    tlas_.buffer = ctx_->createBuffer(sizes.accelerationStructureSize,
                                      VK_BUFFER_USAGE_ACCELERATION_STRUCTURE_STORAGE_BIT_KHR | VK_BUFFER_USAGE_SHADER_DEVICE_ADDRESS_BIT,
                                      VK_MEMORY_PROPERTY_DEVICE_LOCAL_BIT);
    VkAccelerationStructureCreateInfoKHR ci{VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_CREATE_INFO_KHR};
    ci.buffer = tlas_.buffer.buffer;
    ci.size = sizes.accelerationStructureSize;
    ci.type = VK_ACCELERATION_STRUCTURE_TYPE_TOP_LEVEL_KHR;
    vkCheck(vkCreateAccelerationStructureKHR(ctx_->device, &ci, nullptr, &tlas_.handle), "vkCreateAccelerationStructureKHR (TLAS)");

    VkPhysicalDeviceAccelerationStructurePropertiesKHR asProps{VK_STRUCTURE_TYPE_PHYSICAL_DEVICE_ACCELERATION_STRUCTURE_PROPERTIES_KHR};
    VkPhysicalDeviceProperties2 props2{VK_STRUCTURE_TYPE_PHYSICAL_DEVICE_PROPERTIES_2};
    props2.pNext = &asProps;
    vkGetPhysicalDeviceProperties2(ctx_->physical, &props2);
    VkDeviceSize align = asProps.minAccelerationStructureScratchOffsetAlignment;
    Buffer scratch = ctx_->createBuffer(sizes.buildScratchSize + align,
                                        VK_BUFFER_USAGE_STORAGE_BUFFER_BIT | VK_BUFFER_USAGE_SHADER_DEVICE_ADDRESS_BIT,
                                        VK_MEMORY_PROPERTY_DEVICE_LOCAL_BIT);
    build.dstAccelerationStructure = tlas_.handle;
    build.scratchData.deviceAddress = alignUp(scratch.address, align);
    VkAccelerationStructureBuildRangeInfoKHR range{count, 0, 0, 0};
    const VkAccelerationStructureBuildRangeInfoKHR* ranges = &range;
    VkCommandBuffer cmd = ctx_->beginOneShot();
    vkCmdBuildAccelerationStructuresKHR(cmd, 1, &build, &ranges);
    ctx_->endOneShot(cmd);
    ctx_->destroyBuffer(scratch);
}

void Renderer::buildAccelerationStructures(const Scene& scene) {
    const uint32_t triCount = static_cast<uint32_t>(scene.triangleCount());
    const uint32_t vertexCount = static_cast<uint32_t>(scene.vertices.size());
    std::vector<VkAccelerationStructureInstanceKHR> instances;

    auto makeInstance = [](VkDeviceAddress blas, uint32_t customIndex, uint32_t mask) {
        VkAccelerationStructureInstanceKHR inst{};
        inst.transform.matrix[0][0] = inst.transform.matrix[1][1] = inst.transform.matrix[2][2] = 1.0f;
        inst.instanceCustomIndex = customIndex;
        inst.mask = mask;
        inst.flags = VK_GEOMETRY_INSTANCE_TRIANGLE_FACING_CULL_DISABLE_BIT_KHR;
        inst.accelerationStructureReference = blas;
        return inst;
    };

    // 形ごとに、不透明な部分とガラスの部分の加速構造を 1 つずつ作る（何度置いても 1 回だけ）。
    // 置き場所ごとに TLAS のインスタンスを作る: 印（マスク）は不透明 0x01・ガラス 0x02（影の計算は 0x01 だけを見るので、ガラスは光を通す）。
    // 番号（custom index）は「置き場所の番号 × 2 + ガラスなら 1」。シェーダーはこれで形と三角形を見つける
    std::vector<int32_t> opaqueBlas(scene.meshes.size(), -1), glassBlas(scene.meshes.size(), -1);
    for (size_t mi = 0; mi < scene.meshes.size(); ++mi) {
        const MeshRange& m = scene.meshes[mi];
        if (m.opaqueCount > 0) {
            opaqueBlas[mi] = static_cast<int32_t>(blas_.size());
            blas_.push_back(buildBlas(static_cast<VkDeviceSize>(m.firstTri) * 3 * sizeof(uint32_t), m.opaqueCount, vertexCount));
        }
        if (m.triCount > m.opaqueCount) {
            glassBlas[mi] = static_cast<int32_t>(blas_.size());
            blas_.push_back(buildBlas(static_cast<VkDeviceSize>(m.firstTri + m.opaqueCount) * 3 * sizeof(uint32_t),
                                      m.triCount - m.opaqueCount, vertexCount));
        }
    }
    if (scene.instances.size() >= (1u << 23)) fatal("%s\n(placements %zu)", tr("ModelTooLarge"), scene.instances.size());
    for (size_t i = 0; i < scene.instances.size(); ++i) {
        const Instance& in = scene.instances[i];
        auto add = [&](int32_t b, uint32_t glass, uint32_t mask) {
            if (b < 0) return;
            VkAccelerationStructureInstanceKHR inst = makeInstance(blas_[b].address, static_cast<uint32_t>(i) * 2u + glass, mask);
            for (int r = 0; r < 3; ++r)
                for (int c = 0; c < 4; ++c) inst.transform.matrix[r][c] = in.m[r * 4 + c];
            instances.push_back(inst);
        };
        add(opaqueBlas[in.mesh], 0u, 0x01);
        add(glassBlas[in.mesh], 1u, 0x02);
    }
    (void)triCount;
    VkDeviceSize blasBytes = 0;
    for (const AccelStruct& b : blas_) blasBytes += b.buffer.allocated;
    logInfo("BLAS: %zu (%.1f MB), placements in TLAS: %zu", blas_.size(), blasBytes / 1048576.0, instances.size());
    buildTlas(instances);
    logInfo("Acceleration structures built");
}

void Renderer::destroyAccel(AccelStruct& a) {
    if (a.handle) vkDestroyAccelerationStructureKHR(ctx_->device, a.handle, nullptr);
    ctx_->destroyBuffer(a.buffer);
    a = AccelStruct{};
}


// --- シェーダーと入出力のつなぎ ---------------------------------------------------

void Renderer::createPasses() {
    VkDevice dev = ctx_->device;
    const VkDescriptorType IMG = VK_DESCRIPTOR_TYPE_STORAGE_IMAGE, BUF = VK_DESCRIPTOR_TYPE_STORAGE_BUFFER;

    // パストレース: 0 = 形（加速構造 / BVH の箱）, 1 = 明るさ, 2 = 位置, 3〜6 = 形・材質, 7 = 面の向き, 8 = BVH の三角形, 9 = 面の色,
    //               10 = テクスチャの置き場所, 11 = テクスチャの画素, 12 = 照明器具の光源, 13 = 空と雲の画像
    std::vector<BindingSpec> pt = {
        {0, hardwareRt_ ? VK_DESCRIPTOR_TYPE_ACCELERATION_STRUCTURE_KHR : BUF},
        {1, IMG}, {2, IMG}, {3, BUF}, {4, BUF}, {5, BUF}, {6, BUF}, {7, IMG}, {9, IMG}, {10, BUF}, {11, BUF}, {12, BUF},
        {13, IMG}, {14, BUF}, {15, BUF}, {16, BUF},
    };
    if (!hardwareRt_) pt.push_back({8, BUF});
    pathTrace_ = hardwareRt_
        ? createPass(dev, pt, 2, sizeof(PathTracePush), kPathTraceHwSpv, sizeof(kPathTraceHwSpv), "path trace (hw)")
        : createPass(dev, pt, 2, sizeof(PathTracePush), kPathTraceSwSpv, sizeof(kPathTraceSwSpv), "path trace (sw)");

    std::vector<BindingSpec> temporal;
    for (uint32_t i = 0; i < 13; ++i) temporal.push_back({i, IMG});
    temporal_ = createPass(dev, temporal, 2, sizeof(TemporalPush), kTemporalSpv, sizeof(kTemporalSpv), "temporal");

    std::vector<BindingSpec> atrous;
    for (uint32_t i = 0; i < 5; ++i) atrous.push_back({i, IMG});
    atrous_ = createPass(dev, atrous, 4, sizeof(AtrousPush), kAtrousSpv, sizeof(kAtrousSpv), "atrous");

    localDown_ = createPass(dev, {{0, IMG}, {1, IMG}, {2, IMG}, {3, BUF}}, 2, sizeof(LocalDownPush), kLocalDownSpv,
                            sizeof(kLocalDownSpv), "local exposure (down)");
    localBlur_ = createPass(dev, {{0, IMG}, {1, IMG}}, 2, sizeof(LocalBlurPush), kLocalBlurSpv, sizeof(kLocalBlurSpv),
                            "local exposure (blur)");
    composite_ = createPass(dev, {{0, IMG}, {1, IMG}, {2, IMG}, {3, IMG}, {4, IMG}}, 2, sizeof(CompositePush),
                            kCompositeSpv, sizeof(kCompositeSpv), "composite");
    skyPass_ = createPass(dev, {{0, IMG}}, 1, sizeof(SkyPush), kSkySpv, sizeof(kSkySpv), "sky");
    skyCdfPass_ = createPass(dev, {{0, IMG}, {1, BUF}}, 1, 4, kSkyCdfSpv, sizeof(kSkyCdfSpv), "sky distribution");
    overlayPass_ = createPass(dev, {{0, IMG}, {1, IMG}}, kOverlaySlots, sizeof(OverlayPush), kOverlaySpv, sizeof(kOverlaySpv), "overlay");
}

void Renderer::writeDescriptors() {
    VkDevice dev = ctx_->device;
    for (int cur = 0; cur < 2; ++cur) {
        int prev = 1 - cur;
        std::vector<Bind> pt = {
            {1, noisy_.view}, {2, gpos_[cur].view}, {3, {}, vertices_.buffer}, {4, {}, indices_.buffer},
            {5, {}, triMaterials_.buffer}, {6, {}, materials_.buffer}, {7, gnormal_[cur].view}, {9, galbedo_.view},
            {10, {}, texInfos_.buffer}, {11, {}, texels_.buffer}, {12, {}, lights_.buffer}, {13, skyImage_.view},
            {14, {}, sceneInstances_.buffer}, {15, {}, meshes_.buffer}, {16, {}, skyDist_.buffer},
        };
        if (hardwareRt_) {
            Bind b{0};
            b.accel = &tlas_.handle;
            pt.push_back(b);
        } else {
            pt.push_back({0, {}, bvhNodes_.buffer});
            pt.push_back({8, {}, bvhTris_.buffer});
        }
        writeSet(dev, pathTrace_.sets[cur], pt);

        writeSet(dev, temporal_.sets[cur], {
            {0, noisy_.view}, {1, gpos_[cur].view}, {2, gnormal_[cur].view}, {3, gpos_[prev].view},
            {4, gnormal_[prev].view}, {5, history_[prev].view}, {6, moments_[prev].view}, {7, history_[cur].view},
            {8, moments_[cur].view}, {9, filtered_[0].view}, {10, galbedo_.view}, {11, albedoHist_[prev].view},
            {12, albedoHist_[cur].view},
        });
        for (int dir = 0; dir < 2; ++dir) {
            writeSet(dev, atrous_.sets[cur * 2 + dir], {
                {0, filtered_[dir].view}, {1, filtered_[1 - dir].view}, {2, gpos_[cur].view},
                {3, gnormal_[cur].view}, {4, history_[cur].view},
            });
        }
        // 平均を奇数回くり返すと、結果は filtered_[1] に入る
        writeSet(dev, localDown_.sets[cur], {
            {0, filtered_[kAtrousIterations % 2].view}, {1, albedoHist_[cur].view}, {2, local_[0].view},
            {3, {}, exposure_.buffer},
        });
        // ぼかし: 横は local_[0] → local_[1]、縦は local_[1] → local_[0]
        writeSet(dev, localBlur_.sets[cur], {{0, local_[cur].view}, {1, local_[1 - cur].view}});
        writeSet(dev, composite_.sets[cur], {
            {0, filtered_[kAtrousIterations % 2].view}, {1, history_[cur].view}, {2, albedoHist_[cur].view},
            {3, output_.view}, {4, local_[0].view},
        });
    }
}

void Renderer::writeOverlayDescriptors() {
    if (!overlayTargetView_) return;
    for (uint32_t i = 0; i < kOverlaySlots; ++i)
        if (overlay_[i].view) writeSet(ctx_->device, overlayPass_.sets[i], {{0, overlayTargetView_}, {1, overlay_[i].view}});
}

void Renderer::setOverlayTarget(const Image& target) {
    overlayTargetView_ = target.view;
    writeOverlayDescriptors();
}

void Renderer::setOverlay(uint32_t slot, const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height) {
    if (slot >= kOverlaySlots) return;
    vkDeviceWaitIdle(ctx_->device);
    Image& overlay = overlay_[slot];
    ctx_->destroyImage(overlay);
    if (width == 0 || height == 0 || rgba.size() < static_cast<size_t>(width) * height * 4) return;
    overlay = ctx_->createImage(width, height, VK_FORMAT_R8G8B8A8_UNORM,
                                 VK_IMAGE_USAGE_STORAGE_BIT | VK_IMAGE_USAGE_TRANSFER_DST_BIT);
    Buffer staging = ctx_->createBuffer(rgba.size(), VK_BUFFER_USAGE_TRANSFER_SRC_BIT,
                                        VK_MEMORY_PROPERTY_HOST_VISIBLE_BIT | VK_MEMORY_PROPERTY_HOST_COHERENT_BIT);
    memcpy(staging.mapped, rgba.data(), rgba.size());
    VkCommandBuffer cmd = ctx_->beginOneShot();
    transitionImage(cmd, overlay.image, VK_IMAGE_LAYOUT_UNDEFINED, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL,
                    VK_PIPELINE_STAGE_TOP_OF_PIPE_BIT, 0, VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_WRITE_BIT);
    VkBufferImageCopy region{};
    region.imageSubresource = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 0, 1};
    region.imageExtent = {width, height, 1};
    vkCmdCopyBufferToImage(cmd, staging.buffer, overlay.image, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, 1, &region);
    transitionImage(cmd, overlay.image, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, VK_IMAGE_LAYOUT_GENERAL,
                    VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_WRITE_BIT,
                    VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_ACCESS_SHADER_READ_BIT);
    ctx_->endOneShot(cmd);
    ctx_->destroyBuffer(staging);
    writeOverlayDescriptors();
}

void Renderer::recordOverlay(VkCommandBuffer cmd, uint32_t slot, int x, int y, float opacity) {
    if (slot >= kOverlaySlots || !overlay_[slot].view || !overlayTargetView_) return;
    // target は GENERAL レイアウトで、書き込みが終わっている前提（呼び出し側で切り替える）。
    // 続けて重ねるときは、前の重ね合わせの書き込みを待つ
    computeBarrier(cmd);
    OverlayPush pc{{x, y}, opacity};
    dispatch(cmd, overlayPass_, slot, &pc, sizeof(pc), overlay_[slot].width, overlay_[slot].height);
}

void Renderer::setLook(uint32_t background, uint32_t mono, bool hideGround, uint32_t ground) {
    uint32_t flags = hideGround ? 1u : 0u;
    if (background == lookBackground_ && mono == lookMono_ && flags == lookFlags_ && ground == lookGround_) return;
    lookBackground_ = background;
    lookMono_ = mono;
    lookFlags_ = flags;
    lookGround_ = ground;
    resetHistory_ = true;  // 見た目が変わったので、前のフレームを重ねない
}

void Renderer::setSunDirection(Vec3 dir) {
    if (std::fabs(dir.x - sunDir_.x) + std::fabs(dir.y - sunDir_.y) + std::fabs(dir.z - sunDir_.z) < 1e-6f) return;
    sunDir_ = dir;
    computeSky();
    resetHistory_ = true;
}

void Renderer::setLighting(bool sun, bool artificial) {
    if (sun == sunOn_ && artificial == lightsOn_) return;
    bool skyChanged = sun != sunOn_;
    sunOn_ = sun;
    lightsOn_ = artificial;
    if (skyChanged) computeSky();
    resetHistory_ = true;
}

void Renderer::setCloudAmount(float amount) {
    amount = std::max(0.0f, std::min(1.0f, amount));
    if (std::fabs(amount - cloudAmount_) < 1e-4f) return;
    cloudAmount_ = amount;
    computeSky();
    resetHistory_ = true;
}

void Renderer::setTransparentOutput(bool transparent) {
    if (transparent == transparentOutput_) return;
    transparentOutput_ = transparent;
    resetHistory_ = true;
}

void Renderer::resize(uint32_t width, uint32_t height) {
    vkDeviceWaitIdle(ctx_->device);
    Image* all[] = {&noisy_, &gpos_[0], &gpos_[1], &gnormal_[0], &gnormal_[1], &galbedo_, &history_[0], &history_[1],
                    &moments_[0], &moments_[1], &filtered_[0], &filtered_[1], &albedoHist_[0], &albedoHist_[1], &output_,
                    &local_[0], &local_[1]};
    for (Image* img : all) ctx_->destroyImage(*img);
    const VkImageUsageFlags storage = VK_IMAGE_USAGE_STORAGE_BIT;
    for (Image* img : {&local_[0], &local_[1]})
        *img = ctx_->createImage((width + 7) / 8, (height + 7) / 8, VK_FORMAT_R32G32B32A32_SFLOAT, storage);
    for (Image* img : {&gpos_[0], &gpos_[1], &history_[0], &history_[1],
                       &moments_[0], &moments_[1], &filtered_[0], &filtered_[1], &albedoHist_[0], &albedoHist_[1]})
        *img = ctx_->createImage(width, height, VK_FORMAT_R32G32B32A32_SFLOAT, storage);
    // 精度の要らないもの（今回の明るさ・面の向き）は半分の大きさの数（16 ビット）で持ち、GPU のメモリを節約する。
    // ノイズ除去の途中の値（ばらつき）は大きな値になるので 32 ビットのまま
    for (Image* img : {&noisy_, &gnormal_[0], &gnormal_[1]})
        *img = ctx_->createImage(width, height, VK_FORMAT_R16G16B16A16_SFLOAT, storage);
    galbedo_ = ctx_->createImage(width, height, VK_FORMAT_R8G8B8A8_UNORM, storage);
    output_ = ctx_->createImage(width, height, VK_FORMAT_R8G8B8A8_UNORM, storage | VK_IMAGE_USAGE_TRANSFER_SRC_BIT);
    imagesFresh_ = true;
    resetHistory_ = true;
    writeDescriptors();
}

VkDeviceSize Renderer::imageBytes(uint32_t width, uint32_t height) {
    // resize と同じ並び: 32 ビット×4 が 10 枚、16 ビット×4 が 3 枚、8 ビット×4 が 2 枚、縦横 1/8 の 32 ビット×4 が 2 枚
    const VkDeviceSize px = static_cast<VkDeviceSize>(width) * height;
    const VkDeviceSize small = static_cast<VkDeviceSize>((width + 7) / 8) * ((height + 7) / 8);
    return px * (10 * 16 + 3 * 8 + 2 * 4) + small * 2 * 16;
}

void Renderer::dispatch(VkCommandBuffer cmd, const ComputePass& pass, uint32_t set, const void* push, uint32_t pushSize,
                        uint32_t width, uint32_t height) {
    vkCmdBindPipeline(cmd, VK_PIPELINE_BIND_POINT_COMPUTE, pass.pipeline);
    vkCmdBindDescriptorSets(cmd, VK_PIPELINE_BIND_POINT_COMPUTE, pass.pipelineLayout, 0, 1, &pass.sets[set], 0, nullptr);
    vkCmdPushConstants(cmd, pass.pipelineLayout, VK_SHADER_STAGE_COMPUTE_BIT, 0, pushSize, push);
    vkCmdDispatch(cmd, (width + 7) / 8, (height + 7) / 8, 1);
}

void Renderer::record(VkCommandBuffer cmd, const ViewParams& view) {
    // 帯に分けて描くときは、帯ごとに先に GPU へ送って終わるのを待つ（cmd より前に実行される）ので、
    // 作り直した画像の準備もそれより前に済ませる
    const uint32_t bands = std::max(1u, std::min(ptBands_, (output_.height + 7) / 8));
    if (imagesFresh_) {
        VkCommandBuffer prep = bands > 1 ? ctx_->beginOneShot() : cmd;
        Image* all[] = {&noisy_, &gpos_[0], &gpos_[1], &gnormal_[0], &gnormal_[1], &galbedo_, &history_[0],
                        &history_[1], &moments_[0], &moments_[1], &filtered_[0], &filtered_[1], &albedoHist_[0], &albedoHist_[1], &output_,
                        &local_[0], &local_[1]};
        for (Image* img : all)
            transitionImage(prep, img->image, VK_IMAGE_LAYOUT_UNDEFINED, VK_IMAGE_LAYOUT_GENERAL,
                            VK_PIPELINE_STAGE_TOP_OF_PIPE_BIT, 0, VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT,
                            VK_ACCESS_SHADER_READ_BIT | VK_ACCESS_SHADER_WRITE_BIT);
        if (bands > 1) ctx_->endOneShot(prep);
        imagesFresh_ = false;
    }
    memset(exposure_.mapped, 0, sizeof(ExposureData));

    const bool reset = resetHistory_ || !hasLastView_;
    const bool moved = !reset && !sameView(view, lastView_);
    if (reset || moved) staticFrames_ = 0;
    const uint32_t cur = frameIndex_ & 1u;
    const uint32_t w = output_.width, h = output_.height;
    const float tanHalf = std::tan(view.fovY * 0.5f * 3.14159265f / 180.0f);
    const float aspect = static_cast<float>(w) / static_cast<float>(h);

    // 1. パストレース
    PathTracePush pt{};
    pt.camPos[0] = view.eye.x; pt.camPos[1] = view.eye.y; pt.camPos[2] = view.eye.z; pt.camPos[3] = tanHalf;
    pt.camForward[0] = view.forward.x; pt.camForward[1] = view.forward.y; pt.camForward[2] = view.forward.z;
    pt.camForward[3] = view.panorama ? 1.0f : 0.0f;
    pt.camRight[0] = view.right.x; pt.camRight[1] = view.right.y; pt.camRight[2] = view.right.z; pt.camRight[3] = aspect;
    pt.camUp[0] = view.up.x; pt.camUp[1] = view.up.y; pt.camUp[2] = view.up.z; pt.camUp[3] = groundZ_;
    pt.sunDir[0] = sunDir_.x; pt.sunDir[1] = sunDir_.y; pt.sunDir[2] = sunDir_.z; pt.sunDir[3] = effectiveSun();
    pt.misc[0] = lightsOn_ ? static_cast<float>(lightCount_) : 0.0f;
    pt.misc[1] = effectiveSky();
    pt.misc[2] = skySampling_ ? 1.0f : 0.0f;  // 空と地面の光を直接狙う
    // 1 画素が 1 m 先で何 m の幅になるか（テクスチャの縮小した段を選ぶのに使う）
    pt.misc[3] = view.panorama ? 2.0f * 3.14159265f / static_cast<float>(w) : 2.0f * tanHalf / static_cast<float>(h);
    pt.frame[0] = frameIndex_;
    pt.frame[1] = staticFrames_;
    pt.frame[2] = view.maxBounces;
    pt.frame[3] = std::min(view.samplesPerFrame, 255u);
    // 透明にするときは、空が見える所を「背景」として印を付ける（色は使わない）
    pt.look[0] = transparentOutput_ ? (lookBackground_ | 0xFF000000u) : lookBackground_;
    pt.look[1] = lookMono_;
    pt.look[2] = lookFlags_;
    pt.look[3] = lookGround_;
    if (bands <= 1) {
        dispatch(cmd, pathTrace_, cur, &pt, sizeof(pt), w, h);
    } else {
        // 重い 1 枚を横長の帯に分け、帯ごとに GPU へ送って終わるのを待つ（1 回の処理が長いと Windows が GPU を
        // リセットして落ちるため。約 2 秒）。最後の帯だけ cmd に入れ、続く処理と一緒に送る
        const uint32_t rowsPerBand = ((h + bands - 1) / bands + 7) / 8 * 8;
        uint32_t y = 0;
        while (y < h) {
            const uint32_t rows = std::min(rowsPerBand, h - y);
            PathTracePush band = pt;
            band.frame[3] = pt.frame[3] | (y << 8);
            if (y + rows < h) {
                VkCommandBuffer one = ctx_->beginOneShot();
                computeBarrier(one);
                dispatch(one, pathTrace_, cur, &band, sizeof(band), w, rows);
                ctx_->endOneShot(one);
            } else {
                computeBarrier(cmd);  // 先に送った帯の書き込みを、この後の処理から見えるようにする
                dispatch(cmd, pathTrace_, cur, &band, sizeof(band), w, rows);
            }
            y += rows;
        }
    }
    computeBarrier(cmd);

    // 2. 前のフレームを重ねる（前のカメラの位置で、今の位置を写し直す）
    const ViewParams& prev = hasLastView_ ? lastView_ : view;
    TemporalPush tp{};
    tp.prevEye[0] = prev.eye.x; tp.prevEye[1] = prev.eye.y; tp.prevEye[2] = prev.eye.z; tp.prevEye[3] = moved ? 1.0f : 0.0f;
    tp.prevForward[0] = prev.forward.x; tp.prevForward[1] = prev.forward.y; tp.prevForward[2] = prev.forward.z;
    tp.prevForward[3] = reset ? 1.0f : 0.0f;
    tp.prevRight[0] = prev.right.x; tp.prevRight[1] = prev.right.y; tp.prevRight[2] = prev.right.z; tp.prevRight[3] = aspect;
    tp.prevUp[0] = prev.up.x; tp.prevUp[1] = prev.up.y; tp.prevUp[2] = prev.up.z;
    tp.prevUp[3] = std::tan(prev.fovY * 0.5f * 3.14159265f / 180.0f);
    tp.params[0] = kMovingHistoryAlpha;
    tp.params[1] = kMaxHistory;
    dispatch(cmd, temporal_, cur, &tp, sizeof(tp), w, h);
    computeBarrier(cmd);

    // 3. 周りの画素と平均する（間隔を広げながらくり返す）
    for (int i = 0; i < kAtrousIterations; ++i) {
        AtrousPush ap{spatialDenoise_ ? 1 << i : 0};
        dispatch(cmd, atrous_, cur * 2 + (i % 2), &ap, sizeof(ap), w, h);
        computeBarrier(cmd);
    }

    // 4. 大まかな明るさ（ローカルトーンマッピング用。自動露出の集計も）
    LocalDownPush ldp{lookBackground_};
    dispatch(cmd, localDown_, cur, &ldp, sizeof(ldp), w, h);
    computeBarrier(cmd);
    const uint32_t sw = local_[0].width, sh = local_[0].height;
    LocalBlurPush lbx{{1, 0}}, lby{{0, 1}};
    dispatch(cmd, localBlur_, 0, &lbx, sizeof(lbx), sw, sh);
    computeBarrier(cmd);
    dispatch(cmd, localBlur_, 1, &lby, sizeof(lby), sw, sh);
    computeBarrier(cmd);

    // 5. 仕上げ。縮める中心は前のフレームで測った画面の平均（まだ無ければ露出から逆算）
    float pivot = hasAvgLog2_ ? avgLog2_ : std::log2(0.32f / std::max(view.exposure, 1e-6f));
    // 明るい所の収め先: 平均との差が小さい画面（普通の屋外など）はそのまま。差が大きい画面（屋根の下から外を見る）ほど強く縮め、
    // kBrightTarget 段 + 超えた分の 1 割に収める（外が真っ白にならず、明るい外として見える）
    const float kBrightTarget = 1.6f;
    float brightIn = hasAvgLog2_ ? std::max(brightLog2_ - avgLog2_, 0.0f) : 0.0f;
    float brightOut = std::min(brightIn, kBrightTarget + 0.1f * std::max(brightIn - kBrightTarget, 0.0f));
    CompositePush cp{view.exposure, lookBackground_, transparentOutput_ ? 1u : 0u, pivot, localStrength_, brightIn, brightOut};
    dispatch(cmd, composite_, cur, &cp, sizeof(cp), w, h);

    // 自動露出の集計を CPU から読めるようにする
    VkMemoryBarrier toHost{VK_STRUCTURE_TYPE_MEMORY_BARRIER};
    toHost.srcAccessMask = VK_ACCESS_SHADER_WRITE_BIT;
    toHost.dstAccessMask = VK_ACCESS_HOST_READ_BIT;
    vkCmdPipelineBarrier(cmd, VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_PIPELINE_STAGE_HOST_BIT, 0, 1, &toHost, 0,
                         nullptr, 0, nullptr);

    lastView_ = view;
    hasLastView_ = true;
    resetHistory_ = false;
    staticFrames_++;
    frameIndex_++;
}

float Renderer::finishFrame() {
    const auto* data = static_cast<const ExposureData*>(exposure_.mapped);
    if (data->logLumCount <= 0) return -1.0f;
    float avgLog2 = static_cast<float>(data->logLumSum) / 64.0f / static_cast<float>(data->logLumCount);
    // 明るい所: 度数分布の上位 5% の明るさ（区分の中で数に比例して補間し、段差が出ないようにする）
    int64_t total = 0;
    for (int i = 0; i < 64; ++i) total += data->histogram[i];
    float bright = avgLog2;
    int64_t above = 0;
    const double want = total * 0.05;
    for (int i = 63; i >= 0; --i) {
        int64_t n = data->histogram[i];
        if (n > 0 && above + n >= want) {
            float inBin = static_cast<float>((want - above) / static_cast<double>(n));  // 区分の上の端から何割下か
            bright = (i + 1 - inBin) * 0.5f - 16.0f;
            break;
        }
        above += n;
    }
    // 歩き回っている間に外の明るさがちらつかないよう、前のフレームの値となめらかにつなぐ
    if (!hasAvgLog2_) {
        avgLog2_ = avgLog2;
        brightLog2_ = bright;
    } else {
        avgLog2_ += (avgLog2 - avgLog2_) * 0.2f;
        brightLog2_ += (bright - brightLog2_) * 0.2f;
    }
    hasAvgLog2_ = true;
    return std::exp2(avgLog2);
}

void Renderer::destroy() {
    if (!ctx_) return;
    VkDevice dev = ctx_->device;
    vkDeviceWaitIdle(dev);
    for (ComputePass* p : {&pathTrace_, &temporal_, &atrous_, &localDown_, &localBlur_, &composite_, &skyPass_, &skyCdfPass_, &overlayPass_})
        destroyPass(dev, *p);
    Image* all[] = {&noisy_, &gpos_[0], &gpos_[1], &gnormal_[0], &gnormal_[1], &galbedo_, &history_[0], &history_[1],
                    &moments_[0], &moments_[1], &filtered_[0], &filtered_[1], &albedoHist_[0], &albedoHist_[1], &output_,
                    &local_[0], &local_[1], &skyImage_, &overlay_[0], &overlay_[1], &overlay_[2], &overlay_[3]};
    for (Image* img : all) ctx_->destroyImage(*img);
    destroyScene();
    ctx_->destroyBuffer(exposure_);
    ctx_->destroyBuffer(skyDist_);
    ctx_ = nullptr;
}

}  // namespace t28

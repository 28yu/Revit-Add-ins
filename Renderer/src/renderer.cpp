#include "renderer.h"

#include <algorithm>
#include <cmath>
#include <cstring>

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
static const uint32_t kOverlaySpv[] =      // 操作の案内を重ねる
#include "overlay.spv.inc"
    ;

namespace {

const int kAtrousIterations = 5;           // 平均する回数（間隔 1, 2, 4, 8, 16 画素）
const float kMovingHistoryAlpha = 0.125f;  // 動いている間、今回の結果を混ぜる最小の割合（= 最大 8 フレーム分を重ねる）
const float kMaxHistory = 65536.0f;        // 止まっている間に重ねる回数の上限

struct PathTracePush {
    float camPos[4];
    float camForward[4];
    float camRight[4];
    float camUp[4];
    float sunDir[4];
    float misc[4];
    uint32_t frame[4];
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
};

struct OverlayPush {
    int32_t offset[2];
    float opacity;
};

struct ExposureData {
    int32_t logLumSum;
    int32_t logLumCount;
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

// テクスチャ 1 枚分の置き場所（シェーダーの TexInfo と同じ並び）
struct GpuTexInfo {
    uint32_t offset;  // texels の何番目から始まるか（一番大きい段）
    uint32_t width, height;
    uint32_t levels;  // 縮小した段の数（ミップマップ。遠くの模様がちらつかないよう、離れたら小さい段を使う）
};

float srgbToLinear(uint8_t c) {
    float s = c / 255.0f;
    return s <= 0.04045f ? s / 12.92f : std::pow((s + 0.055f) / 1.055f, 2.4f);
}

uint8_t linearToSrgb(float l) {
    l = std::min(std::max(l, 0.0f), 1.0f);
    float s = l <= 0.0031308f ? l * 12.92f : 1.055f * std::pow(l, 1.0f / 2.4f) - 0.055f;
    return static_cast<uint8_t>(s * 255.0f + 0.5f);
}

// 全テクスチャを 1 つの並びに詰める。段ごとに縦横半分にした画像も後ろに続けて入れる
void packTextures(const std::vector<Texture>& textures, std::vector<GpuTexInfo>& infos, std::vector<uint32_t>& texels) {
    for (const Texture& t : textures) {
        GpuTexInfo info{static_cast<uint32_t>(texels.size()), t.width, t.height, 0};
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
            for (size_t i = 0; i < static_cast<size_t>(w) * h; ++i) {
                uint32_t r = toByte(level[i * 4]), g = toByte(level[i * 4 + 1]), b = toByte(level[i * 4 + 2]);
                uint32_t a = static_cast<uint32_t>(std::min(std::max(level[i * 4 + 3], 0.0f), 1.0f) * 255.0f + 0.5f);
                texels.push_back(r | (g << 8) | (b << 16) | (a << 24));
            }
            ++info.levels;
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
    uploadScene(scene, glassStart, bvh);
    createPasses();
}

void Renderer::reloadScene(const Scene& scene, uint32_t glassStart, const Bvh& bvh) {
    vkDeviceWaitIdle(ctx_->device);
    destroyScene();
    uploadScene(scene, glassStart, bvh);
    if (output_.view) writeDescriptors();
    resetHistory_ = true;
}

void Renderer::destroyScene() {
    destroyAccel(tlas_);
    destroyAccel(blasOpaque_);
    destroyAccel(blasGlass_);
    for (Buffer* b : {&vertices_, &indices_, &triMaterials_, &materials_, &texInfos_, &texels_, &lights_, &instances_, &bvhNodes_, &bvhTris_})
        ctx_->destroyBuffer(*b);
}

void Renderer::uploadScene(const Scene& scene, uint32_t glassStart, const Bvh& bvh) {
    VkContext& ctx = *ctx_;
    glassStart_ = glassStart;
    sunDir_ = scene.sunDir;
    sunIntensity_ = scene.sunIntensity;
    skyIntensity_ = scene.skyIntensity;
    logInfo("Scene: %zu vertices, %zu triangles (%u opaque, %zu glass), %zu materials",
            scene.vertices.size(), scene.triangleCount(), glassStart_, scene.triangleCount() - glassStart_,
            scene.materials.size());

    // RT コアを使うときだけ、加速構造の材料として使う印を付ける（予備の方式ではその機能自体が無い）
    const VkBufferUsageFlags geomUsage =
        VK_BUFFER_USAGE_STORAGE_BUFFER_BIT |
        (hardwareRt_ ? VK_BUFFER_USAGE_SHADER_DEVICE_ADDRESS_BIT | VK_BUFFER_USAGE_ACCELERATION_STRUCTURE_BUILD_INPUT_READ_ONLY_BIT_KHR : 0);
    vertices_ = ctx.createDeviceBuffer(scene.vertices.data(), scene.vertices.size() * sizeof(Vertex), geomUsage);
    indices_ = ctx.createDeviceBuffer(scene.indices.data(), scene.indices.size() * sizeof(uint32_t), geomUsage);
    triMaterials_ = ctx.createDeviceBuffer(scene.triMaterials.data(), scene.triMaterials.size() * sizeof(uint32_t),
                                           VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);

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
    if (lights.empty()) lights.push_back({});  // 空のバッファは作れないので 1 つ入れておく（数は 0 のまま）
    lights_ = ctx.createDeviceBuffer(lights.data(), lights.size() * sizeof(GpuLight), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);
    logInfo("Lights: %u", lightCount_);

    std::vector<GpuTexInfo> texInfos;
    std::vector<uint32_t> texels;
    packTextures(scene.textures, texInfos, texels);
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
    tri.vertexStride = sizeof(Vertex);
    tri.maxVertex = vertexCount - 1;
    tri.indexType = VK_INDEX_TYPE_UINT32;
    tri.indexData.deviceAddress = indices_.address + indexOffsetBytes;

    VkAccelerationStructureBuildGeometryInfoKHR build{VK_STRUCTURE_TYPE_ACCELERATION_STRUCTURE_BUILD_GEOMETRY_INFO_KHR};
    build.type = VK_ACCELERATION_STRUCTURE_TYPE_BOTTOM_LEVEL_KHR;
    build.flags = VK_BUILD_ACCELERATION_STRUCTURE_PREFER_FAST_TRACE_BIT_KHR;
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
    VkCommandBuffer cmd = ctx_->beginOneShot();
    vkCmdBuildAccelerationStructuresKHR(cmd, 1, &build, &ranges);
    ctx_->endOneShot(cmd);
    ctx_->destroyBuffer(scratch);

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

    // 不透明な物（インスタンス 0, マスク 0x01）とガラス（インスタンス 1, マスク 0x02）に分ける。
    // 影の計算ではマスク 0x01 だけを見るので、ガラスは光を通す
    if (glassStart_ > 0) {
        blasOpaque_ = buildBlas(0, glassStart_, vertexCount);
        instances.push_back(makeInstance(blasOpaque_.address, 0, 0x01));
    }
    if (triCount > glassStart_) {
        blasGlass_ = buildBlas(static_cast<VkDeviceSize>(glassStart_) * 3 * sizeof(uint32_t), triCount - glassStart_, vertexCount);
        instances.push_back(makeInstance(blasGlass_.address, 1, 0x02));
    }
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
    //               10 = テクスチャの置き場所, 11 = テクスチャの画素, 12 = 照明器具の光源
    std::vector<BindingSpec> pt = {
        {0, hardwareRt_ ? VK_DESCRIPTOR_TYPE_ACCELERATION_STRUCTURE_KHR : BUF},
        {1, IMG}, {2, IMG}, {3, BUF}, {4, BUF}, {5, BUF}, {6, BUF}, {7, IMG}, {9, IMG}, {10, BUF}, {11, BUF}, {12, BUF},
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

    composite_ = createPass(dev, {{0, IMG}, {1, IMG}, {2, IMG}, {3, IMG}, {4, BUF}}, 2, sizeof(CompositePush),
                            kCompositeSpv, sizeof(kCompositeSpv), "composite");
    overlayPass_ = createPass(dev, {{0, IMG}, {1, IMG}}, kOverlaySlots, sizeof(OverlayPush), kOverlaySpv, sizeof(kOverlaySpv), "overlay");
}

void Renderer::writeDescriptors() {
    VkDevice dev = ctx_->device;
    for (int cur = 0; cur < 2; ++cur) {
        int prev = 1 - cur;
        std::vector<Bind> pt = {
            {1, noisy_.view}, {2, gpos_[cur].view}, {3, {}, vertices_.buffer}, {4, {}, indices_.buffer},
            {5, {}, triMaterials_.buffer}, {6, {}, materials_.buffer}, {7, gnormal_[cur].view}, {9, galbedo_.view},
            {10, {}, texInfos_.buffer}, {11, {}, texels_.buffer}, {12, {}, lights_.buffer},
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
        writeSet(dev, composite_.sets[cur], {
            {0, filtered_[kAtrousIterations % 2].view}, {1, history_[cur].view}, {2, albedoHist_[cur].view},
            {3, output_.view}, {4, {}, exposure_.buffer},
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

void Renderer::resize(uint32_t width, uint32_t height) {
    vkDeviceWaitIdle(ctx_->device);
    Image* all[] = {&noisy_, &gpos_[0], &gpos_[1], &gnormal_[0], &gnormal_[1], &galbedo_, &history_[0], &history_[1],
                    &moments_[0], &moments_[1], &filtered_[0], &filtered_[1], &albedoHist_[0], &albedoHist_[1], &output_};
    for (Image* img : all) ctx_->destroyImage(*img);
    const VkImageUsageFlags storage = VK_IMAGE_USAGE_STORAGE_BIT;
    for (Image* img : {&noisy_, &gpos_[0], &gpos_[1], &gnormal_[0], &gnormal_[1], &history_[0], &history_[1],
                       &moments_[0], &moments_[1], &filtered_[0], &filtered_[1], &albedoHist_[0], &albedoHist_[1]})
        *img = ctx_->createImage(width, height, VK_FORMAT_R32G32B32A32_SFLOAT, storage);
    galbedo_ = ctx_->createImage(width, height, VK_FORMAT_R8G8B8A8_UNORM, storage);
    output_ = ctx_->createImage(width, height, VK_FORMAT_R8G8B8A8_UNORM, storage | VK_IMAGE_USAGE_TRANSFER_SRC_BIT);
    imagesFresh_ = true;
    resetHistory_ = true;
    writeDescriptors();
}

void Renderer::dispatch(VkCommandBuffer cmd, const ComputePass& pass, uint32_t set, const void* push, uint32_t pushSize,
                        uint32_t width, uint32_t height) {
    vkCmdBindPipeline(cmd, VK_PIPELINE_BIND_POINT_COMPUTE, pass.pipeline);
    vkCmdBindDescriptorSets(cmd, VK_PIPELINE_BIND_POINT_COMPUTE, pass.pipelineLayout, 0, 1, &pass.sets[set], 0, nullptr);
    vkCmdPushConstants(cmd, pass.pipelineLayout, VK_SHADER_STAGE_COMPUTE_BIT, 0, pushSize, push);
    vkCmdDispatch(cmd, (width + 7) / 8, (height + 7) / 8, 1);
}

void Renderer::record(VkCommandBuffer cmd, const ViewParams& view) {
    if (imagesFresh_) {
        Image* all[] = {&noisy_, &gpos_[0], &gpos_[1], &gnormal_[0], &gnormal_[1], &galbedo_, &history_[0],
                        &history_[1], &moments_[0], &moments_[1], &filtered_[0], &filtered_[1], &albedoHist_[0], &albedoHist_[1], &output_};
        for (Image* img : all)
            transitionImage(cmd, img->image, VK_IMAGE_LAYOUT_UNDEFINED, VK_IMAGE_LAYOUT_GENERAL,
                            VK_PIPELINE_STAGE_TOP_OF_PIPE_BIT, 0, VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT,
                            VK_ACCESS_SHADER_READ_BIT | VK_ACCESS_SHADER_WRITE_BIT);
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
    pt.camUp[0] = view.up.x; pt.camUp[1] = view.up.y; pt.camUp[2] = view.up.z;
    pt.sunDir[0] = sunDir_.x; pt.sunDir[1] = sunDir_.y; pt.sunDir[2] = sunDir_.z; pt.sunDir[3] = sunIntensity_;
    pt.misc[0] = static_cast<float>(lightCount_);
    pt.misc[1] = skyIntensity_;
    pt.misc[2] = static_cast<float>(glassStart_);
    // 1 画素が 1 m 先で何 m の幅になるか（テクスチャの縮小した段を選ぶのに使う）
    pt.misc[3] = view.panorama ? 2.0f * 3.14159265f / static_cast<float>(w) : 2.0f * tanHalf / static_cast<float>(h);
    pt.frame[0] = frameIndex_;
    pt.frame[1] = staticFrames_;
    pt.frame[2] = view.maxBounces;
    pt.frame[3] = view.samplesPerFrame;
    dispatch(cmd, pathTrace_, cur, &pt, sizeof(pt), w, h);
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
        AtrousPush ap{1 << i};
        dispatch(cmd, atrous_, cur * 2 + (i % 2), &ap, sizeof(ap), w, h);
        computeBarrier(cmd);
    }

    // 4. 仕上げ
    CompositePush cp{view.exposure};
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
    return std::exp2(avgLog2);
}

void Renderer::destroy() {
    if (!ctx_) return;
    VkDevice dev = ctx_->device;
    vkDeviceWaitIdle(dev);
    for (ComputePass* p : {&pathTrace_, &temporal_, &atrous_, &composite_, &overlayPass_}) destroyPass(dev, *p);
    Image* all[] = {&noisy_, &gpos_[0], &gpos_[1], &gnormal_[0], &gnormal_[1], &galbedo_, &history_[0], &history_[1],
                    &moments_[0], &moments_[1], &filtered_[0], &filtered_[1], &albedoHist_[0], &albedoHist_[1], &output_, &overlay_[0], &overlay_[1], &overlay_[2]};
    for (Image* img : all) ctx_->destroyImage(*img);
    destroyScene();
    ctx_->destroyBuffer(exposure_);
    ctx_ = nullptr;
}

}  // namespace t28

#include "renderer.h"

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
static const uint32_t kOverlaySpv[] =      // 操作の案内を重ねる
#include "overlay.spv.inc"
    ;

namespace {
struct PushConstants {
    float camPos[4];
    float camForward[4];
    float camRight[4];
    float camUp[4];
    float sunDir[4];
    float misc[4];
    uint32_t frame[4];
};
static_assert(sizeof(PushConstants) <= 128, "push constants must fit in 128 bytes");

struct ExposureData {
    int32_t logLumSum;
    int32_t logLumCount;
};

struct OverlayPushConstants {
    int32_t offset[2];
    float opacity;
};

struct GpuMaterial {
    float color[4];
    float params[4];
};

VkDeviceSize alignUp(VkDeviceSize v, VkDeviceSize a) { return (v + a - 1) / a * a; }

VkShaderModule makeShader(VkDevice device, const uint32_t* code, size_t size) {
    VkShaderModuleCreateInfo smci{VK_STRUCTURE_TYPE_SHADER_MODULE_CREATE_INFO};
    smci.codeSize = size;
    smci.pCode = code;
    VkShaderModule module;
    vkCheck(vkCreateShaderModule(device, &smci, nullptr, &module), "vkCreateShaderModule");
    return module;
}

VkPipeline makeComputePipeline(VkDevice device, VkShaderModule module, VkPipelineLayout layout) {
    VkComputePipelineCreateInfo cpci{VK_STRUCTURE_TYPE_COMPUTE_PIPELINE_CREATE_INFO};
    cpci.stage.sType = VK_STRUCTURE_TYPE_PIPELINE_SHADER_STAGE_CREATE_INFO;
    cpci.stage.stage = VK_SHADER_STAGE_COMPUTE_BIT;
    cpci.stage.module = module;
    cpci.stage.pName = "main";
    cpci.layout = layout;
    VkPipeline pipeline;
    vkCheck(vkCreateComputePipelines(device, VK_NULL_HANDLE, 1, &cpci, nullptr, &pipeline), "vkCreateComputePipelines");
    vkDestroyShaderModule(device, module, nullptr);
    return pipeline;
}
}  // namespace

namespace {
// 予備の方式（BVH）で 0 番、ハード版で 0 番の種類が違うので、番号と種類の一覧を作って使う
struct BindingSpec {
    uint32_t binding;
    VkDescriptorType type;
};

std::vector<BindingSpec> pathTraceBindings(bool hardwareRt) {
    std::vector<BindingSpec> b = {
        {0, hardwareRt ? VK_DESCRIPTOR_TYPE_ACCELERATION_STRUCTURE_KHR : VK_DESCRIPTOR_TYPE_STORAGE_BUFFER},
        {1, VK_DESCRIPTOR_TYPE_STORAGE_IMAGE},  {2, VK_DESCRIPTOR_TYPE_STORAGE_IMAGE},
        {3, VK_DESCRIPTOR_TYPE_STORAGE_BUFFER}, {4, VK_DESCRIPTOR_TYPE_STORAGE_BUFFER},
        {5, VK_DESCRIPTOR_TYPE_STORAGE_BUFFER}, {6, VK_DESCRIPTOR_TYPE_STORAGE_BUFFER},
        {7, VK_DESCRIPTOR_TYPE_STORAGE_BUFFER},
    };
    if (!hardwareRt) b.push_back({8, VK_DESCRIPTOR_TYPE_STORAGE_BUFFER});
    return b;
}
}  // namespace

void Renderer::init(VkContext& ctx, const Scene& scene, uint32_t glassStart, const Bvh& bvh) {
    ctx_ = &ctx;
    hardwareRt_ = ctx.hardwareRt;
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
        mats.push_back({{m.r, m.g, m.b, m.opacity}, {m.roughness, m.metallic, m.emissive, 0.0f}});
    materials_ = ctx.createDeviceBuffer(mats.data(), mats.size() * sizeof(GpuMaterial), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);

    exposure_ = ctx.createBuffer(sizeof(ExposureData), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT,
                                 VK_MEMORY_PROPERTY_HOST_VISIBLE_BIT | VK_MEMORY_PROPERTY_HOST_COHERENT_BIT);
    memset(exposure_.mapped, 0, sizeof(ExposureData));

    if (hardwareRt_) {
        buildAccelerationStructures(scene);
    } else {
        bvhNodes_ = ctx.createDeviceBuffer(bvh.nodes().data(), bvh.nodes().size() * sizeof(BvhNode), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);
        bvhTris_ = ctx.createDeviceBuffer(bvh.triIndices().data(), bvh.triIndices().size() * sizeof(uint32_t), VK_BUFFER_USAGE_STORAGE_BUFFER_BIT);
        logInfo("BVH uploaded: %zu nodes", bvh.nodes().size());
    }
    createPipeline();
    createOverlayPipeline();
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

// --- パイプライン（シェーダーと入出力のつなぎ） --------------------------------

void Renderer::createPipeline() {
    std::vector<BindingSpec> specs = pathTraceBindings(hardwareRt_);
    std::vector<VkDescriptorSetLayoutBinding> bindings(specs.size());
    uint32_t imageCount = 0, bufferCount = 0, accelCount = 0;
    for (size_t i = 0; i < specs.size(); ++i) {
        bindings[i].binding = specs[i].binding;
        bindings[i].descriptorType = specs[i].type;
        bindings[i].descriptorCount = 1;
        bindings[i].stageFlags = VK_SHADER_STAGE_COMPUTE_BIT;
        if (specs[i].type == VK_DESCRIPTOR_TYPE_STORAGE_IMAGE) imageCount++;
        else if (specs[i].type == VK_DESCRIPTOR_TYPE_STORAGE_BUFFER) bufferCount++;
        else accelCount++;
    }
    VkDescriptorSetLayoutCreateInfo lci{VK_STRUCTURE_TYPE_DESCRIPTOR_SET_LAYOUT_CREATE_INFO};
    lci.bindingCount = static_cast<uint32_t>(bindings.size());
    lci.pBindings = bindings.data();
    vkCheck(vkCreateDescriptorSetLayout(ctx_->device, &lci, nullptr, &setLayout_), "vkCreateDescriptorSetLayout");

    std::vector<VkDescriptorPoolSize> poolSizes = {
        {VK_DESCRIPTOR_TYPE_STORAGE_IMAGE, imageCount},
        {VK_DESCRIPTOR_TYPE_STORAGE_BUFFER, bufferCount},
    };
    if (accelCount > 0) poolSizes.push_back({VK_DESCRIPTOR_TYPE_ACCELERATION_STRUCTURE_KHR, accelCount});
    VkDescriptorPoolCreateInfo pci{VK_STRUCTURE_TYPE_DESCRIPTOR_POOL_CREATE_INFO};
    pci.maxSets = 1;
    pci.poolSizeCount = static_cast<uint32_t>(poolSizes.size());
    pci.pPoolSizes = poolSizes.data();
    vkCheck(vkCreateDescriptorPool(ctx_->device, &pci, nullptr, &pool_), "vkCreateDescriptorPool");
    VkDescriptorSetAllocateInfo ai{VK_STRUCTURE_TYPE_DESCRIPTOR_SET_ALLOCATE_INFO};
    ai.descriptorPool = pool_;
    ai.descriptorSetCount = 1;
    ai.pSetLayouts = &setLayout_;
    vkCheck(vkAllocateDescriptorSets(ctx_->device, &ai, &set_), "vkAllocateDescriptorSets");

    VkPushConstantRange pcr{VK_SHADER_STAGE_COMPUTE_BIT, 0, sizeof(PushConstants)};
    VkPipelineLayoutCreateInfo plci{VK_STRUCTURE_TYPE_PIPELINE_LAYOUT_CREATE_INFO};
    plci.setLayoutCount = 1;
    plci.pSetLayouts = &setLayout_;
    plci.pushConstantRangeCount = 1;
    plci.pPushConstantRanges = &pcr;
    vkCheck(vkCreatePipelineLayout(ctx_->device, &plci, nullptr, &pipelineLayout_), "vkCreatePipelineLayout");

    VkShaderModule module = hardwareRt_ ? makeShader(ctx_->device, kPathTraceHwSpv, sizeof(kPathTraceHwSpv))
                                        : makeShader(ctx_->device, kPathTraceSwSpv, sizeof(kPathTraceSwSpv));
    pipeline_ = makeComputePipeline(ctx_->device, module, pipelineLayout_);
}

void Renderer::createOverlayPipeline() {
    VkDescriptorSetLayoutBinding bindings[2]{};
    for (uint32_t i = 0; i < 2; ++i) {
        bindings[i].binding = i;
        bindings[i].descriptorType = VK_DESCRIPTOR_TYPE_STORAGE_IMAGE;
        bindings[i].descriptorCount = 1;
        bindings[i].stageFlags = VK_SHADER_STAGE_COMPUTE_BIT;
    }
    VkDescriptorSetLayoutCreateInfo lci{VK_STRUCTURE_TYPE_DESCRIPTOR_SET_LAYOUT_CREATE_INFO};
    lci.bindingCount = 2;
    lci.pBindings = bindings;
    vkCheck(vkCreateDescriptorSetLayout(ctx_->device, &lci, nullptr, &overlayLayout_), "vkCreateDescriptorSetLayout (overlay)");

    VkDescriptorPoolSize poolSize{VK_DESCRIPTOR_TYPE_STORAGE_IMAGE, 2};
    VkDescriptorPoolCreateInfo pci{VK_STRUCTURE_TYPE_DESCRIPTOR_POOL_CREATE_INFO};
    pci.maxSets = 1;
    pci.poolSizeCount = 1;
    pci.pPoolSizes = &poolSize;
    vkCheck(vkCreateDescriptorPool(ctx_->device, &pci, nullptr, &overlayPool_), "vkCreateDescriptorPool (overlay)");
    VkDescriptorSetAllocateInfo ai{VK_STRUCTURE_TYPE_DESCRIPTOR_SET_ALLOCATE_INFO};
    ai.descriptorPool = overlayPool_;
    ai.descriptorSetCount = 1;
    ai.pSetLayouts = &overlayLayout_;
    vkCheck(vkAllocateDescriptorSets(ctx_->device, &ai, &overlaySet_), "vkAllocateDescriptorSets (overlay)");

    VkPushConstantRange pcr{VK_SHADER_STAGE_COMPUTE_BIT, 0, sizeof(OverlayPushConstants)};
    VkPipelineLayoutCreateInfo plci{VK_STRUCTURE_TYPE_PIPELINE_LAYOUT_CREATE_INFO};
    plci.setLayoutCount = 1;
    plci.pSetLayouts = &overlayLayout_;
    plci.pushConstantRangeCount = 1;
    plci.pPushConstantRanges = &pcr;
    vkCheck(vkCreatePipelineLayout(ctx_->device, &plci, nullptr, &overlayPipelineLayout_), "vkCreatePipelineLayout (overlay)");
    overlayPipeline_ = makeComputePipeline(ctx_->device, makeShader(ctx_->device, kOverlaySpv, sizeof(kOverlaySpv)),
                                           overlayPipelineLayout_);
}

void Renderer::writeDescriptors() {
    VkWriteDescriptorSetAccelerationStructureKHR asInfo{VK_STRUCTURE_TYPE_WRITE_DESCRIPTOR_SET_ACCELERATION_STRUCTURE_KHR};
    asInfo.accelerationStructureCount = 1;
    asInfo.pAccelerationStructures = &tlas_.handle;
    VkDescriptorImageInfo accumInfo{VK_NULL_HANDLE, accum_.view, VK_IMAGE_LAYOUT_GENERAL};
    VkDescriptorImageInfo outInfo{VK_NULL_HANDLE, output_.view, VK_IMAGE_LAYOUT_GENERAL};
    // 番号 → 中身（予備の方式では 0 番が BVH の箱、8 番が BVH の三角形の並び）
    VkDescriptorBufferInfo bufs[9] = {
        {bvhNodes_.buffer, 0, VK_WHOLE_SIZE}, {}, {},
        {vertices_.buffer, 0, VK_WHOLE_SIZE}, {indices_.buffer, 0, VK_WHOLE_SIZE},
        {triMaterials_.buffer, 0, VK_WHOLE_SIZE}, {materials_.buffer, 0, VK_WHOLE_SIZE},
        {exposure_.buffer, 0, VK_WHOLE_SIZE}, {bvhTris_.buffer, 0, VK_WHOLE_SIZE},
    };

    std::vector<BindingSpec> specs = pathTraceBindings(hardwareRt_);
    std::vector<VkWriteDescriptorSet> writes(specs.size());
    for (size_t i = 0; i < specs.size(); ++i) {
        VkWriteDescriptorSet& w = writes[i];
        w = VkWriteDescriptorSet{VK_STRUCTURE_TYPE_WRITE_DESCRIPTOR_SET};
        w.dstSet = set_;
        w.dstBinding = specs[i].binding;
        w.descriptorCount = 1;
        w.descriptorType = specs[i].type;
        if (specs[i].type == VK_DESCRIPTOR_TYPE_ACCELERATION_STRUCTURE_KHR) w.pNext = &asInfo;
        else if (specs[i].binding == 1) w.pImageInfo = &accumInfo;
        else if (specs[i].binding == 2) w.pImageInfo = &outInfo;
        else w.pBufferInfo = &bufs[specs[i].binding];
    }
    vkUpdateDescriptorSets(ctx_->device, static_cast<uint32_t>(writes.size()), writes.data(), 0, nullptr);
}

void Renderer::setOverlayTarget(const Image& target) {
    overlayTargetView_ = target.view;
    writeOverlayDescriptors();
}

void Renderer::writeOverlayDescriptors() {
    if (!overlay_.view || !overlayTargetView_) return;
    VkDescriptorImageInfo outInfo{VK_NULL_HANDLE, overlayTargetView_, VK_IMAGE_LAYOUT_GENERAL};
    VkDescriptorImageInfo ovInfo{VK_NULL_HANDLE, overlay_.view, VK_IMAGE_LAYOUT_GENERAL};
    VkWriteDescriptorSet writes[2]{};
    for (uint32_t i = 0; i < 2; ++i) {
        writes[i].sType = VK_STRUCTURE_TYPE_WRITE_DESCRIPTOR_SET;
        writes[i].dstSet = overlaySet_;
        writes[i].dstBinding = i;
        writes[i].descriptorCount = 1;
        writes[i].descriptorType = VK_DESCRIPTOR_TYPE_STORAGE_IMAGE;
    }
    writes[0].pImageInfo = &outInfo;
    writes[1].pImageInfo = &ovInfo;
    vkUpdateDescriptorSets(ctx_->device, 2, writes, 0, nullptr);
}

void Renderer::setOverlay(const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height) {
    vkDeviceWaitIdle(ctx_->device);
    ctx_->destroyImage(overlay_);
    if (width == 0 || height == 0) return;
    overlay_ = ctx_->createImage(width, height, VK_FORMAT_R8G8B8A8_UNORM,
                                 VK_IMAGE_USAGE_STORAGE_BIT | VK_IMAGE_USAGE_TRANSFER_DST_BIT);
    Buffer staging = ctx_->createBuffer(rgba.size(), VK_BUFFER_USAGE_TRANSFER_SRC_BIT,
                                        VK_MEMORY_PROPERTY_HOST_VISIBLE_BIT | VK_MEMORY_PROPERTY_HOST_COHERENT_BIT);
    memcpy(staging.mapped, rgba.data(), rgba.size());
    VkCommandBuffer cmd = ctx_->beginOneShot();
    transitionImage(cmd, overlay_.image, VK_IMAGE_LAYOUT_UNDEFINED, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL,
                    VK_PIPELINE_STAGE_TOP_OF_PIPE_BIT, 0, VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_WRITE_BIT);
    VkBufferImageCopy region{};
    region.imageSubresource = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 0, 1};
    region.imageExtent = {width, height, 1};
    vkCmdCopyBufferToImage(cmd, staging.buffer, overlay_.image, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, 1, &region);
    transitionImage(cmd, overlay_.image, VK_IMAGE_LAYOUT_TRANSFER_DST_OPTIMAL, VK_IMAGE_LAYOUT_GENERAL,
                    VK_PIPELINE_STAGE_TRANSFER_BIT, VK_ACCESS_TRANSFER_WRITE_BIT,
                    VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_ACCESS_SHADER_READ_BIT);
    ctx_->endOneShot(cmd);
    ctx_->destroyBuffer(staging);
    writeOverlayDescriptors();
}

void Renderer::recordOverlay(VkCommandBuffer cmd, int x, int y, float opacity) {
    if (!overlay_.view || !overlayTargetView_) return;
    // target は GENERAL レイアウトで、書き込みが終わっている前提（呼び出し側で切り替える）
    OverlayPushConstants pc{{x, y}, opacity};
    vkCmdBindPipeline(cmd, VK_PIPELINE_BIND_POINT_COMPUTE, overlayPipeline_);
    vkCmdBindDescriptorSets(cmd, VK_PIPELINE_BIND_POINT_COMPUTE, overlayPipelineLayout_, 0, 1, &overlaySet_, 0, nullptr);
    vkCmdPushConstants(cmd, overlayPipelineLayout_, VK_SHADER_STAGE_COMPUTE_BIT, 0, sizeof(pc), &pc);
    vkCmdDispatch(cmd, (overlay_.width + 7) / 8, (overlay_.height + 7) / 8, 1);
}

void Renderer::resize(uint32_t width, uint32_t height) {
    vkDeviceWaitIdle(ctx_->device);
    ctx_->destroyImage(accum_);
    ctx_->destroyImage(output_);
    accum_ = ctx_->createImage(width, height, VK_FORMAT_R32G32B32A32_SFLOAT, VK_IMAGE_USAGE_STORAGE_BIT);
    output_ = ctx_->createImage(width, height, VK_FORMAT_R8G8B8A8_UNORM,
                                VK_IMAGE_USAGE_STORAGE_BIT | VK_IMAGE_USAGE_TRANSFER_SRC_BIT);
    imagesFresh_ = true;
    accumulated_ = 0;
    writeDescriptors();
}

void Renderer::record(VkCommandBuffer cmd, const ViewParams& view) {
    if (imagesFresh_) {
        for (VkImage img : {accum_.image, output_.image})
            transitionImage(cmd, img, VK_IMAGE_LAYOUT_UNDEFINED, VK_IMAGE_LAYOUT_GENERAL,
                            VK_PIPELINE_STAGE_TOP_OF_PIPE_BIT, 0, VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT,
                            VK_ACCESS_SHADER_READ_BIT | VK_ACCESS_SHADER_WRITE_BIT);
        imagesFresh_ = false;
    }
    memset(exposure_.mapped, 0, sizeof(ExposureData));

    PushConstants pc{};
    float tanHalf = std::tan(view.fovY * 0.5f * 3.14159265f / 180.0f);
    float aspect = static_cast<float>(output_.width) / static_cast<float>(output_.height);
    pc.camPos[0] = view.eye.x; pc.camPos[1] = view.eye.y; pc.camPos[2] = view.eye.z; pc.camPos[3] = tanHalf;
    pc.camForward[0] = view.forward.x; pc.camForward[1] = view.forward.y; pc.camForward[2] = view.forward.z;
    pc.camForward[3] = view.panorama ? 1.0f : 0.0f;
    pc.camRight[0] = view.right.x; pc.camRight[1] = view.right.y; pc.camRight[2] = view.right.z; pc.camRight[3] = aspect;
    pc.camUp[0] = view.up.x; pc.camUp[1] = view.up.y; pc.camUp[2] = view.up.z;
    pc.sunDir[0] = sunDir_.x; pc.sunDir[1] = sunDir_.y; pc.sunDir[2] = sunDir_.z; pc.sunDir[3] = sunIntensity_;
    pc.misc[0] = view.exposure;
    pc.misc[1] = skyIntensity_;
    pc.misc[2] = static_cast<float>(glassStart_);
    pc.frame[0] = frameIndex_++;
    pc.frame[1] = accumulated_;
    pc.frame[2] = view.maxBounces;

    vkCmdBindPipeline(cmd, VK_PIPELINE_BIND_POINT_COMPUTE, pipeline_);
    vkCmdBindDescriptorSets(cmd, VK_PIPELINE_BIND_POINT_COMPUTE, pipelineLayout_, 0, 1, &set_, 0, nullptr);
    vkCmdPushConstants(cmd, pipelineLayout_, VK_SHADER_STAGE_COMPUTE_BIT, 0, sizeof(pc), &pc);
    vkCmdDispatch(cmd, (output_.width + 7) / 8, (output_.height + 7) / 8, 1);
    accumulated_++;

    // 自動露出の集計を CPU から読めるようにする
    VkMemoryBarrier toHost{VK_STRUCTURE_TYPE_MEMORY_BARRIER};
    toHost.srcAccessMask = VK_ACCESS_SHADER_WRITE_BIT;
    toHost.dstAccessMask = VK_ACCESS_HOST_READ_BIT;
    vkCmdPipelineBarrier(cmd, VK_PIPELINE_STAGE_COMPUTE_SHADER_BIT, VK_PIPELINE_STAGE_HOST_BIT, 0, 1, &toHost, 0,
                         nullptr, 0, nullptr);
}

float Renderer::finishFrame() {
    const auto* data = static_cast<const ExposureData*>(exposure_.mapped);
    if (data->logLumCount <= 0) return -1.0f;
    float avgLog2 = static_cast<float>(data->logLumSum) / 64.0f / static_cast<float>(data->logLumCount);
    return std::exp2(avgLog2);
}

void Renderer::destroy() {
    if (!ctx_) return;
    vkDeviceWaitIdle(ctx_->device);
    vkDestroyPipeline(ctx_->device, pipeline_, nullptr);
    vkDestroyPipelineLayout(ctx_->device, pipelineLayout_, nullptr);
    vkDestroyDescriptorPool(ctx_->device, pool_, nullptr);
    vkDestroyDescriptorSetLayout(ctx_->device, setLayout_, nullptr);
    ctx_->destroyImage(accum_);
    ctx_->destroyImage(output_);
    vkDestroyPipeline(ctx_->device, overlayPipeline_, nullptr);
    vkDestroyPipelineLayout(ctx_->device, overlayPipelineLayout_, nullptr);
    vkDestroyDescriptorPool(ctx_->device, overlayPool_, nullptr);
    vkDestroyDescriptorSetLayout(ctx_->device, overlayLayout_, nullptr);
    ctx_->destroyImage(overlay_);
    destroyAccel(tlas_);
    destroyAccel(blasOpaque_);
    destroyAccel(blasGlass_);
    for (Buffer* b : {&vertices_, &indices_, &triMaterials_, &materials_, &exposure_, &instances_, &bvhNodes_, &bvhTris_})
        ctx_->destroyBuffer(*b);
    ctx_ = nullptr;
}

}  // namespace t28

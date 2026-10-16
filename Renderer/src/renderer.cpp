#include "renderer.h"

#include <cmath>
#include <cstring>

#include "log.h"

namespace t28 {

// シェーダー（GPU 用プログラム）。shaders/pathtrace.comp を glslc で変換したもの
static const uint32_t kPathTraceSpv[] =
#include "pathtrace.spv.inc"
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

struct GpuMaterial {
    float color[4];
    float params[4];
};

VkDeviceSize alignUp(VkDeviceSize v, VkDeviceSize a) { return (v + a - 1) / a * a; }
}  // namespace

void Renderer::init(VkContext& ctx, Scene& scene) {
    ctx_ = &ctx;
    glassStart_ = scene.sortGlassLast();
    sunDir_ = scene.sunDir;
    sunIntensity_ = scene.sunIntensity;
    skyIntensity_ = scene.skyIntensity;
    logInfo("Scene: %zu vertices, %zu triangles (%u opaque, %zu glass), %zu materials",
            scene.vertices.size(), scene.triangleCount(), glassStart_, scene.triangleCount() - glassStart_,
            scene.materials.size());

    const VkBufferUsageFlags geomUsage = VK_BUFFER_USAGE_STORAGE_BUFFER_BIT |
                                         VK_BUFFER_USAGE_SHADER_DEVICE_ADDRESS_BIT |
                                         VK_BUFFER_USAGE_ACCELERATION_STRUCTURE_BUILD_INPUT_READ_ONLY_BIT_KHR;
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

    buildAccelerationStructures(scene);
    createPipeline();
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

void Renderer::buildAccelerationStructures(Scene& scene) {
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
    VkDescriptorSetLayoutBinding bindings[8]{};
    VkDescriptorType types[8] = {
        VK_DESCRIPTOR_TYPE_ACCELERATION_STRUCTURE_KHR, VK_DESCRIPTOR_TYPE_STORAGE_IMAGE, VK_DESCRIPTOR_TYPE_STORAGE_IMAGE,
        VK_DESCRIPTOR_TYPE_STORAGE_BUFFER, VK_DESCRIPTOR_TYPE_STORAGE_BUFFER, VK_DESCRIPTOR_TYPE_STORAGE_BUFFER,
        VK_DESCRIPTOR_TYPE_STORAGE_BUFFER, VK_DESCRIPTOR_TYPE_STORAGE_BUFFER,
    };
    for (uint32_t i = 0; i < 8; ++i) {
        bindings[i].binding = i;
        bindings[i].descriptorType = types[i];
        bindings[i].descriptorCount = 1;
        bindings[i].stageFlags = VK_SHADER_STAGE_COMPUTE_BIT;
    }
    VkDescriptorSetLayoutCreateInfo lci{VK_STRUCTURE_TYPE_DESCRIPTOR_SET_LAYOUT_CREATE_INFO};
    lci.bindingCount = 8;
    lci.pBindings = bindings;
    vkCheck(vkCreateDescriptorSetLayout(ctx_->device, &lci, nullptr, &setLayout_), "vkCreateDescriptorSetLayout");

    VkDescriptorPoolSize poolSizes[] = {
        {VK_DESCRIPTOR_TYPE_ACCELERATION_STRUCTURE_KHR, 1},
        {VK_DESCRIPTOR_TYPE_STORAGE_IMAGE, 2},
        {VK_DESCRIPTOR_TYPE_STORAGE_BUFFER, 5},
    };
    VkDescriptorPoolCreateInfo pci{VK_STRUCTURE_TYPE_DESCRIPTOR_POOL_CREATE_INFO};
    pci.maxSets = 1;
    pci.poolSizeCount = 3;
    pci.pPoolSizes = poolSizes;
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

    VkShaderModuleCreateInfo smci{VK_STRUCTURE_TYPE_SHADER_MODULE_CREATE_INFO};
    smci.codeSize = sizeof(kPathTraceSpv);
    smci.pCode = kPathTraceSpv;
    VkShaderModule module;
    vkCheck(vkCreateShaderModule(ctx_->device, &smci, nullptr, &module), "vkCreateShaderModule");

    VkComputePipelineCreateInfo cpci{VK_STRUCTURE_TYPE_COMPUTE_PIPELINE_CREATE_INFO};
    cpci.stage.sType = VK_STRUCTURE_TYPE_PIPELINE_SHADER_STAGE_CREATE_INFO;
    cpci.stage.stage = VK_SHADER_STAGE_COMPUTE_BIT;
    cpci.stage.module = module;
    cpci.stage.pName = "main";
    cpci.layout = pipelineLayout_;
    vkCheck(vkCreateComputePipelines(ctx_->device, VK_NULL_HANDLE, 1, &cpci, nullptr, &pipeline_), "vkCreateComputePipelines");
    vkDestroyShaderModule(ctx_->device, module, nullptr);
}

void Renderer::writeDescriptors() {
    VkWriteDescriptorSetAccelerationStructureKHR asInfo{VK_STRUCTURE_TYPE_WRITE_DESCRIPTOR_SET_ACCELERATION_STRUCTURE_KHR};
    asInfo.accelerationStructureCount = 1;
    asInfo.pAccelerationStructures = &tlas_.handle;
    VkDescriptorImageInfo accumInfo{VK_NULL_HANDLE, accum_.view, VK_IMAGE_LAYOUT_GENERAL};
    VkDescriptorImageInfo outInfo{VK_NULL_HANDLE, output_.view, VK_IMAGE_LAYOUT_GENERAL};
    VkDescriptorBufferInfo bufs[5] = {
        {vertices_.buffer, 0, VK_WHOLE_SIZE}, {indices_.buffer, 0, VK_WHOLE_SIZE},
        {triMaterials_.buffer, 0, VK_WHOLE_SIZE}, {materials_.buffer, 0, VK_WHOLE_SIZE},
        {exposure_.buffer, 0, VK_WHOLE_SIZE},
    };

    VkWriteDescriptorSet writes[8]{};
    for (uint32_t i = 0; i < 8; ++i) {
        writes[i].sType = VK_STRUCTURE_TYPE_WRITE_DESCRIPTOR_SET;
        writes[i].dstSet = set_;
        writes[i].dstBinding = i;
        writes[i].descriptorCount = 1;
    }
    writes[0].descriptorType = VK_DESCRIPTOR_TYPE_ACCELERATION_STRUCTURE_KHR;
    writes[0].pNext = &asInfo;
    writes[1].descriptorType = VK_DESCRIPTOR_TYPE_STORAGE_IMAGE;
    writes[1].pImageInfo = &accumInfo;
    writes[2].descriptorType = VK_DESCRIPTOR_TYPE_STORAGE_IMAGE;
    writes[2].pImageInfo = &outInfo;
    for (uint32_t i = 0; i < 5; ++i) {
        writes[3 + i].descriptorType = VK_DESCRIPTOR_TYPE_STORAGE_BUFFER;
        writes[3 + i].pBufferInfo = &bufs[i];
    }
    vkUpdateDescriptorSets(ctx_->device, 8, writes, 0, nullptr);
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
    float avgLog2 = static_cast<float>(data->logLumSum) / 256.0f / static_cast<float>(data->logLumCount);
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
    destroyAccel(tlas_);
    destroyAccel(blasOpaque_);
    destroyAccel(blasGlass_);
    for (Buffer* b : {&vertices_, &indices_, &triMaterials_, &materials_, &exposure_, &instances_}) ctx_->destroyBuffer(*b);
    ctx_ = nullptr;
}

}  // namespace t28

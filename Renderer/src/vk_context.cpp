#include "vk_context.h"

#include <cstring>

#include "log.h"
#include "i18n.h"

namespace t28 {

void vkCheck(VkResult result, const char* what) {
    if (result != VK_SUCCESS) fatal("%s (Vulkan error %d)", what, static_cast<int>(result));
}

// レイトレーシングに必要な拡張機能
static const char* kRayTracingExtensions[] = {
    VK_KHR_ACCELERATION_STRUCTURE_EXTENSION_NAME,
    VK_KHR_RAY_QUERY_EXTENSION_NAME,
    VK_KHR_DEFERRED_HOST_OPERATIONS_EXTENSION_NAME,
};

static bool hasExtension(const std::vector<VkExtensionProperties>& list, const char* name) {
    for (const auto& e : list)
        if (strcmp(e.extensionName, name) == 0) return true;
    return false;
}

static const char* deviceTypeName(VkPhysicalDeviceType t) {
    switch (t) {
        case VK_PHYSICAL_DEVICE_TYPE_DISCRETE_GPU: return "discrete";
        case VK_PHYSICAL_DEVICE_TYPE_INTEGRATED_GPU: return "integrated";
        case VK_PHYSICAL_DEVICE_TYPE_CPU: return "cpu";
        default: return "other";
    }
}

void VkContext::init(const std::vector<const char*>& instanceExtensions,
                     const std::function<VkSurfaceKHR(VkInstance)>& makeSurface, int preferredGpu,
                     bool allowHardwareRt) {
    if (volkInitialize() != VK_SUCCESS) fatal("%s", tr("NoVulkan"));

    VkApplicationInfo app{VK_STRUCTURE_TYPE_APPLICATION_INFO};
    app.pApplicationName = "Tools28 Render";
    app.apiVersion = VK_API_VERSION_1_2;
    VkInstanceCreateInfo ici{VK_STRUCTURE_TYPE_INSTANCE_CREATE_INFO};
    ici.pApplicationInfo = &app;
    ici.enabledExtensionCount = static_cast<uint32_t>(instanceExtensions.size());
    ici.ppEnabledExtensionNames = instanceExtensions.data();
    vkCheck(vkCreateInstance(&ici, nullptr, &instance), "vkCreateInstance");
    volkLoadInstance(instance);

    if (makeSurface) surface = makeSurface(instance);

    // --- GPU を選ぶ -------------------------------------------------------
    // ノートPCでは「CPU 内蔵の GPU」と「NVIDIA などの GPU」の2つがあることが多い。
    // RT コア（光の計算専用回路）がある GPU を最優先し、無ければ RT コアの無い GPU でも予備の方式で描く。
    // 同じ条件なら、性能の高い独立 GPU を内蔵 GPU より優先する。
    uint32_t count = 0;
    vkEnumeratePhysicalDevices(instance, &count, nullptr);
    std::vector<VkPhysicalDevice> devices(count);
    vkEnumeratePhysicalDevices(instance, &count, devices.data());
    if (count == 0) fatal("%s", tr("NoGpu"));

    int bestScore = -1;
    for (uint32_t i = 0; i < count; ++i) {
        VkPhysicalDeviceProperties props;
        vkGetPhysicalDeviceProperties(devices[i], &props);
        uint32_t extCount = 0;
        vkEnumerateDeviceExtensionProperties(devices[i], nullptr, &extCount, nullptr);
        std::vector<VkExtensionProperties> exts(extCount);
        vkEnumerateDeviceExtensionProperties(devices[i], nullptr, &extCount, exts.data());

        bool rt = allowHardwareRt && props.apiVersion >= VK_API_VERSION_1_2;
        for (const char* name : kRayTracingExtensions) rt = rt && hasExtension(exts, name);
        bool swap = !surface || hasExtension(exts, VK_KHR_SWAPCHAIN_EXTENSION_NAME);

        // 計算と描画の両方ができ、画面にも出せる窓口（キュー）を探す
        uint32_t qCount = 0;
        vkGetPhysicalDeviceQueueFamilyProperties(devices[i], &qCount, nullptr);
        std::vector<VkQueueFamilyProperties> families(qCount);
        vkGetPhysicalDeviceQueueFamilyProperties(devices[i], &qCount, families.data());
        int family = -1;
        for (uint32_t q = 0; q < qCount; ++q) {
            bool ok = (families[q].queueFlags & VK_QUEUE_GRAPHICS_BIT) && (families[q].queueFlags & VK_QUEUE_COMPUTE_BIT);
            if (ok && surface) {
                VkBool32 present = VK_FALSE;
                vkGetPhysicalDeviceSurfaceSupportKHR(devices[i], q, surface, &present);
                ok = present == VK_TRUE;
            }
            if (ok) { family = static_cast<int>(q); break; }
        }

        logInfo("GPU %u: %s (%s) raytracing=%s swapchain=%s queue=%d", i, props.deviceName,
                deviceTypeName(props.deviceType), rt ? "yes" : "no", swap ? "yes" : "no", family);
        if (!swap || family < 0) continue;

        int score = props.deviceType == VK_PHYSICAL_DEVICE_TYPE_DISCRETE_GPU ? 200
                  : props.deviceType == VK_PHYSICAL_DEVICE_TYPE_INTEGRATED_GPU ? 100 : 1;
        if (rt) score += 1000;
        if (static_cast<int>(i) == preferredGpu) score = 100000;
        if (score > bestScore) {
            bestScore = score;
            physical = devices[i];
            queueFamily = static_cast<uint32_t>(family);
            properties = props;
            hardwareRt = rt;
        }
    }
    if (!physical) fatal("%s", tr("NoGpu"));
    gpuName = properties.deviceName;
    logInfo("Using GPU: %s (%s)", gpuName.c_str(), hardwareRt ? "hardware ray tracing" : "software BVH");

    // --- 論理デバイスを作る -----------------------------------------------
    std::vector<const char*> deviceExts;
    if (hardwareRt) deviceExts.assign(std::begin(kRayTracingExtensions), std::end(kRayTracingExtensions));
    if (surface) deviceExts.push_back(VK_KHR_SWAPCHAIN_EXTENSION_NAME);

    VkPhysicalDeviceRayQueryFeaturesKHR rayQuery{VK_STRUCTURE_TYPE_PHYSICAL_DEVICE_RAY_QUERY_FEATURES_KHR};
    rayQuery.rayQuery = VK_TRUE;
    VkPhysicalDeviceAccelerationStructureFeaturesKHR accel{VK_STRUCTURE_TYPE_PHYSICAL_DEVICE_ACCELERATION_STRUCTURE_FEATURES_KHR};
    accel.accelerationStructure = VK_TRUE;
    accel.pNext = &rayQuery;
    VkPhysicalDeviceVulkan12Features v12{VK_STRUCTURE_TYPE_PHYSICAL_DEVICE_VULKAN_1_2_FEATURES};
    v12.bufferDeviceAddress = VK_TRUE;
    v12.pNext = &accel;

    float priority = 1.0f;
    VkDeviceQueueCreateInfo qci{VK_STRUCTURE_TYPE_DEVICE_QUEUE_CREATE_INFO};
    qci.queueFamilyIndex = queueFamily;
    qci.queueCount = 1;
    qci.pQueuePriorities = &priority;
    VkDeviceCreateInfo dci{VK_STRUCTURE_TYPE_DEVICE_CREATE_INFO};
    dci.pNext = hardwareRt ? &v12 : nullptr;  // 予備の方式では追加の機能は要らない
    dci.queueCreateInfoCount = 1;
    dci.pQueueCreateInfos = &qci;
    dci.enabledExtensionCount = static_cast<uint32_t>(deviceExts.size());
    dci.ppEnabledExtensionNames = deviceExts.data();
    vkCheck(vkCreateDevice(physical, &dci, nullptr, &device), "vkCreateDevice");
    volkLoadDevice(device);
    vkGetDeviceQueue(device, queueFamily, 0, &queue);

    VkCommandPoolCreateInfo pci{VK_STRUCTURE_TYPE_COMMAND_POOL_CREATE_INFO};
    pci.flags = VK_COMMAND_POOL_CREATE_RESET_COMMAND_BUFFER_BIT;
    pci.queueFamilyIndex = queueFamily;
    vkCheck(vkCreateCommandPool(device, &pci, nullptr, &commandPool), "vkCreateCommandPool");
}

void VkContext::destroy() {
    if (device) {
        vkDeviceWaitIdle(device);
        vkDestroyCommandPool(device, commandPool, nullptr);
        vkDestroyDevice(device, nullptr);
    }
    if (surface) vkDestroySurfaceKHR(instance, surface, nullptr);
    if (instance) vkDestroyInstance(instance, nullptr);
    device = VK_NULL_HANDLE;
    instance = VK_NULL_HANDLE;
}

uint32_t VkContext::findMemoryType(uint32_t typeBits, VkMemoryPropertyFlags props) const {
    VkPhysicalDeviceMemoryProperties mem;
    vkGetPhysicalDeviceMemoryProperties(physical, &mem);
    for (uint32_t i = 0; i < mem.memoryTypeCount; ++i)
        if ((typeBits & (1u << i)) && (mem.memoryTypes[i].propertyFlags & props) == props) return i;
    fatal("No suitable GPU memory type");
}

Buffer VkContext::createBuffer(VkDeviceSize size, VkBufferUsageFlags usage, VkMemoryPropertyFlags props) {
    Buffer b;
    b.size = size;
    VkBufferCreateInfo bci{VK_STRUCTURE_TYPE_BUFFER_CREATE_INFO};
    bci.size = size;
    bci.usage = usage;
    vkCheck(vkCreateBuffer(device, &bci, nullptr, &b.buffer), "vkCreateBuffer");

    VkMemoryRequirements req;
    vkGetBufferMemoryRequirements(device, b.buffer, &req);
    VkMemoryAllocateFlagsInfo flags{VK_STRUCTURE_TYPE_MEMORY_ALLOCATE_FLAGS_INFO};
    flags.flags = VK_MEMORY_ALLOCATE_DEVICE_ADDRESS_BIT;
    VkMemoryAllocateInfo mai{VK_STRUCTURE_TYPE_MEMORY_ALLOCATE_INFO};
    mai.pNext = (usage & VK_BUFFER_USAGE_SHADER_DEVICE_ADDRESS_BIT) ? &flags : nullptr;
    mai.allocationSize = req.size;
    mai.memoryTypeIndex = findMemoryType(req.memoryTypeBits, props);
    vkCheck(vkAllocateMemory(device, &mai, nullptr, &b.memory), "vkAllocateMemory (buffer)");
    vkCheck(vkBindBufferMemory(device, b.buffer, b.memory, 0), "vkBindBufferMemory");

    if (props & VK_MEMORY_PROPERTY_HOST_VISIBLE_BIT)
        vkCheck(vkMapMemory(device, b.memory, 0, size, 0, &b.mapped), "vkMapMemory");
    if (usage & VK_BUFFER_USAGE_SHADER_DEVICE_ADDRESS_BIT) {
        VkBufferDeviceAddressInfo info{VK_STRUCTURE_TYPE_BUFFER_DEVICE_ADDRESS_INFO};
        info.buffer = b.buffer;
        b.address = vkGetBufferDeviceAddress(device, &info);
    }
    return b;
}

void VkContext::destroyBuffer(Buffer& b) {
    if (b.buffer) vkDestroyBuffer(device, b.buffer, nullptr);
    if (b.memory) vkFreeMemory(device, b.memory, nullptr);
    b = Buffer{};
}

Buffer VkContext::createDeviceBuffer(const void* data, VkDeviceSize size, VkBufferUsageFlags usage) {
    if (size == 0) size = 16;  // 空のデータでも作れるように最小サイズを確保する
    Buffer staging = createBuffer(size, VK_BUFFER_USAGE_TRANSFER_SRC_BIT,
                                  VK_MEMORY_PROPERTY_HOST_VISIBLE_BIT | VK_MEMORY_PROPERTY_HOST_COHERENT_BIT);
    if (data) memcpy(staging.mapped, data, static_cast<size_t>(size));
    else memset(staging.mapped, 0, static_cast<size_t>(size));
    Buffer b = createBuffer(size, usage | VK_BUFFER_USAGE_TRANSFER_DST_BIT, VK_MEMORY_PROPERTY_DEVICE_LOCAL_BIT);
    VkCommandBuffer cmd = beginOneShot();
    VkBufferCopy region{0, 0, size};
    vkCmdCopyBuffer(cmd, staging.buffer, b.buffer, 1, &region);
    endOneShot(cmd);
    destroyBuffer(staging);
    return b;
}

Image VkContext::createImage(uint32_t w, uint32_t h, VkFormat format, VkImageUsageFlags usage) {
    Image img;
    img.width = w;
    img.height = h;
    img.format = format;
    VkImageCreateInfo ici{VK_STRUCTURE_TYPE_IMAGE_CREATE_INFO};
    ici.imageType = VK_IMAGE_TYPE_2D;
    ici.format = format;
    ici.extent = {w, h, 1};
    ici.mipLevels = 1;
    ici.arrayLayers = 1;
    ici.samples = VK_SAMPLE_COUNT_1_BIT;
    ici.tiling = VK_IMAGE_TILING_OPTIMAL;
    ici.usage = usage;
    vkCheck(vkCreateImage(device, &ici, nullptr, &img.image), "vkCreateImage");
    VkMemoryRequirements req;
    vkGetImageMemoryRequirements(device, img.image, &req);
    VkMemoryAllocateInfo mai{VK_STRUCTURE_TYPE_MEMORY_ALLOCATE_INFO};
    mai.allocationSize = req.size;
    mai.memoryTypeIndex = findMemoryType(req.memoryTypeBits, VK_MEMORY_PROPERTY_DEVICE_LOCAL_BIT);
    vkCheck(vkAllocateMemory(device, &mai, nullptr, &img.memory), "vkAllocateMemory (image)");
    vkCheck(vkBindImageMemory(device, img.image, img.memory, 0), "vkBindImageMemory");

    VkImageViewCreateInfo vci{VK_STRUCTURE_TYPE_IMAGE_VIEW_CREATE_INFO};
    vci.image = img.image;
    vci.viewType = VK_IMAGE_VIEW_TYPE_2D;
    vci.format = format;
    vci.subresourceRange = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 1, 0, 1};
    vkCheck(vkCreateImageView(device, &vci, nullptr, &img.view), "vkCreateImageView");
    return img;
}

void VkContext::destroyImage(Image& img) {
    if (img.view) vkDestroyImageView(device, img.view, nullptr);
    if (img.image) vkDestroyImage(device, img.image, nullptr);
    if (img.memory) vkFreeMemory(device, img.memory, nullptr);
    img = Image{};
}

VkCommandBuffer VkContext::beginOneShot() {
    VkCommandBufferAllocateInfo ai{VK_STRUCTURE_TYPE_COMMAND_BUFFER_ALLOCATE_INFO};
    ai.commandPool = commandPool;
    ai.level = VK_COMMAND_BUFFER_LEVEL_PRIMARY;
    ai.commandBufferCount = 1;
    VkCommandBuffer cmd;
    vkCheck(vkAllocateCommandBuffers(device, &ai, &cmd), "vkAllocateCommandBuffers");
    VkCommandBufferBeginInfo bi{VK_STRUCTURE_TYPE_COMMAND_BUFFER_BEGIN_INFO};
    bi.flags = VK_COMMAND_BUFFER_USAGE_ONE_TIME_SUBMIT_BIT;
    vkBeginCommandBuffer(cmd, &bi);
    return cmd;
}

void VkContext::endOneShot(VkCommandBuffer cmd) {
    vkEndCommandBuffer(cmd);
    VkSubmitInfo si{VK_STRUCTURE_TYPE_SUBMIT_INFO};
    si.commandBufferCount = 1;
    si.pCommandBuffers = &cmd;
    vkCheck(vkQueueSubmit(queue, 1, &si, VK_NULL_HANDLE), "vkQueueSubmit (one-shot)");
    vkQueueWaitIdle(queue);
    vkFreeCommandBuffers(device, commandPool, 1, &cmd);
}

void transitionImage(VkCommandBuffer cmd, VkImage image, VkImageLayout from, VkImageLayout to,
                     VkPipelineStageFlags srcStage, VkAccessFlags srcAccess,
                     VkPipelineStageFlags dstStage, VkAccessFlags dstAccess) {
    VkImageMemoryBarrier b{VK_STRUCTURE_TYPE_IMAGE_MEMORY_BARRIER};
    b.oldLayout = from;
    b.newLayout = to;
    b.srcAccessMask = srcAccess;
    b.dstAccessMask = dstAccess;
    b.srcQueueFamilyIndex = VK_QUEUE_FAMILY_IGNORED;
    b.dstQueueFamilyIndex = VK_QUEUE_FAMILY_IGNORED;
    b.image = image;
    b.subresourceRange = {VK_IMAGE_ASPECT_COLOR_BIT, 0, 1, 0, 1};
    vkCmdPipelineBarrier(cmd, srcStage, dstStage, 0, 0, nullptr, 0, nullptr, 1, &b);
}

}  // namespace t28

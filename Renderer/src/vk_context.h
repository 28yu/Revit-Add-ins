#pragma once
// Vulkan（GPU を直接動かす仕組み）の準備と、記憶領域・画像を作る道具。
#include <volk.h>

#include <functional>
#include <string>
#include <vector>

namespace t28 {

void vkCheck(VkResult result, const char* what);

struct Buffer {
    VkBuffer buffer = VK_NULL_HANDLE;
    VkDeviceMemory memory = VK_NULL_HANDLE;
    VkDeviceSize size = 0;
    void* mapped = nullptr;  // CPU から見える領域なら書き込み先
    VkDeviceAddress address = 0;
};

struct Image {
    VkImage image = VK_NULL_HANDLE;
    VkDeviceMemory memory = VK_NULL_HANDLE;
    VkImageView view = VK_NULL_HANDLE;
    VkFormat format = VK_FORMAT_UNDEFINED;
    uint32_t width = 0, height = 0;
};

class VkContext {
public:
    // instanceExtensions: 画面に出すときに必要な拡張機能（画像だけ作るなら空）
    // makeSurface: 画面表示用の面を作る関数（画像だけ作るなら空）。GPU 選びの前に呼ぶ
    // preferredGpu: 使う GPU の番号（-1 なら自動で選ぶ）
    void init(const std::vector<const char*>& instanceExtensions,
              const std::function<VkSurfaceKHR(VkInstance)>& makeSurface, int preferredGpu);
    void destroy();

    Buffer createBuffer(VkDeviceSize size, VkBufferUsageFlags usage, VkMemoryPropertyFlags props);
    void destroyBuffer(Buffer& b);
    // データを GPU 専用の記憶領域に送る（一時的な CPU 側の領域を経由する）
    Buffer createDeviceBuffer(const void* data, VkDeviceSize size, VkBufferUsageFlags usage);

    Image createImage(uint32_t w, uint32_t h, VkFormat format, VkImageUsageFlags usage);
    void destroyImage(Image& img);

    VkCommandBuffer beginOneShot();
    void endOneShot(VkCommandBuffer cmd);

    VkInstance instance = VK_NULL_HANDLE;
    VkPhysicalDevice physical = VK_NULL_HANDLE;
    VkDevice device = VK_NULL_HANDLE;
    VkQueue queue = VK_NULL_HANDLE;
    uint32_t queueFamily = 0;
    VkCommandPool commandPool = VK_NULL_HANDLE;
    VkPhysicalDeviceProperties properties{};
    std::string gpuName;
    VkSurfaceKHR surface = VK_NULL_HANDLE;

private:
    uint32_t findMemoryType(uint32_t typeBits, VkMemoryPropertyFlags props) const;
};

void transitionImage(VkCommandBuffer cmd, VkImage image, VkImageLayout from, VkImageLayout to,
                     VkPipelineStageFlags srcStage, VkAccessFlags srcAccess,
                     VkPipelineStageFlags dstStage, VkAccessFlags dstAccess);

}  // namespace t28

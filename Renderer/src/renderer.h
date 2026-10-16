#pragma once
// レイトレーシングで1フレームずつ描く本体。
#include "scene.h"
#include "vk_context.h"

namespace t28 {

struct ViewParams {
    Vec3 eye;
    Vec3 forward, right, up;  // 単位ベクトル
    float fovY = 60.0f;       // 縦の画角（度）
    bool panorama = false;    // 360°パノラマで描く
    float exposure = 1.0f;    // 明るさの倍率
    uint32_t maxBounces = 4;  // 光の反射を追う回数
};

class Renderer {
public:
    void init(VkContext& ctx, Scene& scene);
    void destroy();

    // 描画する大きさを変える（足し合わせもやり直し）
    void resize(uint32_t width, uint32_t height);
    // カメラが動いたときなどに、足し合わせをやり直す
    void resetAccumulation() { accumulated_ = 0; }
    uint32_t accumulatedSamples() const { return accumulated_; }

    // 1フレーム分の命令を記録する。終わると出力画像は GENERAL レイアウトのまま
    void record(VkCommandBuffer cmd, const ViewParams& view);
    // record した命令の実行が終わった後に呼ぶ（自動露出の集計を読む）
    // 戻り値: 画面全体の明るさ（対数平均）。集計がなければ負の値
    float finishFrame();

    const Image& output() const { return output_; }

private:
    void buildAccelerationStructures(Scene& scene);
    void createPipeline();
    void writeDescriptors();

    struct AccelStruct {
        VkAccelerationStructureKHR handle = VK_NULL_HANDLE;
        Buffer buffer;
        VkDeviceAddress address = 0;
    };
    AccelStruct buildBlas(VkDeviceSize indexOffsetBytes, uint32_t triangleCount, uint32_t vertexCount);
    void buildTlas(const std::vector<VkAccelerationStructureInstanceKHR>& instances);
    void destroyAccel(AccelStruct& a);

    VkContext* ctx_ = nullptr;
    Buffer vertices_, indices_, triMaterials_, materials_, exposure_, instances_;
    AccelStruct blasOpaque_, blasGlass_, tlas_;
    uint32_t glassStart_ = 0;
    float skyIntensity_ = 1.0f;
    Vec3 sunDir_;
    float sunIntensity_ = 5.0f;

    Image accum_, output_;
    bool imagesFresh_ = false;  // 作り直した直後（レイアウトの切り替えが必要）
    uint32_t accumulated_ = 0;
    uint32_t frameIndex_ = 0;

    VkDescriptorSetLayout setLayout_ = VK_NULL_HANDLE;
    VkDescriptorPool pool_ = VK_NULL_HANDLE;
    VkDescriptorSet set_ = VK_NULL_HANDLE;
    VkPipelineLayout pipelineLayout_ = VK_NULL_HANDLE;
    VkPipeline pipeline_ = VK_NULL_HANDLE;
};

}  // namespace t28

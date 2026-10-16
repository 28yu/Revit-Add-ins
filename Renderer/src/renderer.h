#pragma once
// レイトレーシングで1フレームずつ描く本体。
#include "bvh.h"
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
    // scene は sortGlassLast() 済み（glassStart = ガラスの最初の三角形）、bvh はその並びで作ったもの。
    // RT コアが無い GPU では bvh を GPU に送って、予備の方式で描く
    void init(VkContext& ctx, const Scene& scene, uint32_t glassStart, const Bvh& bvh);
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

    // 操作の案内（RGBA の画像）を設定し、recordOverlay で target に重ねる。
    // target は画面と同じ大きさの画像（小さく描いて拡大するときも、文字がぼやけないように拡大後に重ねる）
    void setOverlay(const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height);
    void setOverlayTarget(const Image& target);
    void recordOverlay(VkCommandBuffer cmd, int x, int y, float opacity);
    uint32_t overlayWidth() const { return overlay_.width; }
    uint32_t overlayHeight() const { return overlay_.height; }

private:
    void buildAccelerationStructures(const Scene& scene);
    void createPipeline();
    void createOverlayPipeline();
    void writeDescriptors();
    void writeOverlayDescriptors();

    struct AccelStruct {
        VkAccelerationStructureKHR handle = VK_NULL_HANDLE;
        Buffer buffer;
        VkDeviceAddress address = 0;
    };
    AccelStruct buildBlas(VkDeviceSize indexOffsetBytes, uint32_t triangleCount, uint32_t vertexCount);
    void buildTlas(const std::vector<VkAccelerationStructureInstanceKHR>& instances);
    void destroyAccel(AccelStruct& a);

    VkContext* ctx_ = nullptr;
    bool hardwareRt_ = false;
    Buffer vertices_, indices_, triMaterials_, materials_, exposure_, instances_;
    Buffer bvhNodes_, bvhTris_;  // 予備の方式（RT コア無し）で使う
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

    // 操作の案内を重ねる処理
    Image overlay_;
    VkImageView overlayTargetView_ = VK_NULL_HANDLE;
    VkDescriptorSetLayout overlayLayout_ = VK_NULL_HANDLE;
    VkDescriptorPool overlayPool_ = VK_NULL_HANDLE;
    VkDescriptorSet overlaySet_ = VK_NULL_HANDLE;
    VkPipelineLayout overlayPipelineLayout_ = VK_NULL_HANDLE;
    VkPipeline overlayPipeline_ = VK_NULL_HANDLE;
};

}  // namespace t28

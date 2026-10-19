#pragma once
// レイトレーシングで1フレームずつ描く本体。
//
// 1フレームの流れ:
//   1. パストレース（pathtrace_hw / pathtrace_sw）: 1 画素 1 本の光の通り道を計算（ざらついた明るさ＋面の情報）
//   2. ノイズ除去 その1（temporal）: 前のフレームまでの結果を今のカメラ位置に合わせて重ねる
//   3. ノイズ除去 その2（atrous ×5）: 面の境目を守りながら周りの画素と平均する
//   4. 仕上げ（composite）: 面の色を掛け戻し、トーンマッピングして画面に出す色にする（自動露出の集計も）
// カメラが止まっている間は 2 で全フレームを均等に足し合わせるので、時間とともに正確な結果に近づく。
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
    uint32_t samplesPerFrame = 1;  // 1 フレームで計算する回数（速い GPU ほど多くして、動いている間の画質を上げる）
};

class Renderer {
public:
    // scene は sortGlassLast() 済み（glassStart = ガラスの最初の三角形）、bvh はその並びで作ったもの。
    // RT コアが無い GPU では bvh を GPU に送って、予備の方式で描く
    // 空と雲の画像の横の画素数（init の前に設定する。縦はその 1/4）
    uint32_t skyWidth_ = 4096;
    void init(VkContext& ctx, const Scene& scene, uint32_t glassStart, const Bvh& bvh);
    void destroy();
    // モデルを入れ替える（Revit で変更があったとき）。描画の大きさ・カメラはそのまま
    void reloadScene(const Scene& scene, uint32_t glassStart, const Bvh& bvh);

    // 描画する大きさを変える（履歴も捨てる）
    void resize(uint32_t width, uint32_t height);
    // 見た目の切り替え。色は RGBA8（r が下位 8 ビット）で、a が 0 なら使わない。
    // background: 空・雲の代わりに出す背景色 / mono: 物を全部この色にする
    void setLook(uint32_t background, uint32_t mono);
    // 履歴を捨てて最初からやり直す
    void resetAccumulation() { resetHistory_ = true; }
    // カメラが止まってから描いたフレーム数
    uint32_t accumulatedSamples() const { return staticFrames_; }

    // 1フレーム分の命令を記録する（カメラが動いたかは前のフレームと比べて自動で判断する）。
    // 終わると出力画像は GENERAL レイアウトのまま
    void record(VkCommandBuffer cmd, const ViewParams& view);
    // record した命令の実行が終わった後に呼ぶ（自動露出の集計を読む）
    // 戻り値: 画面全体の明るさ（対数平均）。集計がなければ負の値
    float finishFrame();

    const Image& output() const { return output_; }

    // 画面に重ねる絵（操作の案内・ツールバー・進み具合）。slot ごとに 1 枚の RGBA の画像を持ち、recordOverlay で target に重ねる。
    // target は画面と同じ大きさの画像（小さく描いて拡大するときも、文字がぼやけないように拡大後に重ねる）
    static const uint32_t kOverlaySlots = 4;
    void setOverlay(uint32_t slot, const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height);
    void setOverlayTarget(const Image& target);
    void recordOverlay(VkCommandBuffer cmd, uint32_t slot, int x, int y, float opacity);

    // シェーダー1つ分（入出力の一覧・パイプライン）
    struct ComputePass {
        VkDescriptorSetLayout layout = VK_NULL_HANDLE;
        VkDescriptorPool pool = VK_NULL_HANDLE;
        std::vector<VkDescriptorSet> sets;
        VkPipelineLayout pipelineLayout = VK_NULL_HANDLE;
        VkPipeline pipeline = VK_NULL_HANDLE;
    };

private:
    void uploadScene(const Scene& scene, uint32_t glassStart, const Bvh& bvh);
    void computeSky();
    void destroyScene();
    void buildAccelerationStructures(const Scene& scene);
    void createPasses();
    void writeDescriptors();
    void writeOverlayDescriptors();
    void dispatch(VkCommandBuffer cmd, const ComputePass& pass, uint32_t set, const void* push, uint32_t pushSize,
                  uint32_t width, uint32_t height);

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
    Buffer texInfos_, texels_;  // テクスチャ（Revit の外観の「イメージ」「レリーフ パターン」）
    Buffer lights_;             // 照明器具の光源
    uint32_t lightCount_ = 0;
    float groundZ_ = 0.0f;      // 地面の高さ（モデルの一番下）
    Buffer bvhNodes_, bvhTris_;  // 予備の方式（RT コア無し）で使う
    AccelStruct blasOpaque_, blasGlass_, tlas_;
    uint32_t glassStart_ = 0;
    float skyIntensity_ = 1.0f;
    Vec3 sunDir_;
    float sunIntensity_ = 5.0f;

    // 描画とノイズ除去の画像。[2] は「今回」と「前回」を交互に使う
    Image noisy_, gpos_[2], gnormal_[2], galbedo_, history_[2], moments_[2], filtered_[2], albedoHist_[2], output_;
    bool imagesFresh_ = false;  // 作り直した直後（レイアウトの切り替えが必要）
    bool resetHistory_ = true;
    uint32_t lookBackground_ = 0, lookMono_ = 0;
    bool hasLastView_ = false;
    ViewParams lastView_;
    uint32_t staticFrames_ = 0;
    uint32_t frameIndex_ = 0;

    ComputePass pathTrace_, temporal_, atrous_, composite_, skyPass_, overlayPass_;
    Image skyImage_;  // 空と雲（起動時・太陽が変わったときに描く）
    Vec3 skySun_;

    // 操作の案内
    Image overlay_[kOverlaySlots];
    VkImageView overlayTargetView_ = VK_NULL_HANDLE;
};

}  // namespace t28

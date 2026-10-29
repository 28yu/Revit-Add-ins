#pragma once
// Intel Open Image Denoise（OIDN）で、止まっている間の画像のノイズを取る（AI のノイズ除去。CPU で動く）。
//
// DLL は実行時に読み込む（exe と同じ場所の oidn フォルダ。Linux では liboidn フォルダ）。
// 見つからない・読めないときは使わない（ビューアの exe には入れていないので、ビューアでは今まで通りのノイズ除去になる）。
#include <string>

namespace t28 {

class OidnDenoiser {
public:
    OidnDenoiser() = default;
    ~OidnDenoiser();
    OidnDenoiser(const OidnDenoiser&) = delete;
    OidnDenoiser& operator=(const OidnDenoiser&) = delete;

    // exe と同じ場所の DLL を読み込み、CPU の装置を用意する。使えれば true（理由はログに出す）
    bool load();
    bool available() const { return device_ != nullptr; }

    // color / albedo / normal: 1 画素 3 個の float（横 width × 縦 height）。output に結果を書く（color と同じでもよい）。
    // color は明るさ（HDR）、albedo・normal はノイズの無い面の色・向き。別のスレッドから呼んでよい（同時に 1 つだけ）
    bool denoise(const float* color, const float* albedo, const float* normal, float* output, int width, int height);

private:
    void* lib_ = nullptr;
    void* device_ = nullptr;
    struct Api;
    Api* api_ = nullptr;
};

}  // namespace t28

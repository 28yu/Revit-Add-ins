#pragma once
// 起動時の設定と、カメラ・自動露出の共通処理。
#include <algorithm>
#include <cmath>
#include <string>
#include <vector>

#include "bvh.h"
#include "renderer.h"
#include "scene.h"

namespace t28 {

struct Options {
    std::string scenePath;         // 空なら確認用の部屋
    std::string outputPath;        // 指定があれば画像だけ作って終わる
    std::string logPath;
    std::string title;             // モデル名・ビュー名（ウィンドウの題名に使う）
    std::string viewName;          // ビュー名（書き出すファイルの名前に使う）。無ければ題名の「 - 」より後ろ
    int width = 1280, height = 720;
    int samples = 256;             // 画像だけ作るときの足し合わせ回数
    bool panorama = false;
    int gpu = -1;
    uint32_t bounces = 4;
    float exposureEv = 0.0f;       // 自動露出への上乗せ（+1 で2倍明るい）
    bool autoExposure = true;
    float renderScale = 0.0f;      // ウィンドウより小さく描いて拡大する（速さ優先）。0 = 自動
    bool forceSoftware = false;    // RT コアがあっても予備の方式で描く（動作確認用）
    // 画像だけ作るときの見た目（動作確認用）: 0 = 使わない、それ以外は 0xFFBBGGRR
    uint32_t lookBackground = 0, lookMono = 0;
    bool hideGround = false, transparent = false;
    float timeHour = -1.0f;  // 時間帯（時。負なら Revit の太陽のまま）
    std::string lighting;    // 確認用: "sun" / "artificial" / "both"（レンダリング設定の照明を置き換える）
    float clouds = -1.0f;    // 確認用: 雲の量（負なら置き換えない）
    bool liveSync = false;         // Revit での変更を受け取る（シーンファイルの更新を見張る）。Revit から --live で指定
    int exportSamples = 0;
    int skySize = 4096;            // 空と雲の画像の横の画素数（動作確認で短くするときは小さくする）
    int motionReference = 0;       // 動作確認用: カメラを --motion-test と同じだけ動かした位置で、止めたまま描く（比べる正解）
    int motionTest = 0;            // 動作確認用: 画像だけ作るとき、カメラをこのフレーム数だけ動かし続けて最後の絵を保存する（残像の確認）         // 書き出しの足し合わせ回数（0 = 自動。動作確認で短くするとき用）
    int exitAfterFrames = 0;       // 動作確認用: このフレーム数を描いたら自動で閉じる（0 = 閉じない）
    // 性能測定（--bench）: 読み込み・準備の時間、GPU メモリ、止まっている間／動いている間の 1 フレームの時間、
    // シーンファイル・画像の大きさを測り、ログに「BENCH {...}」の 1 行（JSON）で出す。画像だけ作るモードで使う
    int gpuMemoryMb = 0;           // 動作確認用: GPU のメモリをこの量（MB）とみなす（少ない GPU での動きを確かめる）
    int bands = 0;                 // 光の計算を帯に分ける数（0 = 自動。--bands で決め打ち。動作確認用）
    bool instancing = true;        // 同じ形を 1 回だけ持つ（--no-instancing で切る。比較用）
    float localToneMapping = 1.0f;  // ローカルトーンマッピングの強さ（0 = 使わない。--local-tonemap で変える。比較用）
    bool skySampling = true;       // 空と地面の光を直接狙う（--sky-sampling 0 で切る。比較用）
    bool spatialDenoise = true;    // 周りの画素と平均する（--no-spatial-denoise で切る。ノイズの量を測るとき用）
    bool finalDenoise = true;      // 止まっている間・書き出しの仕上げに Open Image Denoise を使う（--final-denoise 0 で切る）
    int sparse = -1;               // 動いている間は 4 画素に 1 つだけ光を追う（-1 = 自動: GPU が速くないとき、0 = 使わない、1 = いつも）
    bool bench = false;
    int benchFrames = 16;          // 止まっている間・動いている間それぞれ何フレーム測るか
    std::string exportViewerPath;  // 指定があれば、ビューアの exe を書き出して終わる（動作確認用）
    std::string saveScenePath;     // 指定があれば、読み込んだシーンを .t28s で保存する（ファイルの大きさの確認用）
    double sceneLoadSeconds = 0;   // シーンの読み込み（または作成）にかかった時間（性能測定用。main が入れる）
};

// 歩き回るカメラ。向きは「方位（yaw）」と「見上げ角（pitch）」で持つ。Z が上
struct CameraController {
    Vec3 eye;
    float yaw = 0;    // ラジアン。0 = +Y（北）、正で東向き
    float pitch = 0;  // ラジアン。正で上向き
    float fovY = 60;

    void setFrom(const Camera& c) {
        eye = c.eye;
        float len = std::sqrt(c.forward.x * c.forward.x + c.forward.y * c.forward.y + c.forward.z * c.forward.z);
        if (len < 1e-6f) len = 1;
        yaw = std::atan2(c.forward.x / len, c.forward.y / len);
        pitch = std::asin(std::fmax(-1.0f, std::fmin(1.0f, c.forward.z / len)));
        fovY = c.fovY > 1 && c.fovY < 170 ? c.fovY : 60;
    }
    Vec3 forward() const {
        return {std::sin(yaw) * std::cos(pitch), std::cos(yaw) * std::cos(pitch), std::sin(pitch)};
    }
    Vec3 right() const { return {std::cos(yaw), -std::sin(yaw), 0}; }
    Vec3 up() const {
        Vec3 f = forward(), r = right();
        return {r.y * f.z - r.z * f.y, r.z * f.x - r.x * f.z, r.x * f.y - r.y * f.x};
    }
    // 前方向から向きを決め直す（回転の操作の後に使う）
    void setForward(Vec3 f) {
        float len = std::sqrt(f.x * f.x + f.y * f.y + f.z * f.z);
        if (len < 1e-6f) return;
        yaw = std::atan2(f.x / len, f.y / len);
        pitch = std::asin(std::fmax(-1.0f, std::fmin(1.0f, f.z / len)));
    }
    // 画面上の位置（-1〜1、上が +1）を通る光線の向き
    Vec3 rayThrough(float ndcX, float ndcY, float aspect) const {
        float t = std::tan(fovY * 0.5f * 3.14159265f / 180.0f);
        Vec3 f = forward(), r = right(), u = up();
        Vec3 d{f.x + (r.x * ndcX * aspect + u.x * ndcY) * t, f.y + (r.y * ndcX * aspect + u.y * ndcY) * t,
               f.z + (r.z * ndcX * aspect + u.z * ndcY) * t};
        float len = std::sqrt(d.x * d.x + d.y * d.y + d.z * d.z);
        return {d.x / len, d.y / len, d.z / len};
    }
    ViewParams view() const {
        ViewParams v;
        v.eye = eye;
        v.forward = forward();
        v.right = right();
        v.up = up();
        v.fovY = fovY;
        return v;
    }
};

// 自動露出: 画面の明るさの平均（対数平均）が kExposureKey になるように合わせる
const float kExposureKey = 0.32f;
struct AutoExposure {
    float current = 1.0f;
    bool initialized = false;

    // key: 画面の明るさ（対数平均）。smooth: 0 = 即座に合わせる、1 に近いほどゆっくり
    void update(float key, float evOffset, float smooth) {
        if (key <= 0) return;
        // 写真の標準（18% グレー）より明るめに合わせる（Enscape のような明るく軽い仕上がりにするため）
        float target = kExposureKey / key * std::exp2(evOffset);
        target = std::fmax(1.0f / 64.0f, std::fmin(1024.0f, target));
        if (!initialized) {
            current = target;
            initialized = true;
        } else {
            // 明るさの比で滑らかに近づける（急に明るさが変わると見づらいため）
            current = std::exp2(std::log2(current) * smooth + std::log2(target) * (1.0f - smooth));
        }
    }
};

// 光の計算を横長の帯に分ける数を決める。帯 1 本の計算を kMaxBandSeconds 以下に保つ
// （Windows は GPU の 1 回の処理が約 2 秒を超えるとリセットし、描画プログラムが落ちるため）
struct BandGovernor {
    static constexpr double kMaxBandSeconds = 0.3;
    uint32_t bands = 1;
    int fastFrames = 0;
    // 画像の大きさから、最初の分け方を決める（4K なら 4 本。測る前の 1 枚目で落ちないように）
    static uint32_t initialFor(uint32_t width, uint32_t height) {
        double pixels = static_cast<double>(width) * height;
        return static_cast<uint32_t>(std::fmin(16.0, std::fmax(1.0, std::ceil(pixels / 2.1e6))));
    }
    // 1 枚にかかった時間（全部の帯の合計）から分け方を直す。変えたら true
    bool update(double frameSeconds) {
        double perBand = frameSeconds / bands;
        if (perBand > kMaxBandSeconds && bands < 64) {
            bands = std::min(64u, bands * 2);
            fastFrames = 0;
            return true;
        }
        if (perBand < kMaxBandSeconds * 0.2 && bands > 1) {
            if (++fastFrames >= 10) {
                bands /= 2;
                fastFrames = 0;
                return true;
            }
        } else {
            fastFrames = 0;
        }
        return false;
    }
};

// 描画の準備: 1 回しか使わない形を 1 つにまとめ、形ごとにガラスの三角形を後ろに並べ替え、BVH を作る。
// instancing = false なら全部を 1 つの形にまとめる（同じ形を何度も持つ。比較・動作確認用）。戻り値は不透明な三角形の数
inline uint32_t prepareScene(Scene& scene, Bvh& bvh, bool instancing = true) {
    scene.ensureInstances();
    scene.mergeRarelyUsedMeshes(instancing ? 1u : 0xFFFFFFFFu);
    uint32_t opaque = scene.sortGlassLast();
    bvh.build(scene);
    return opaque;
}

// GPU の画像（RGBA8、GENERAL レイアウト）を CPU 側に写す。終わると GENERAL に戻す
// keepAlpha: 不透明度をそのまま残す（背景を透明にした PNG の書き出し）
std::vector<uint8_t> readImageRgba(VkContext& ctx, const Image& image, bool keepAlpha = false);

int runOffscreen(const Options& opt, Scene& scene);
int runWindow(const Options& opt, Scene& scene);

}  // namespace t28

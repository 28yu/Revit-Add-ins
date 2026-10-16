#pragma once
// 描画するモデルのデータ。座標はメートル、Z が上（Revit と同じ）。
#include <cstdint>
#include <string>
#include <vector>

namespace t28 {

struct Vec3 {
    float x = 0, y = 0, z = 0;
};

struct Vertex {
    float px, py, pz, u;  // 位置と、テクスチャ画像の横の位置（1 = 画像 1 枚分）
    float nx, ny, nz, v;  // 法線（面の向き）と、テクスチャ画像の縦の位置
};

struct Material {
    float r = 0.8f, g = 0.8f, b = 0.8f;  // 色（リニア）
    float opacity = 1.0f;                // 1 = 不透明、それ未満はガラスとして扱う
    float roughness = 0.5f;              // 0 = 鏡のようにつるつる、1 = ざらざら
    float metallic = 0.0f;
    float emissive = 0.0f;               // 自己発光の強さ
    float f0 = 0.04f;                    // 正面から見たときの反射率（Revit の外観の「反射率」）
    int32_t texture = -1;                // 色に使うテクスチャ画像の番号（-1 = 無し）
    std::string name;
};

// テクスチャ画像（Revit の外観の「イメージ」）。色は sRGB のまま 1 画素 4 バイト
struct Texture {
    uint32_t width = 0, height = 0;
    std::vector<uint8_t> rgba;
};

struct Camera {
    Vec3 eye{0, -5, 1.6f};
    Vec3 forward{0, 1, 0};
    Vec3 up{0, 0, 1};
    float fovY = 50.0f;  // 縦の画角（度）
};

struct Scene {
    std::vector<Vertex> vertices;
    std::vector<uint32_t> indices;       // 3つで1つの三角形
    std::vector<uint32_t> triMaterials;  // 三角形ごとのマテリアル番号
    std::vector<Material> materials;
    std::vector<Texture> textures;
    Camera camera;
    Vec3 sunDir{0.5f, -0.6f, 0.62f};     // 太陽のある方向
    float sunIntensity = 5.0f;
    float skyIntensity = 1.0f;

    // 不透明な三角形を前、ガラスの三角形を後ろに並べ替える。
    // 戻り値: ガラスの最初の三角形番号（＝不透明な三角形の数）
    uint32_t sortGlassLast();

    size_t triangleCount() const { return indices.size() / 3; }
};

// Revit 側（Tools28）が書き出すシーンファイル（.t28s）を読む
bool loadScene(const std::string& path, Scene& scene, std::string& error);
// 動作確認用の部屋（窓・柱・ガラス・鏡）
Scene makeTestScene();

}  // namespace t28

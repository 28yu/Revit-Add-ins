#include "scene.h"

#include <algorithm>
#include <cmath>
#include <cstdio>
#include <cstring>
#include <numeric>

#include "log.h"

namespace t28 {

// --- シーンファイル（.t28s）の形式 ------------------------------------------
// すべてリトルエンディアン。Revit 側は Commands/RealtimeRender/SceneWriter.cs が書く。
//   char[4]  "T28S"
//   uint32   版（1）
//   float[3] カメラ位置, float[3] 前方向, float[3] 上方向, float 縦の画角（度）
//   float[3] 太陽のある方向, float 太陽の明るさ, float 空の明るさ
//   uint32   マテリアル数 → 各: float r,g,b,不透明度,粗さ,金属,発光, uint32 名前のバイト数, 名前（UTF-8）
//   uint32   頂点数       → 各: float 位置x,y,z, 法線x,y,z
//   uint32   三角形数     → 各: uint32 頂点番号×3, マテリアル番号
static const uint32_t kVersion = 1;

namespace {
class Reader {
public:
    explicit Reader(FILE* f) : f_(f) {}
    template <typename T>
    bool read(T& value) { return fread(&value, sizeof(T), 1, f_) == 1; }
    bool readFloats(float* v, size_t n) { return fread(v, sizeof(float), n, f_) == n; }
    bool readBytes(void* p, size_t n) { return n == 0 || fread(p, 1, n, f_) == n; }

private:
    FILE* f_;
};
}  // namespace

bool loadScene(const std::string& path, Scene& scene, std::string& error) {
    FILE* f = openFile(path, "rb");
    if (!f) {
        error = "cannot open file: " + path;
        return false;
    }
    Reader r(f);
    auto fail = [&](const char* what) {
        error = what;
        fclose(f);
        return false;
    };

    char magic[4];
    uint32_t version = 0;
    if (!r.readBytes(magic, 4) || memcmp(magic, "T28S", 4) != 0) return fail("not a T28S file");
    if (!r.read(version) || version != kVersion) return fail("unsupported version");

    float cam[10];
    if (!r.readFloats(cam, 10)) return fail("truncated camera");
    scene.camera.eye = {cam[0], cam[1], cam[2]};
    scene.camera.forward = {cam[3], cam[4], cam[5]};
    scene.camera.up = {cam[6], cam[7], cam[8]};
    scene.camera.fovY = cam[9];

    float sun[5];
    if (!r.readFloats(sun, 5)) return fail("truncated sun");
    scene.sunDir = {sun[0], sun[1], sun[2]};
    scene.sunIntensity = sun[3];
    scene.skyIntensity = sun[4];

    uint32_t materialCount = 0;
    if (!r.read(materialCount) || materialCount > 1000000) return fail("bad material count");
    scene.materials.resize(materialCount);
    for (Material& m : scene.materials) {
        float v[7];
        uint32_t nameLen = 0;
        if (!r.readFloats(v, 7) || !r.read(nameLen) || nameLen > 4096) return fail("truncated material");
        m.r = v[0]; m.g = v[1]; m.b = v[2];
        m.opacity = v[3]; m.roughness = v[4]; m.metallic = v[5]; m.emissive = v[6];
        m.name.resize(nameLen);
        if (!r.readBytes(&m.name[0], nameLen)) return fail("truncated material name");
    }
    if (scene.materials.empty()) scene.materials.push_back(Material{});

    uint32_t vertexCount = 0;
    if (!r.read(vertexCount)) return fail("truncated vertex count");
    scene.vertices.resize(vertexCount);
    for (Vertex& v : scene.vertices) {
        float p[6];
        if (!r.readFloats(p, 6)) return fail("truncated vertices");
        v = {p[0], p[1], p[2], 1.0f, p[3], p[4], p[5], 0.0f};
    }

    uint32_t triCount = 0;
    if (!r.read(triCount)) return fail("truncated triangle count");
    scene.indices.resize(static_cast<size_t>(triCount) * 3);
    scene.triMaterials.resize(triCount);
    for (uint32_t t = 0; t < triCount; ++t) {
        uint32_t tri[4];
        if (!r.readBytes(tri, sizeof(tri))) return fail("truncated triangles");
        for (int k = 0; k < 3; ++k) {
            if (tri[k] >= vertexCount) return fail("vertex index out of range");
            scene.indices[t * 3 + k] = tri[k];
        }
        scene.triMaterials[t] = tri[3] < scene.materials.size() ? tri[3] : 0;
    }
    fclose(f);
    if (triCount == 0) {
        error = "the model has no triangles";
        return false;
    }
    return true;
}

uint32_t Scene::sortGlassLast() {
    const size_t n = triangleCount();
    std::vector<uint32_t> order(n);
    std::iota(order.begin(), order.end(), 0u);
    auto isGlass = [&](uint32_t t) { return materials[triMaterials[t]].opacity < 0.99f; };
    std::stable_partition(order.begin(), order.end(), [&](uint32_t t) { return !isGlass(t); });

    std::vector<uint32_t> newIndices(indices.size());
    std::vector<uint32_t> newMats(n);
    uint32_t glassStart = static_cast<uint32_t>(n);
    for (size_t i = 0; i < n; ++i) {
        uint32_t t = order[i];
        for (int k = 0; k < 3; ++k) newIndices[i * 3 + k] = indices[t * 3 + k];
        newMats[i] = triMaterials[t];
        if (isGlass(t) && glassStart == n) glassStart = static_cast<uint32_t>(i);
    }
    indices.swap(newIndices);
    triMaterials.swap(newMats);
    return glassStart;
}

// --- 動作確認用の部屋 --------------------------------------------------------
namespace {
struct Builder {
    Scene& s;
    uint32_t material(const char* name, float r, float g, float b, float rough, float metal = 0, float opacity = 1) {
        Material m;
        m.name = name;
        m.r = r; m.g = g; m.b = b;
        m.roughness = rough;
        m.metallic = metal;
        m.opacity = opacity;
        s.materials.push_back(m);
        return static_cast<uint32_t>(s.materials.size() - 1);
    }
    void quad(Vec3 a, Vec3 b, Vec3 c, Vec3 d, Vec3 n, uint32_t mat) {
        uint32_t base = static_cast<uint32_t>(s.vertices.size());
        for (const Vec3& p : {a, b, c, d}) s.vertices.push_back({p.x, p.y, p.z, 1, n.x, n.y, n.z, 0});
        uint32_t idx[6] = {0, 1, 2, 0, 2, 3};
        for (uint32_t i : idx) s.indices.push_back(base + i);
        s.triMaterials.push_back(mat);
        s.triMaterials.push_back(mat);
    }
    // 直方体（中心と大きさ）
    void box(Vec3 c, Vec3 size, uint32_t mat) {
        float x0 = c.x - size.x / 2, x1 = c.x + size.x / 2;
        float y0 = c.y - size.y / 2, y1 = c.y + size.y / 2;
        float z0 = c.z - size.z / 2, z1 = c.z + size.z / 2;
        quad({x0, y0, z1}, {x1, y0, z1}, {x1, y1, z1}, {x0, y1, z1}, {0, 0, 1}, mat);
        quad({x0, y1, z0}, {x1, y1, z0}, {x1, y0, z0}, {x0, y0, z0}, {0, 0, -1}, mat);
        quad({x0, y0, z0}, {x1, y0, z0}, {x1, y0, z1}, {x0, y0, z1}, {0, -1, 0}, mat);
        quad({x1, y1, z0}, {x0, y1, z0}, {x0, y1, z1}, {x1, y1, z1}, {0, 1, 0}, mat);
        quad({x0, y1, z0}, {x0, y0, z0}, {x0, y0, z1}, {x0, y1, z1}, {-1, 0, 0}, mat);
        quad({x1, y0, z0}, {x1, y1, z0}, {x1, y1, z1}, {x1, y0, z1}, {1, 0, 0}, mat);
    }
};
}  // namespace

Scene makeTestScene() {
    Scene s;
    Builder b{s};
    uint32_t wall = b.material("Wall", 0.8f, 0.8f, 0.78f, 0.8f);
    uint32_t floor = b.material("Floor", 0.42f, 0.27f, 0.15f, 0.35f);
    uint32_t ground = b.material("Ground", 0.25f, 0.33f, 0.15f, 0.9f);
    uint32_t steel = b.material("Steel", 0.75f, 0.75f, 0.77f, 0.3f, 1.0f);
    uint32_t glass = b.material("Glass", 0.85f, 0.95f, 0.95f, 0.02f, 0.0f, 0.2f);
    uint32_t red = b.material("Red", 0.7f, 0.06f, 0.05f, 0.6f);
    uint32_t green = b.material("Green", 0.08f, 0.5f, 0.12f, 0.5f);
    uint32_t mirror = b.material("Mirror", 0.95f, 0.95f, 0.95f, 0.02f, 1.0f);

    b.box({0, 0, -0.15f}, {60, 60, 0.3f}, ground);
    b.box({0, 0, 0.05f}, {6, 8, 0.1f}, floor);
    b.box({0, 0, 2.9f}, {6.4f, 8.4f, 0.2f}, wall);
    b.box({-3.1f, 0, 1.45f}, {0.2f, 8.4f, 2.9f}, wall);
    b.box({0, -4.1f, 1.45f}, {6.4f, 0.2f, 2.9f}, wall);
    // 北の壁（大きな窓あり）
    b.box({-2.6f, 4.1f, 1.45f}, {1.2f, 0.2f, 2.9f}, wall);
    b.box({2.6f, 4.1f, 1.45f}, {1.2f, 0.2f, 2.9f}, wall);
    b.box({0, 4.1f, 0.45f}, {4, 0.2f, 0.9f}, wall);
    b.box({0, 4.1f, 2.65f}, {4, 0.2f, 0.5f}, wall);
    b.box({0, 4.1f, 1.6f}, {4, 0.02f, 1.4f}, glass);
    // 東の壁（窓あり）
    b.box({3.1f, -2.5f, 1.45f}, {0.2f, 3.4f, 2.9f}, wall);
    b.box({3.1f, 3.2f, 1.45f}, {0.2f, 2.0f, 2.9f}, wall);
    b.box({3.1f, 0.7f, 0.4f}, {0.2f, 3.0f, 0.8f}, wall);
    b.box({3.1f, 0.7f, 2.6f}, {0.2f, 3.0f, 0.6f}, wall);
    b.box({3.1f, 0.7f, 1.5f}, {0.02f, 3.0f, 1.6f}, glass);

    b.box({-2.0f, -2.0f, 1.4f}, {0.3f, 0.3f, 2.8f}, steel);
    b.box({1.8f, 1.5f, 0.45f}, {0.8f, 0.8f, 0.9f}, green);
    b.box({0, 12, 1}, {2, 2, 2}, red);
    b.box({-2.98f, 0.5f, 1.4f}, {0.02f, 1.6f, 1.6f}, mirror);

    s.camera.eye = {0.5f, -3.2f, 1.5f};
    s.camera.forward = {0.0f, 1.0f, -0.05f};
    s.camera.up = {0, 0, 1};
    s.camera.fovY = 60.0f;
    // 北東・高さ 35° の太陽（窓から日が差し込む）
    float az = 30.0f * 3.14159265f / 180.0f, el = 35.0f * 3.14159265f / 180.0f;
    s.sunDir = {std::sin(az) * std::cos(el), std::cos(az) * std::cos(el), std::sin(el)};
    return s;
}

}  // namespace t28

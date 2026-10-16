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
//   uint32   版（2）
//   float[3] カメラ位置, float[3] 前方向, float[3] 上方向, float 縦の画角（度）
//   float[3] 太陽のある方向, float 太陽の明るさ, float 空の明るさ
//   uint32   マテリアル数 → 各: float r,g,b,不透明度,粗さ,金属,発光,反射率,
//                               色の画像: int32 テクスチャ番号（-1 = 無し）, float 貼り方×6,
//                               凹凸の画像: int32 テクスチャ番号, float 貼り方×6, int32 種類（0 = 高さ, 1 = 法線）, float 深さ,
//                               uint32 名前のバイト数, 名前（UTF-8）
//   uint32   頂点数       → 各: float 位置x,y,z, 法線x,y,z, 面の上の位置 u,v（メートル）
//   uint32   三角形数     → 各: uint32 頂点番号×3, マテリアル番号
//   uint32   テクスチャ数 → 各: uint32 幅, 高さ, 種類（bit0: 1 = 色ではないデータ）, 画素（RGBA 各 1 バイト）
//   uint32   光源数       → 各: float 位置x,y,z, 大きさ, 明るさr,g,b, 種類（0 = 全方向, 1 = 半球, 2 = スポット）,
//                               向きx,y,z, cos（落ち始め）, cos（届く範囲）
static const uint32_t kVersion = 3;

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

bool loadScene(const std::string& path, Scene& scene, std::string& error, long offset) {
    FILE* f = openFile(path, "rb");
    if (!f) {
        error = "cannot open file: " + path;
        return false;
    }
    if (offset > 0 && fseek(f, offset, SEEK_SET) != 0) {
        fclose(f);
        error = "cannot seek: " + path;
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
    if (!r.read(version) || version != kVersion) return fail("unsupported version (please update Tools28)");

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
        float v[8];
        int32_t bumpType = 0;
        uint32_t nameLen = 0;
        if (!r.readFloats(v, 8) || !r.read(m.color.texture) || !r.readFloats(m.color.xform, 6) || !r.read(m.bump.texture) ||
            !r.readFloats(m.bump.xform, 6) || !r.read(bumpType) || !r.read(m.bumpDepth) || !r.read(nameLen) || nameLen > 4096)
            return fail("truncated material");
        m.r = v[0]; m.g = v[1]; m.b = v[2];
        m.opacity = v[3]; m.roughness = v[4]; m.metallic = v[5]; m.emissive = v[6]; m.f0 = v[7];
        m.bumpType = bumpType == 1 ? BumpType::Normal : BumpType::Height;
        m.name.resize(nameLen);
        if (!r.readBytes(&m.name[0], nameLen)) return fail("truncated material name");
    }
    if (scene.materials.empty()) scene.materials.push_back(Material{});

    uint32_t vertexCount = 0;
    if (!r.read(vertexCount)) return fail("truncated vertex count");
    scene.vertices.resize(vertexCount);
    for (Vertex& v : scene.vertices) {
        float p[8];
        if (!r.readFloats(p, 8)) return fail("truncated vertices");
        v = {p[0], p[1], p[2], p[6], p[3], p[4], p[5], p[7]};
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

    uint32_t texCount = 0;
    if (!r.read(texCount) || texCount > 4096) return fail("bad texture count");
    scene.textures.resize(texCount);
    for (Texture& tex : scene.textures) {
        uint32_t flags = 0;
        if (!r.read(tex.width) || !r.read(tex.height) || !r.read(flags) || tex.width == 0 || tex.height == 0 ||
            tex.width > 8192 || tex.height > 8192)
            return fail("bad texture size");
        tex.linear = (flags & 1u) != 0;
        tex.rgba.resize(static_cast<size_t>(tex.width) * tex.height * 4);
        if (!r.readBytes(tex.rgba.data(), tex.rgba.size())) return fail("truncated texture");
    }
    for (Material& m : scene.materials) {
        if (m.color.texture >= static_cast<int32_t>(texCount)) m.color.texture = -1;
        if (m.bump.texture >= static_cast<int32_t>(texCount)) m.bump.texture = -1;
    }

    uint32_t lightCount = 0;
    if (!r.read(lightCount) || lightCount > 100000) return fail("bad light count");
    scene.lights.resize(lightCount);
    for (Light& l : scene.lights) {
        float v[13];
        if (!r.readFloats(v, 13)) return fail("truncated lights");
        l.position = {v[0], v[1], v[2]};
        l.radius = v[3];
        l.intensity = {v[4], v[5], v[6]};
        int type = static_cast<int>(v[7]);
        l.type = type == 1 ? LightType::Hemisphere : type == 2 ? LightType::Spot : LightType::Omni;
        l.direction = {v[8], v[9], v[10]};
        l.cosInner = v[11];
        l.cosOuter = v[12];
    }
    fclose(f);
    if (triCount == 0) {
        error = "the model has no triangles";
        return false;
    }
    return true;
}

namespace {
template <typename T>
void put(std::vector<uint8_t>& out, const T& v) {
    const uint8_t* p = reinterpret_cast<const uint8_t*>(&v);
    out.insert(out.end(), p, p + sizeof(T));
}
}  // namespace

void saveScene(const Scene& scene, std::vector<uint8_t>& out) {
    out.insert(out.end(), {'T', '2', '8', 'S'});
    put(out, kVersion);
    const Camera& c = scene.camera;
    for (float v : {c.eye.x, c.eye.y, c.eye.z, c.forward.x, c.forward.y, c.forward.z, c.up.x, c.up.y, c.up.z, c.fovY}) put(out, v);
    for (float v : {scene.sunDir.x, scene.sunDir.y, scene.sunDir.z, scene.sunIntensity, scene.skyIntensity}) put(out, v);
    put(out, static_cast<uint32_t>(scene.materials.size()));
    for (const Material& m : scene.materials) {
        for (float v : {m.r, m.g, m.b, m.opacity, m.roughness, m.metallic, m.emissive, m.f0}) put(out, v);
        put(out, m.color.texture);
        for (float v : m.color.xform) put(out, v);
        put(out, m.bump.texture);
        for (float v : m.bump.xform) put(out, v);
        put(out, static_cast<int32_t>(m.bumpType));
        put(out, m.bumpDepth);
        put(out, static_cast<uint32_t>(m.name.size()));
        out.insert(out.end(), m.name.begin(), m.name.end());
    }
    put(out, static_cast<uint32_t>(scene.vertices.size()));
    for (const Vertex& v : scene.vertices)
        for (float x : {v.px, v.py, v.pz, v.nx, v.ny, v.nz, v.u, v.v}) put(out, x);
    put(out, static_cast<uint32_t>(scene.triangleCount()));
    for (size_t t = 0; t < scene.triangleCount(); ++t) {
        for (int k = 0; k < 3; ++k) put(out, scene.indices[t * 3 + k]);
        put(out, scene.triMaterials[t]);
    }
    put(out, static_cast<uint32_t>(scene.textures.size()));
    for (const Texture& t : scene.textures) {
        put(out, t.width);
        put(out, t.height);
        put(out, static_cast<uint32_t>(t.linear ? 1u : 0u));
        out.insert(out.end(), t.rgba.begin(), t.rgba.end());
    }
    put(out, static_cast<uint32_t>(scene.lights.size()));
    for (const Light& l : scene.lights)
        for (float v : {l.position.x, l.position.y, l.position.z, l.radius, l.intensity.x, l.intensity.y, l.intensity.z,
                        static_cast<float>(l.type), l.direction.x, l.direction.y, l.direction.z, l.cosInner, l.cosOuter})
            put(out, v);
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
    // 四角形。テクスチャは a→b を横、a→d を縦にして 1 m で画像 1 枚分を貼る
    void quad(Vec3 a, Vec3 b, Vec3 c, Vec3 d, Vec3 n, uint32_t mat) {
        uint32_t base = static_cast<uint32_t>(s.vertices.size());
        auto sub = [](Vec3 p, Vec3 q) { return Vec3{p.x - q.x, p.y - q.y, p.z - q.z}; };
        auto len = [](Vec3 p) { return std::sqrt(p.x * p.x + p.y * p.y + p.z * p.z); };
        Vec3 eu = sub(b, a), ev = sub(d, a);
        float lu = std::max(len(eu), 1e-6f), lv = std::max(len(ev), 1e-6f);
        for (const Vec3& p : {a, b, c, d}) {
            Vec3 r = sub(p, a);
            float u = (r.x * eu.x + r.y * eu.y + r.z * eu.z) / lu, v = (r.x * ev.x + r.y * ev.y + r.z * ev.z) / lv;
            s.vertices.push_back({p.x, p.y, p.z, u, n.x, n.y, n.z, v});
        }
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
    // 床に木のフローリングの模様（テクスチャの確認用）: 幅 1/4 の板を 4 枚、板ごとに色を変えて木目の筋を入れる
    {
        Texture tex;
        tex.width = tex.height = 256;
        tex.rgba.resize(256 * 256 * 4);
        for (uint32_t y = 0; y < 256; ++y) {
            for (uint32_t x = 0; x < 256; ++x) {
                uint32_t plank = y / 64;
                float shade = 0.8f + 0.2f * std::sin(plank * 2.7f + 1.3f);
                float grain = 0.9f + 0.1f * std::sin(x * 0.11f + plank * 5.0f + std::sin(x * 0.03f + y * 0.2f) * 2.0f);
                float seam = (y % 64 < 2 || (x + plank * 97) % 256 < 2) ? 0.55f : 1.0f;
                float k = shade * grain * seam;
                uint8_t* p = &tex.rgba[(y * 256 + x) * 4];
                p[0] = static_cast<uint8_t>(std::min(255.0f, 176 * k));
                p[1] = static_cast<uint8_t>(std::min(255.0f, 124 * k));
                p[2] = static_cast<uint8_t>(std::min(255.0f, 78 * k));
                p[3] = 255;
            }
        }
        s.textures.push_back(std::move(tex));
        s.materials[floor].color.texture = 0;

        // 板の継ぎ目を溝にする凹凸（高さの画像。白 = 高い）
        Texture bump;
        bump.width = bump.height = 256;
        bump.linear = true;
        bump.rgba.resize(256 * 256 * 4);
        for (uint32_t y = 0; y < 256; ++y)
            for (uint32_t x = 0; x < 256; ++x) {
                uint32_t plank = y / 64;
                bool seam = y % 64 < 3 || (x + plank * 97) % 256 < 3;
                uint8_t* p = &bump.rgba[(y * 256 + x) * 4];
                p[0] = p[1] = p[2] = seam ? 0 : 255;
                p[3] = 255;
            }
        s.textures.push_back(std::move(bump));
        s.materials[floor].bump.texture = 1;
        s.materials[floor].bumpDepth = 0.003f;
    }
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

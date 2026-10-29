// 性能測定用のモデル（立体駐車場）。--bench-scene small|large|huge で使う。
//
// 目的: 大きなモデルでの速さ・メモリ・ファイルの大きさを、実物のモデルが無くても同じ条件でくり返し測る。
// 実際の Revit のモデルに近づけるため:
//   - 同じ形（柱・梁・車止め・車・照明）を位置だけ変えて何度も置く（Revit の書き出しは今、同じファミリでも形を毎回持つ）
//   - 屋根の下から外を見る構図（外が白くとびやすい、屋根の下は空の光が届きにくくノイズが多い）
//   - 色のテクスチャ（大きな画像）を使う
#include <cmath>
#include <cstdint>
#include <cstring>

#include "scene.h"

namespace t28 {

namespace {

struct Mesh {
    std::vector<Vertex> vertices;
    std::vector<uint32_t> indices;
    std::vector<int32_t> triMat;  // 三角形ごとのマテリアル（-1 または無し = 置くときのマテリアル）
};

void addQuad(Mesh& m, Vec3 a, Vec3 b, Vec3 c, Vec3 d, Vec3 n) {
    uint32_t base = static_cast<uint32_t>(m.vertices.size());
    auto sub = [](Vec3 p, Vec3 q) { return Vec3{p.x - q.x, p.y - q.y, p.z - q.z}; };
    auto len = [](Vec3 p) { return std::sqrt(p.x * p.x + p.y * p.y + p.z * p.z); };
    Vec3 eu = sub(b, a), ev = sub(d, a);
    float lu = std::fmax(len(eu), 1e-6f), lv = std::fmax(len(ev), 1e-6f);
    for (const Vec3& p : {a, b, c, d}) {
        Vec3 r = sub(p, a);
        float u = (r.x * eu.x + r.y * eu.y + r.z * eu.z) / lu, v = (r.x * ev.x + r.y * ev.y + r.z * ev.z) / lv;
        m.vertices.push_back({p.x, p.y, p.z, u, n.x, n.y, n.z, v});
    }
    for (uint32_t i : {0u, 1u, 2u, 0u, 2u, 3u}) m.indices.push_back(base + i);
}

// 直方体（最小の角と大きさ）
void addBox(Mesh& m, Vec3 o, Vec3 s) {
    float x0 = o.x, x1 = o.x + s.x, y0 = o.y, y1 = o.y + s.y, z0 = o.z, z1 = o.z + s.z;
    addQuad(m, {x0, y0, z1}, {x1, y0, z1}, {x1, y1, z1}, {x0, y1, z1}, {0, 0, 1});
    addQuad(m, {x0, y1, z0}, {x1, y1, z0}, {x1, y0, z0}, {x0, y0, z0}, {0, 0, -1});
    addQuad(m, {x0, y0, z0}, {x1, y0, z0}, {x1, y0, z1}, {x0, y0, z1}, {0, -1, 0});
    addQuad(m, {x1, y1, z0}, {x0, y1, z0}, {x0, y1, z1}, {x1, y1, z1}, {0, 1, 0});
    addQuad(m, {x0, y1, z0}, {x0, y0, z0}, {x0, y0, z1}, {x0, y1, z1}, {-1, 0, 0});
    addQuad(m, {x1, y0, z0}, {x1, y1, z0}, {x1, y1, z1}, {x1, y0, z1}, {1, 0, 0});
}

// 細かく分けた楕円体（車の形の代わり。三角形の数 = seg * seg * 2）。中心が原点、半径 r
Mesh ellipsoid(Vec3 r, int seg) {
    Mesh m;
    for (int j = 0; j <= seg; ++j) {
        float th = 3.14159265f * j / seg;
        for (int i = 0; i <= seg; ++i) {
            float ph = 2.0f * 3.14159265f * i / seg;
            Vec3 n{std::sin(th) * std::cos(ph), std::sin(th) * std::sin(ph), std::cos(th)};
            Vec3 p{n.x * r.x, n.y * r.y, n.z * r.z};
            Vec3 nn{n.x / r.x, n.y / r.y, n.z / r.z};
            float l = std::sqrt(nn.x * nn.x + nn.y * nn.y + nn.z * nn.z);
            m.vertices.push_back({p.x, p.y, p.z, ph, nn.x / l, nn.y / l, nn.z / l, th});
        }
    }
    for (int j = 0; j < seg; ++j)
        for (int i = 0; i < seg; ++i) {
            uint32_t a = j * (seg + 1) + i, b = a + 1, c = a + seg + 1, d = c + 1;
            for (uint32_t k : {a, c, b, b, c, d}) m.indices.push_back(k);
        }
    return m;
}

// 形（とマテリアル）をシーンに 1 回だけ入れ、形の番号を返す（同じ組み合わせは使い回す）
struct ShapeCache {
    std::vector<std::pair<std::pair<const Mesh*, uint32_t>, uint32_t>> known;
};
uint32_t shapeOf(Scene& s, ShapeCache& cache, const Mesh& m, uint32_t mat) {
    for (const auto& k : cache.known)
        if (k.first.first == &m && k.first.second == mat) return k.second;
    MeshRange r;
    r.firstTri = static_cast<uint32_t>(s.triMaterials.size());
    r.triCount = static_cast<uint32_t>(m.indices.size() / 3);
    uint32_t base = static_cast<uint32_t>(s.vertices.size());
    s.vertices.insert(s.vertices.end(), m.vertices.begin(), m.vertices.end());
    for (uint32_t i : m.indices) s.indices.push_back(base + i);
    for (uint32_t t = 0; t < r.triCount; ++t)
        s.triMaterials.push_back(t < m.triMat.size() && m.triMat[t] >= 0 ? static_cast<uint32_t>(m.triMat[t]) : mat);
    s.meshes.push_back(r);
    uint32_t id = static_cast<uint32_t>(s.meshes.size() - 1);
    cache.known.push_back({{&m, mat}, id});
    return id;
}

// 形を位置だけずらして（回転なし）置く。Revit のファミリと同じく、形は 1 回だけ持ち、置き場所を足す
ShapeCache* gCache = nullptr;
void place(Scene& s, const Mesh& m, Vec3 at, uint32_t mat) {
    Instance in;
    in.mesh = shapeOf(s, *gCache, m, mat);
    in.m[3] = at.x;
    in.m[7] = at.y;
    in.m[11] = at.z;
    s.instances.push_back(in);
}

uint32_t addMaterial(Scene& s, const char* name, float r, float g, float b, float rough, int32_t tex = -1,
                     float texSize = 1.0f) {
    Material m;
    m.name = name;
    m.r = r;
    m.g = g;
    m.b = b;
    m.roughness = rough;
    m.color.texture = tex;
    m.color.xform[0] = 1.0f / texSize;
    m.color.xform[4] = 1.0f / texSize;
    s.materials.push_back(m);
    return static_cast<uint32_t>(s.materials.size() - 1);
}

// コンクリートのような、ざらついた灰色の画像（決まった乱数なので毎回同じ）
int32_t addNoiseTexture(Scene& s, uint32_t size, uint8_t base, uint32_t seed) {
    Texture t;
    t.width = t.height = size;
    t.rgba.resize(static_cast<size_t>(size) * size * 4);
    uint32_t x = seed;
    for (size_t i = 0; i < static_cast<size_t>(size) * size; ++i) {
        x ^= x << 13;
        x ^= x >> 17;
        x ^= x << 5;
        int v = base + static_cast<int>(x % 41) - 20;
        uint8_t c = static_cast<uint8_t>(v < 0 ? 0 : v > 255 ? 255 : v);
        t.rgba[i * 4 + 0] = c;
        t.rgba[i * 4 + 1] = c;
        t.rgba[i * 4 + 2] = static_cast<uint8_t>(c > 4 ? c - 4 : 0);
        t.rgba[i * 4 + 3] = 255;
    }
    s.textures.push_back(std::move(t));
    return static_cast<int32_t>(s.textures.size() - 1);
}

}  // namespace

bool makeBenchScene(const std::string& preset, Scene& s) {
    // 柱の間隔（スパン）の数・階数・車の細かさ・画像の大きさ
    int nx, ny, floors, carSeg;
    uint32_t texSize;
    if (preset == "small") {
        nx = 6; ny = 4; floors = 1; carSeg = 16; texSize = 512;      // 約 4 万三角形
    } else if (preset == "large") {
        nx = 16; ny = 10; floors = 4; carSeg = 24; texSize = 2048;   // 約 170 万三角形
    } else if (preset == "huge") {
        nx = 20; ny = 12; floors = 6; carSeg = 32; texSize = 4096;   // 約 670 万三角形
    } else {
        return false;
    }
    const float span = 8.0f, storey = 3.4f, slab = 0.25f;
    ShapeCache cache;
    gCache = &cache;
    std::vector<Mesh> walls(static_cast<size_t>(floors));  // 階ごとの壁（形はそれぞれ別。1 回しか使わない形の例）
    const float width = nx * span, depth = ny * span;

    int32_t concreteTex = addNoiseTexture(s, texSize, 150, 12345u);
    int32_t asphaltTex = addNoiseTexture(s, texSize, 95, 67890u);
    uint32_t mConcrete = addMaterial(s, "コンクリート", 0.6f, 0.6f, 0.58f, 0.8f, concreteTex, 2.0f);
    uint32_t mFloor = addMaterial(s, "床", 0.35f, 0.35f, 0.36f, 0.7f, asphaltTex, 1.5f);
    uint32_t mSteel = addMaterial(s, "鉄骨", 0.55f, 0.52f, 0.58f, 0.5f);
    uint32_t mLine = addMaterial(s, "白線", 0.8f, 0.8f, 0.8f, 0.6f);
    uint32_t mCar[3] = {addMaterial(s, "車_赤", 0.5f, 0.05f, 0.04f, 0.3f), addMaterial(s, "車_白", 0.75f, 0.75f, 0.75f, 0.3f),
                        addMaterial(s, "車_黒", 0.04f, 0.04f, 0.05f, 0.3f)};
    uint32_t mWall = addMaterial(s, "壁", 0.7f, 0.68f, 0.64f, 0.9f);
    uint32_t mLamp = addMaterial(s, "照明", 0.9f, 0.9f, 0.9f, 0.5f);
    s.materials[mLamp].emissive = 2.0f;

    // 同じ形（ファミリに当たるもの）
    Mesh column, beamX, beamY, carStop, line, lamp, slabBay;
    addBox(column, {-0.4f, -0.4f, 0.0f}, {0.8f, 0.8f, storey - slab});
    // 梁: H 形（上下のフランジとウェブ）
    auto hBeam = [&](Mesh& m, bool alongX) {
        float L = span, h = 0.6f, b = 0.3f, tf = 0.03f, tw = 0.015f;
        auto box = [&](float o1, float o2, float z, float s1, float s2, float sz) {
            if (alongX) addBox(m, {o1, o2, z}, {s1, s2, sz});
            else addBox(m, {o2, o1, z}, {s2, s1, sz});
        };
        box(0, -b / 2, -h, L, b, tf);
        box(0, -b / 2, -tf, L, b, tf);
        box(0, -tw / 2, -h + tf, L, tw, h - 2 * tf);
    };
    hBeam(beamX, true);
    hBeam(beamY, false);
    addBox(carStop, {-0.8f, -0.08f, 0.0f}, {1.6f, 0.16f, 0.12f});
    addBox(line, {-0.06f, 0.0f, 0.0f}, {0.12f, 5.0f, 0.005f});
    addBox(lamp, {-0.6f, -0.08f, -0.06f}, {1.2f, 0.16f, 0.06f});
    addBox(slabBay, {0.0f, 0.0f, -slab}, {span, span, slab});
    Mesh car = ellipsoid({2.2f, 0.9f, 0.7f}, carSeg);
    // 車の窓（ガラス）: 同じ形の中に不透明な部分とガラスが混ざる例
    uint32_t mGlass = addMaterial(s, "ガラス", 0.8f, 0.9f, 0.95f, 0.05f);
    s.materials[mGlass].opacity = 0.3f;
    {
        size_t before = car.indices.size() / 3;
        addBox(car, {-0.9f, -0.75f, 0.35f}, {1.4f, 1.5f, 0.45f});
        car.triMat.assign(before, -1);
        car.triMat.resize(car.indices.size() / 3, static_cast<int32_t>(mGlass));
    }

    int carIndex = 0;
    for (int f = 0; f < floors; ++f) {
        float z = f * storey;
        for (int i = 0; i < nx; ++i)
            for (int j = 0; j < ny; ++j) {
                Vec3 o{i * span, j * span, z};
                place(s, slabBay, o, mFloor);
                place(s, slabBay, {o.x, o.y, z + storey}, mConcrete);  // 上の階の床（＝この階の天井）
                // 1 スパンに 3 台分の駐車枠（白線・車止め・車）
                for (int k = 0; k < 3; ++k) {
                    float cx = o.x + 1.4f + k * 2.6f, cy = o.y + 1.0f;
                    place(s, line, {cx - 1.3f, cy, z + 0.002f}, mLine);
                    place(s, carStop, {cx, cy + 0.6f, z}, mConcrete);
                    if ((carIndex++ % 4) != 3) place(s, car, {cx, cy + 3.0f, z + 0.75f}, mCar[carIndex % 3]);
                }
                place(s, lamp, {o.x + span / 2, o.y + span / 2, z + storey - 0.7f}, mLamp);
            }
        for (int i = 0; i <= nx; ++i)
            for (int j = 0; j <= ny; ++j) {
                place(s, column, {i * span, j * span, z}, mConcrete);
                if (i < nx) place(s, beamX, {i * span, j * span, z + storey}, mSteel);
                if (j < ny) place(s, beamY, {i * span, j * span, z + storey}, mSteel);
            }
        // 奥の壁だけ（手前と左右は外に開いている）
        addBox(walls[f], {0.0f, depth, z}, {width, 0.2f, storey});
        place(s, walls[f], {0, 0, 0}, mWall);
    }

    // 天井の照明（数を抑える: 4 スパンに 1 つ）
    for (int f = 0; f < floors; ++f)
        for (int i = 0; i < nx; i += 2)
            for (int j = 0; j < ny; j += 2) {
                Light l;
                l.position = {i * span + span / 2, j * span + span / 2, f * storey + storey - 0.8f};
                l.radius = 0.3f;
                l.intensity = {0.05f, 0.05f, 0.045f};
                l.type = LightType::Hemisphere;
                l.direction = {0, 0, -1};
                s.lights.push_back(l);
            }

    // 1 階の中から、開いている手前（-Y 側）の外を見る（外が白くとびやすい構図）
    s.camera.eye = {std::floor(nx / 2.0f) * span + span * 0.5f, depth * 0.55f, 1.5f};  // 柱と柱の間
    s.camera.forward = {0.15f, -1.0f, 0.03f};
    s.camera.up = {0, 0, 1};
    s.camera.fovY = 55.0f;
    s.sunDir = {0.4f, -0.7f, 0.6f};
    gCache = nullptr;
    return true;
}

}  // namespace t28

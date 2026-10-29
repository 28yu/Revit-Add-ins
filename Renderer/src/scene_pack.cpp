#include "scene_pack.h"

#include <algorithm>
#include <cmath>
#include <cstring>

#include "log.h"
#include "miniz.h"
#include "stb_image_write.h"

#define STB_IMAGE_IMPLEMENTATION
#define STBI_ONLY_JPEG
#define STBI_NO_STDIO
#include "stb_image.h"

namespace t28 {

namespace {

const uint32_t kPackVersion = 1;
const int kJpegQuality = 90;

template <typename T>
void put(std::vector<uint8_t>& out, const T& v) {
    const uint8_t* p = reinterpret_cast<const uint8_t*>(&v);
    out.insert(out.end(), p, p + sizeof(T));
}

uint32_t floatBits(float f) {
    uint32_t u;
    memcpy(&u, &f, 4);
    return u;
}
float bitsFloat(uint32_t u) {
    float f;
    memcpy(&f, &u, 4);
    return f;
}

// 面の向きを八面体に写して 16 ビット×2 に詰める（renderer.cpp の packNormal と同じ。0x8000 = 向き無し）
uint32_t packNormal(float x, float y, float z) {
    float len = std::fabs(x) + std::fabs(y) + std::fabs(z);
    if (len < 1e-12f) return 0x8000u;
    x /= len;
    y /= len;
    if (z < 0.0f) {
        float ox = (1.0f - std::fabs(y)) * (x >= 0.0f ? 1.0f : -1.0f);
        float oy = (1.0f - std::fabs(x)) * (y >= 0.0f ? 1.0f : -1.0f);
        x = ox;
        y = oy;
    }
    auto q = [](float f) {
        int v = static_cast<int>(std::lround(std::max(-1.0f, std::min(1.0f, f)) * 32767.0f));
        return static_cast<uint32_t>(static_cast<uint16_t>(static_cast<int16_t>(v)));
    };
    return q(x) | (q(y) << 16);
}
void unpackNormal(uint32_t p, float& x, float& y, float& z) {
    if ((p & 0xFFFFu) == 0x8000u) {
        x = y = z = 0.0f;
        return;
    }
    float ex = static_cast<int16_t>(p & 0xFFFFu) / 32767.0f, ey = static_cast<int16_t>(p >> 16) / 32767.0f;
    float nz = 1.0f - std::fabs(ex) - std::fabs(ey);
    if (nz < 0.0f) {
        float ox = (1.0f - std::fabs(ey)) * (ex >= 0.0f ? 1.0f : -1.0f);
        float oy = (1.0f - std::fabs(ex)) * (ey >= 0.0f ? 1.0f : -1.0f);
        ex = ox;
        ey = oy;
    }
    float len = std::sqrt(ex * ex + ey * ey + nz * nz);
    x = ex / len;
    y = ey / len;
    z = nz / len;
}

// 4 バイトの数の並びを「前の数との差」にして、バイトごとに並べ替えて書く（読むときは逆の順に戻す）
void putPlanes(std::vector<uint8_t>& out, std::vector<uint32_t> v, bool delta) {
    if (delta)
        for (size_t i = v.size(); i-- > 1;) v[i] -= v[i - 1];
    const size_t n = v.size(), base = out.size();
    out.resize(base + n * 4);
    for (int b = 0; b < 4; ++b)
        for (size_t i = 0; i < n; ++i) out[base + b * n + i] = static_cast<uint8_t>(v[i] >> (8 * b));
}

class Reader {
public:
    Reader(const uint8_t* p, size_t n) : p_(p), n_(n) {}
    template <typename T>
    bool read(T& v) { return bytes(&v, sizeof(T)); }
    bool bytes(void* dst, size_t n) {
        if (n > n_ - pos_) return false;
        memcpy(dst, p_ + pos_, n);
        pos_ += n;
        return true;
    }
    bool planes(std::vector<uint32_t>& v, size_t n, bool delta) {
        if (n > (n_ - pos_) / 4) return false;
        v.assign(n, 0);
        for (int b = 0; b < 4; ++b)
            for (size_t i = 0; i < n; ++i) v[i] |= static_cast<uint32_t>(p_[pos_ + b * n + i]) << (8 * b);
        pos_ += n * 4;
        if (delta)
            for (size_t i = 1; i < n; ++i) v[i] += v[i - 1];
        return true;
    }

private:
    const uint8_t* p_;
    size_t n_, pos_ = 0;
};

}  // namespace

bool isPackedScene(const uint8_t* data, size_t size) { return size >= 16 && memcmp(data, "T28P", 4) == 0; }

void packScene(const Scene& scene, const Camera& camera, std::vector<uint8_t>& out) {
    std::vector<uint8_t> raw;
    const Camera& c = camera;
    for (float v : {c.eye.x, c.eye.y, c.eye.z, c.forward.x, c.forward.y, c.forward.z, c.up.x, c.up.y, c.up.z, c.fovY}) put(raw, v);
    for (float v : {scene.sunDir.x, scene.sunDir.y, scene.sunDir.z, scene.sunIntensity, scene.skyIntensity, scene.latitude,
                    scene.trueNorth, scene.dayOfYear})
        put(raw, v);
    const RevitLook& rl = scene.revitLook;
    put(raw, static_cast<uint32_t>((rl.sun ? 1u : 0u) | (rl.artificial ? 2u : 0u)));
    put(raw, static_cast<uint32_t>(rl.background));
    put(raw, rl.backgroundColor);
    put(raw, rl.groundColor);
    put(raw, rl.cloudAmount);

    put(raw, static_cast<uint32_t>(scene.materials.size()));
    for (const Material& m : scene.materials) {
        for (float v : {m.r, m.g, m.b, m.opacity, m.roughness, m.metallic, m.emissive, m.f0}) put(raw, v);
        put(raw, m.color.texture);
        for (float v : m.color.xform) put(raw, v);
        put(raw, m.bump.texture);
        for (float v : m.bump.xform) put(raw, v);
        put(raw, static_cast<int32_t>(m.bumpType));
        put(raw, m.bumpDepth);
        put(raw, static_cast<uint32_t>(m.name.size()));
        raw.insert(raw.end(), m.name.begin(), m.name.end());
    }
    put(raw, static_cast<uint32_t>(scene.lights.size()));
    for (const Light& l : scene.lights)
        for (float v : {l.position.x, l.position.y, l.position.z, l.radius, l.intensity.x, l.intensity.y, l.intensity.z,
                        static_cast<float>(l.type), l.direction.x, l.direction.y, l.direction.z, l.cosInner, l.cosOuter})
            put(raw, v);
    put(raw, static_cast<uint32_t>(scene.meshes.size()));
    for (const MeshRange& m : scene.meshes) {
        put(raw, m.firstTri);
        put(raw, m.triCount);
    }
    put(raw, static_cast<uint32_t>(scene.meshes.empty() ? 0 : scene.instances.size()));
    if (!scene.meshes.empty())
        for (const Instance& in : scene.instances) {
            put(raw, in.mesh);
            for (float v : in.m) put(raw, v);
        }

    // 頂点（位置 x・y・z、面の向き、面の上の位置 u・v を、それぞれ別の並びにする）
    const size_t nv = scene.vertices.size();
    put(raw, static_cast<uint32_t>(nv));
    std::vector<uint32_t> a(nv);
    auto plane = [&](auto get, bool delta) {
        for (size_t i = 0; i < nv; ++i) a[i] = get(scene.vertices[i]);
        putPlanes(raw, a, delta);
    };
    plane([](const Vertex& v) { return floatBits(v.px); }, true);
    plane([](const Vertex& v) { return floatBits(v.py); }, true);
    plane([](const Vertex& v) { return floatBits(v.pz); }, true);
    plane([](const Vertex& v) { return packNormal(v.nx, v.ny, v.nz); }, true);
    plane([](const Vertex& v) { return floatBits(v.u); }, true);
    plane([](const Vertex& v) { return floatBits(v.v); }, true);

    // 三角形（頂点番号は前の番号との差。正負を交互に並べて小さな数にする）
    const size_t nt = scene.triangleCount();
    put(raw, static_cast<uint32_t>(nt));
    std::vector<uint32_t> idx(nt * 3);
    uint32_t prev = 0;
    for (size_t i = 0; i < nt * 3; ++i) {
        int32_t d = static_cast<int32_t>(scene.indices[i] - prev);
        idx[i] = (static_cast<uint32_t>(d) << 1) ^ static_cast<uint32_t>(d >> 31);
        prev = scene.indices[i];
    }
    putPlanes(raw, idx, false);
    putPlanes(raw, scene.triMaterials, false);

    // テクスチャ（色の画像で透明な部分が無ければ JPEG、それ以外はそのまま）
    put(raw, static_cast<uint32_t>(scene.textures.size()));
    size_t jpegCount = 0;
    for (const Texture& t : scene.textures) {
        bool opaque = true;
        for (size_t i = 3; i < t.rgba.size() && opaque; i += 4) opaque = t.rgba[i] == 255;
        std::vector<uint8_t> jpeg;
        if (!t.linear && opaque && t.width >= 16 && t.height >= 16) {
            std::vector<uint8_t> rgb(static_cast<size_t>(t.width) * t.height * 3);
            for (size_t i = 0, n = rgb.size() / 3; i < n; ++i) memcpy(&rgb[i * 3], &t.rgba[i * 4], 3);
            stbi_write_jpg_to_func(
                [](void* ctx, void* d, int n) {
                    auto* v = static_cast<std::vector<uint8_t>*>(ctx);
                    v->insert(v->end(), static_cast<uint8_t*>(d), static_cast<uint8_t*>(d) + n);
                },
                &jpeg, static_cast<int>(t.width), static_cast<int>(t.height), 3, rgb.data(), kJpegQuality);
        }
        put(raw, t.width);
        put(raw, t.height);
        put(raw, static_cast<uint32_t>((t.linear ? 1u : 0u) | (jpeg.empty() ? 0u : 2u)));  // bit1 = JPEG
        if (!jpeg.empty()) {
            put(raw, static_cast<uint32_t>(jpeg.size()));
            raw.insert(raw.end(), jpeg.begin(), jpeg.end());
            ++jpegCount;
        } else {
            raw.insert(raw.end(), t.rgba.begin(), t.rgba.end());
        }
    }

    size_t packedSize = 0;
    void* packed = tdefl_compress_mem_to_heap(raw.data(), raw.size(), &packedSize, TDEFL_DEFAULT_MAX_PROBES);
    out.insert(out.end(), {'T', '2', '8', 'P'});
    put(out, kPackVersion);
    put(out, static_cast<uint64_t>(raw.size()));
    if (packed) {
        out.insert(out.end(), static_cast<uint8_t*>(packed), static_cast<uint8_t*>(packed) + packedSize);
        mz_free(packed);
    }
    logInfo("Packed scene: %.1f MB -> %.1f MB (%zu vertices, %zu triangles, %zu of %zu textures as JPEG)", raw.size() / 1048576.0,
            packedSize / 1048576.0, nv, nt, jpegCount, scene.textures.size());
}

bool unpackScene(const uint8_t* data, size_t size, Scene& scene, std::string& error) {
    if (!isPackedScene(data, size)) {
        error = "not a packed scene";
        return false;
    }
    uint32_t version = 0;
    uint64_t rawSize = 0;
    memcpy(&version, data + 4, 4);
    memcpy(&rawSize, data + 8, 8);
    if (version != kPackVersion || rawSize > (16ull << 30)) {
        error = "unsupported packed scene";
        return false;
    }
    std::vector<uint8_t> raw(static_cast<size_t>(rawSize));
    size_t got = tinfl_decompress_mem_to_mem(raw.data(), raw.size(), data + 16, size - 16, 0);
    if (got == TINFL_DECOMPRESS_MEM_TO_MEM_FAILED || got != raw.size()) {
        error = "broken packed scene";
        return false;
    }
    Reader r(raw.data(), raw.size());
    auto fail = [&](const char* what) {
        error = what;
        return false;
    };
    float f[18];
    if (!r.bytes(f, sizeof(f))) return fail("truncated header");
    scene.camera.eye = {f[0], f[1], f[2]};
    scene.camera.forward = {f[3], f[4], f[5]};
    scene.camera.up = {f[6], f[7], f[8]};
    scene.camera.fovY = f[9];
    scene.sunDir = {f[10], f[11], f[12]};
    scene.sunIntensity = f[13];
    scene.skyIntensity = f[14];
    scene.latitude = f[15];
    scene.trueNorth = f[16];
    scene.dayOfYear = f[17];
    uint32_t lighting = 0, background = 0;
    if (!r.read(lighting) || !r.read(background) || !r.read(scene.revitLook.backgroundColor) || !r.read(scene.revitLook.groundColor) ||
        !r.read(scene.revitLook.cloudAmount))
        return fail("truncated settings");
    scene.revitLook.sun = (lighting & 1u) != 0;
    scene.revitLook.artificial = (lighting & 2u) != 0;
    scene.revitLook.background = static_cast<int>(std::min<uint32_t>(background, 2));

    uint32_t count = 0;
    if (!r.read(count) || count > 1000000) return fail("bad material count");
    scene.materials.resize(count);
    for (Material& m : scene.materials) {
        float v[8];
        int32_t bumpType = 0;
        uint32_t nameLen = 0;
        if (!r.bytes(v, sizeof(v)) || !r.read(m.color.texture) || !r.bytes(m.color.xform, sizeof(m.color.xform)) ||
            !r.read(m.bump.texture) || !r.bytes(m.bump.xform, sizeof(m.bump.xform)) || !r.read(bumpType) || !r.read(m.bumpDepth) ||
            !r.read(nameLen) || nameLen > 4096)
            return fail("truncated material");
        m.r = v[0]; m.g = v[1]; m.b = v[2];
        m.opacity = v[3]; m.roughness = v[4]; m.metallic = v[5]; m.emissive = v[6]; m.f0 = v[7];
        m.bumpType = bumpType == 1 ? BumpType::Normal : BumpType::Height;
        m.name.resize(nameLen);
        if (nameLen > 0 && !r.bytes(&m.name[0], nameLen)) return fail("truncated material name");
    }
    if (scene.materials.empty()) scene.materials.push_back(Material{});
    if (!r.read(count) || count > 100000) return fail("bad light count");
    scene.lights.resize(count);
    for (Light& l : scene.lights) {
        float v[13];
        if (!r.bytes(v, sizeof(v))) return fail("truncated lights");
        l.position = {v[0], v[1], v[2]};
        l.radius = v[3];
        l.intensity = {v[4], v[5], v[6]};
        int type = static_cast<int>(v[7]);
        l.type = type == 1 ? LightType::Hemisphere : type == 2 ? LightType::Spot : LightType::Omni;
        l.direction = {v[8], v[9], v[10]};
        l.cosInner = v[11];
        l.cosOuter = v[12];
    }
    uint32_t meshCount = 0, instanceCount = 0;
    if (!r.read(meshCount)) return fail("truncated meshes");
    scene.meshes.resize(meshCount);
    for (MeshRange& m : scene.meshes)
        if (!r.read(m.firstTri) || !r.read(m.triCount)) return fail("truncated meshes");
    if (!r.read(instanceCount) || instanceCount > 50000000u) return fail("bad instance count");
    scene.instances.resize(instanceCount);
    for (Instance& in : scene.instances) {
        if (!r.read(in.mesh) || !r.bytes(in.m, sizeof(in.m))) return fail("truncated instances");
        if (in.mesh >= meshCount) return fail("instance mesh out of range");
    }

    uint32_t nv = 0;
    if (!r.read(nv)) return fail("truncated vertices");
    std::vector<uint32_t> px, py, pz, nrm, u, v;
    if (!r.planes(px, nv, true) || !r.planes(py, nv, true) || !r.planes(pz, nv, true) || !r.planes(nrm, nv, true) ||
        !r.planes(u, nv, true) || !r.planes(v, nv, true))
        return fail("truncated vertices");
    scene.vertices.resize(nv);
    for (uint32_t i = 0; i < nv; ++i) {
        Vertex& vx = scene.vertices[i];
        vx.px = bitsFloat(px[i]);
        vx.py = bitsFloat(py[i]);
        vx.pz = bitsFloat(pz[i]);
        unpackNormal(nrm[i], vx.nx, vx.ny, vx.nz);
        vx.u = bitsFloat(u[i]);
        vx.v = bitsFloat(v[i]);
    }

    uint32_t nt = 0;
    if (!r.read(nt)) return fail("truncated triangles");
    std::vector<uint32_t> idx;
    if (!r.planes(idx, static_cast<size_t>(nt) * 3, false) || !r.planes(scene.triMaterials, nt, false))
        return fail("truncated triangles");
    scene.indices.resize(static_cast<size_t>(nt) * 3);
    uint32_t prev = 0;
    for (size_t i = 0; i < idx.size(); ++i) {
        int32_t d = static_cast<int32_t>((idx[i] >> 1) ^ (0u - (idx[i] & 1u)));
        prev += static_cast<uint32_t>(d);
        if (prev >= nv) return fail("vertex index out of range");
        scene.indices[i] = prev;
    }
    for (uint32_t& m : scene.triMaterials)
        if (m >= scene.materials.size()) m = 0;
    for (const MeshRange& m : scene.meshes)
        if (m.firstTri > nt || m.triCount > nt - m.firstTri) return fail("mesh out of range");

    if (!r.read(count) || count > 4096) return fail("bad texture count");
    scene.textures.resize(count);
    for (Texture& t : scene.textures) {
        uint32_t flags = 0;
        if (!r.read(t.width) || !r.read(t.height) || !r.read(flags) || t.width == 0 || t.height == 0 || t.width > 8192 ||
            t.height > 8192)
            return fail("bad texture size");
        t.linear = (flags & 1u) != 0;
        t.rgba.resize(static_cast<size_t>(t.width) * t.height * 4);
        if (flags & 2u) {
            uint32_t len = 0;
            std::vector<uint8_t> jpeg;
            if (!r.read(len)) return fail("truncated texture");
            jpeg.resize(len);
            if (!r.bytes(jpeg.data(), len)) return fail("truncated texture");
            int w = 0, h = 0, comp = 0;
            uint8_t* px4 = stbi_load_from_memory(jpeg.data(), static_cast<int>(len), &w, &h, &comp, 4);
            if (!px4 || static_cast<uint32_t>(w) != t.width || static_cast<uint32_t>(h) != t.height) {
                if (px4) stbi_image_free(px4);
                return fail("broken texture");
            }
            memcpy(t.rgba.data(), px4, t.rgba.size());
            stbi_image_free(px4);
        } else if (!r.bytes(t.rgba.data(), t.rgba.size())) {
            return fail("truncated texture");
        }
    }
    for (Material& m : scene.materials) {
        if (m.color.texture >= static_cast<int32_t>(count)) m.color.texture = -1;
        if (m.bump.texture >= static_cast<int32_t>(count)) m.bump.texture = -1;
    }
    if (nt == 0) return fail("the model has no triangles");
    return true;
}

}  // namespace t28

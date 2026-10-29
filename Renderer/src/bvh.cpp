#include "bvh.h"

#include <algorithm>
#include <cfloat>
#include <cmath>

namespace t28 {

namespace {

struct Box {
    float mn[3] = {FLT_MAX, FLT_MAX, FLT_MAX};
    float mx[3] = {-FLT_MAX, -FLT_MAX, -FLT_MAX};
    void grow(const float p[3]) {
        for (int k = 0; k < 3; ++k) { mn[k] = std::min(mn[k], p[k]); mx[k] = std::max(mx[k], p[k]); }
    }
    void grow(const Box& b) { grow(b.mn); grow(b.mx); }
    float area() const {
        float e[3] = {mx[0] - mn[0], mx[1] - mn[1], mx[2] - mn[2]};
        if (e[0] < 0) return 0;
        return 2.0f * (e[0] * e[1] + e[1] * e[2] + e[2] * e[0]);
    }
};

struct TriInfo {
    Box box;
    float centroid[3];
};

const uint32_t kLeafSize = 4;
const int kBins = 12;

// 箱の分け方は「SAH（表面積の比で、調べる手間が一番少なくなる分け方）」を 12 区間の近似で選ぶ。
// info の箱を入れ子に分けて nodes に足し、末端が指す番号（ids の値）を items に足す。戻り値は根の番号
uint32_t buildTree(std::vector<TriInfo>& info, const std::vector<uint32_t>& ids, std::vector<BvhNode>& nodes,
                   std::vector<uint32_t>& items) {
    const uint32_t n = static_cast<uint32_t>(info.size());
    const uint32_t root = static_cast<uint32_t>(nodes.size());
    const uint32_t itemBase = static_cast<uint32_t>(items.size());
    std::vector<uint32_t> order(n);
    for (uint32_t i = 0; i < n; ++i) order[i] = i;
    auto makeNode = [&](uint32_t first, uint32_t count) {
        Box b;
        for (uint32_t i = first; i < first + count; ++i) b.grow(info[order[i]].box);
        if (count == 0) b.mn[0] = b.mn[1] = b.mn[2] = b.mx[0] = b.mx[1] = b.mx[2] = 0.0f;
        BvhNode node{};
        for (int k = 0; k < 3; ++k) { node.bmin[k] = b.mn[k]; node.bmax[k] = b.mx[k]; }
        node.leftOrFirst = first;
        node.count = count;
        nodes.push_back(node);
        return static_cast<uint32_t>(nodes.size() - 1);
    };
    makeNode(0, n);
    // 空（三角形の無い形）は数 0 の末端のまま
    std::vector<uint32_t> stack = {root};
    while (!stack.empty()) {
        uint32_t ni = stack.back();
        stack.pop_back();
        uint32_t first = nodes[ni].leftOrFirst, count = nodes[ni].count;
        if (count <= kLeafSize) continue;
        Box centroidBox;
        for (uint32_t i = first; i < first + count; ++i) centroidBox.grow(info[order[i]].centroid);
        // 3 方向それぞれで、区間ごとの数と箱を数えて、一番よい分け目を探す
        float bestCost = FLT_MAX;
        int bestAxis = -1, bestSplit = 0;
        for (int axis = 0; axis < 3; ++axis) {
            float lo = centroidBox.mn[axis], hi = centroidBox.mx[axis];
            if (hi - lo < 1e-9f) continue;
            Box binBox[kBins];
            uint32_t binCount[kBins] = {};
            float scale = kBins / (hi - lo);
            for (uint32_t i = first; i < first + count; ++i) {
                const TriInfo& ti = info[order[i]];
                int b = std::min(kBins - 1, static_cast<int>((ti.centroid[axis] - lo) * scale));
                binCount[b]++;
                binBox[b].grow(ti.box);
            }
            float leftArea[kBins - 1], rightArea[kBins - 1];
            uint32_t leftCount[kBins - 1], rightCount[kBins - 1];
            Box lb, rb;
            uint32_t lc = 0, rc = 0;
            for (int i = 0; i < kBins - 1; ++i) {
                lc += binCount[i]; lb.grow(binBox[i]);
                leftCount[i] = lc; leftArea[i] = lb.area();
                rc += binCount[kBins - 1 - i]; rb.grow(binBox[kBins - 1 - i]);
                rightCount[kBins - 2 - i] = rc; rightArea[kBins - 2 - i] = rb.area();
            }
            for (int i = 0; i < kBins - 1; ++i) {
                if (leftCount[i] == 0 || rightCount[i] == 0) continue;
                float cost = leftArea[i] * leftCount[i] + rightArea[i] * rightCount[i];
                if (cost < bestCost) { bestCost = cost; bestAxis = axis; bestSplit = i; }
            }
        }
        uint32_t mid = first;
        if (bestAxis >= 0) {
            float lo = centroidBox.mn[bestAxis], hi = centroidBox.mx[bestAxis];
            float scale = kBins / (hi - lo);
            auto* begin = order.data() + first;
            auto* split = std::partition(begin, begin + count, [&](uint32_t t) {
                int b = std::min(kBins - 1, static_cast<int>((info[t].centroid[bestAxis] - lo) * scale));
                return b <= bestSplit;
            });
            mid = static_cast<uint32_t>(split - order.data());
        }
        if (mid == first || mid == first + count) {
            // 全部が同じ位置に重なっているなど、うまく分けられないときは半分に分ける
            mid = first + count / 2;
        }
        uint32_t left = makeNode(first, mid - first);
        makeNode(mid, first + count - mid);
        nodes[ni].leftOrFirst = left;
        nodes[ni].count = 0;
        stack.push_back(left);
        stack.push_back(left + 1);
    }
    // 末端の番号を、items の中の位置に直す
    for (size_t i = root; i < nodes.size(); ++i)
        if (nodes[i].count > 0) nodes[i].leftOrFirst += itemBase;
    for (uint32_t i = 0; i < n; ++i) items.push_back(ids[order[i]]);
    return root;
}

void transformPoint(const float m[12], const float p[3], float out[3]) {
    for (int r = 0; r < 3; ++r) out[r] = m[r * 4] * p[0] + m[r * 4 + 1] * p[1] + m[r * 4 + 2] * p[2] + m[r * 4 + 3];
}

}  // namespace

bool invertAffine(const float m[12], float out[12]) {
    float a = m[0], b = m[1], c = m[2], d = m[4], e = m[5], f = m[6], g = m[8], h = m[9], k = m[10];
    float det = a * (e * k - f * h) - b * (d * k - f * g) + c * (d * h - e * g);
    if (std::fabs(det) < 1e-20f) {
        const float id[12] = {1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0};
        std::copy(id, id + 12, out);
        return false;
    }
    float inv = 1.0f / det;
    float r[9] = {(e * k - f * h) * inv, (c * h - b * k) * inv, (b * f - c * e) * inv,
                  (f * g - d * k) * inv, (a * k - c * g) * inv, (c * d - a * f) * inv,
                  (d * h - e * g) * inv, (b * g - a * h) * inv, (a * e - b * d) * inv};
    const float t[3] = {m[3], m[7], m[11]};
    for (int row = 0; row < 3; ++row) {
        out[row * 4] = r[row * 3];
        out[row * 4 + 1] = r[row * 3 + 1];
        out[row * 4 + 2] = r[row * 3 + 2];
        out[row * 4 + 3] = -(r[row * 3] * t[0] + r[row * 3 + 1] * t[1] + r[row * 3 + 2] * t[2]);
    }
    return true;
}

void Bvh::build(const Scene& scene) {
    nodes_.clear();
    items_.clear();
    meshRoots_.clear();
    inverse_.clear();
    const size_t meshCount = scene.meshes.size(), instCount = scene.instances.size();
    if (meshCount == 0 || instCount == 0) return;
    // 形ごとの三角形の箱（形の座標）と、形全体の箱
    std::vector<std::vector<TriInfo>> meshInfo(meshCount);
    std::vector<Box> meshBox(meshCount);
    for (size_t mi = 0; mi < meshCount; ++mi) {
        const MeshRange& mr = scene.meshes[mi];
        meshInfo[mi].resize(mr.triCount);
        for (uint32_t i = 0; i < mr.triCount; ++i) {
            TriInfo& ti = meshInfo[mi][i];
            uint32_t t = mr.firstTri + i;
            for (int k = 0; k < 3; ++k) {
                const Vertex& v = scene.vertices[scene.indices[t * 3 + k]];
                float p[3] = {v.px, v.py, v.pz};
                ti.box.grow(p);
            }
            for (int a = 0; a < 3; ++a) ti.centroid[a] = (ti.box.mn[a] + ti.box.mx[a]) * 0.5f;
            meshBox[mi].grow(ti.box);
        }
    }
    // 上の段: 置き場所ごとの箱（形の箱の 8 つの角を世界の座標に動かして囲む）
    std::vector<TriInfo> instInfo(instCount);
    std::vector<uint32_t> instIds(instCount);
    inverse_.resize(instCount);
    for (size_t i = 0; i < instCount; ++i) {
        const Instance& in = scene.instances[i];
        instIds[i] = static_cast<uint32_t>(i);
        invertAffine(in.m, inverse_[i].data());
        const Box& mb = meshBox[in.mesh];
        if (mb.mn[0] > mb.mx[0]) {  // 三角形の無い形
            float zero[3] = {in.m[3], in.m[7], in.m[11]};
            instInfo[i].box.grow(zero);
        } else {
            for (int c = 0; c < 8; ++c) {
                float p[3] = {(c & 1) ? mb.mx[0] : mb.mn[0], (c & 2) ? mb.mx[1] : mb.mn[1], (c & 4) ? mb.mx[2] : mb.mn[2]};
                float w[3];
                transformPoint(in.m, p, w);
                instInfo[i].box.grow(w);
            }
        }
        for (int a = 0; a < 3; ++a) instInfo[i].centroid[a] = (instInfo[i].box.mn[a] + instInfo[i].box.mx[a]) * 0.5f;
    }
    nodes_.reserve(instCount * 2 + scene.triangleCount() * 2);
    buildTree(instInfo, instIds, nodes_, items_);  // 0 番が上の段の根
    meshRoots_.resize(meshCount);
    for (size_t mi = 0; mi < meshCount; ++mi) {
        const MeshRange& mr = scene.meshes[mi];
        std::vector<uint32_t> ids(mr.triCount);
        for (uint32_t i = 0; i < mr.triCount; ++i) ids[i] = mr.firstTri + i;
        meshRoots_[mi] = buildTree(meshInfo[mi], ids, nodes_, items_);
        std::vector<TriInfo>().swap(meshInfo[mi]);
    }
}

namespace {
// 箱に光線が入るか（入る距離 tmin が tmax 以下か）
bool boxHit(const BvhNode& nd, const float org[3], const float inv[3], float tmax) {
    float tmin = 0.0f;
    for (int k = 0; k < 3; ++k) {
        float t0 = (nd.bmin[k] - org[k]) * inv[k];
        float t1 = (nd.bmax[k] - org[k]) * inv[k];
        if (t0 > t1) std::swap(t0, t1);
        tmin = std::max(tmin, t0);
        tmax = std::min(tmax, t1);
    }
    return tmin <= tmax;
}
}  // namespace

float Bvh::intersect(const Scene& scene, Vec3 o, Vec3 d) const {
    if (nodes_.empty()) return -1.0f;
    float best = FLT_MAX;
    // 下の段: 形の座標に直した光線で三角形を調べる（距離 t は世界の座標と同じ値になる。向きを正規化しないため）
    auto meshHit = [&](uint32_t root, const float lo[3], const float ld[3]) {
        float inv[3] = {1.0f / ld[0], 1.0f / ld[1], 1.0f / ld[2]};
        uint32_t stack[64];
        int sp = 0;
        stack[sp++] = root;
        while (sp > 0) {
            const BvhNode& nd = nodes_[stack[--sp]];
            if (!boxHit(nd, lo, inv, best)) continue;
            if (nd.count > 0) {
                for (uint32_t i = nd.leftOrFirst; i < nd.leftOrFirst + nd.count; ++i) {
                    uint32_t t = items_[i];
                    const Vertex& a = scene.vertices[scene.indices[t * 3]];
                    const Vertex& b = scene.vertices[scene.indices[t * 3 + 1]];
                    const Vertex& c = scene.vertices[scene.indices[t * 3 + 2]];
                    // Möller–Trumbore（光線と三角形の交差）
                    float e1[3] = {b.px - a.px, b.py - a.py, b.pz - a.pz};
                    float e2[3] = {c.px - a.px, c.py - a.py, c.pz - a.pz};
                    float p[3] = {ld[1] * e2[2] - ld[2] * e2[1], ld[2] * e2[0] - ld[0] * e2[2], ld[0] * e2[1] - ld[1] * e2[0]};
                    float det = e1[0] * p[0] + e1[1] * p[1] + e1[2] * p[2];
                    if (std::fabs(det) < 1e-12f) continue;
                    float invDet = 1.0f / det;
                    float s[3] = {lo[0] - a.px, lo[1] - a.py, lo[2] - a.pz};
                    float u = (s[0] * p[0] + s[1] * p[1] + s[2] * p[2]) * invDet;
                    if (u < 0 || u > 1) continue;
                    float q[3] = {s[1] * e1[2] - s[2] * e1[1], s[2] * e1[0] - s[0] * e1[2], s[0] * e1[1] - s[1] * e1[0]};
                    float v = (ld[0] * q[0] + ld[1] * q[1] + ld[2] * q[2]) * invDet;
                    if (v < 0 || u + v > 1) continue;
                    float t2 = (e2[0] * q[0] + e2[1] * q[1] + e2[2] * q[2]) * invDet;
                    if (t2 > 1e-4f && t2 < best) best = t2;
                }
            } else if (sp + 2 <= 64) {
                stack[sp++] = nd.leftOrFirst;
                stack[sp++] = nd.leftOrFirst + 1;
            }
        }
    };

    // 上の段: 置き場所の箱をたどる
    float org[3] = {o.x, o.y, o.z}, dir[3] = {d.x, d.y, d.z};
    float inv[3] = {1.0f / d.x, 1.0f / d.y, 1.0f / d.z};
    uint32_t stack[64];
    int sp = 0;
    stack[sp++] = 0;
    while (sp > 0) {
        const BvhNode& nd = nodes_[stack[--sp]];
        if (!boxHit(nd, org, inv, best)) continue;
        if (nd.count > 0) {
            for (uint32_t i = nd.leftOrFirst; i < nd.leftOrFirst + nd.count; ++i) {
                uint32_t inst = items_[i];
                const float* m = inverse_[inst].data();
                float lo[3], ld[3];
                for (int r = 0; r < 3; ++r) {
                    lo[r] = m[r * 4] * org[0] + m[r * 4 + 1] * org[1] + m[r * 4 + 2] * org[2] + m[r * 4 + 3];
                    ld[r] = m[r * 4] * dir[0] + m[r * 4 + 1] * dir[1] + m[r * 4 + 2] * dir[2];
                }
                meshHit(meshRoots_[scene.instances[inst].mesh], lo, ld);
            }
        } else if (sp + 2 <= 64) {
            stack[sp++] = nd.leftOrFirst;
            stack[sp++] = nd.leftOrFirst + 1;
        }
    }
    return best == FLT_MAX ? -1.0f : best;
}

}  // namespace t28

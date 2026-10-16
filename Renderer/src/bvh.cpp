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

}  // namespace

// 箱の分け方は「SAH（表面積の比で、調べる手間が一番少なくなる分け方）」を 12 区間の近似で選ぶ
void Bvh::build(const Scene& scene) {
    const uint32_t n = static_cast<uint32_t>(scene.triangleCount());
    nodes_.clear();
    triIndices_.resize(n);
    std::vector<TriInfo> info(n);
    for (uint32_t t = 0; t < n; ++t) {
        triIndices_[t] = t;
        for (int k = 0; k < 3; ++k) {
            const Vertex& v = scene.vertices[scene.indices[t * 3 + k]];
            float p[3] = {v.px, v.py, v.pz};
            info[t].box.grow(p);
        }
        for (int a = 0; a < 3; ++a) info[t].centroid[a] = (info[t].box.mn[a] + info[t].box.mx[a]) * 0.5f;
    }
    if (n == 0) return;

    nodes_.reserve(static_cast<size_t>(n) * 2);
    auto makeNode = [&](uint32_t first, uint32_t count) {
        Box b;
        for (uint32_t i = first; i < first + count; ++i) b.grow(info[triIndices_[i]].box);
        BvhNode node{};
        for (int k = 0; k < 3; ++k) { node.bmin[k] = b.mn[k]; node.bmax[k] = b.mx[k]; }
        node.leftOrFirst = first;
        node.count = count;
        nodes_.push_back(node);
        return static_cast<uint32_t>(nodes_.size() - 1);
    };
    makeNode(0, n);

    std::vector<uint32_t> stack = {0};
    while (!stack.empty()) {
        uint32_t ni = stack.back();
        stack.pop_back();
        uint32_t first = nodes_[ni].leftOrFirst, count = nodes_[ni].count;
        if (count <= kLeafSize) continue;

        Box centroidBox;
        for (uint32_t i = first; i < first + count; ++i) centroidBox.grow(info[triIndices_[i]].centroid);

        // 3 方向それぞれで、区間ごとの三角形の数と箱を数えて、一番よい分け目を探す
        float bestCost = FLT_MAX;
        int bestAxis = -1, bestSplit = 0;
        for (int axis = 0; axis < 3; ++axis) {
            float lo = centroidBox.mn[axis], hi = centroidBox.mx[axis];
            if (hi - lo < 1e-9f) continue;
            Box binBox[kBins];
            uint32_t binCount[kBins] = {};
            float scale = kBins / (hi - lo);
            for (uint32_t i = first; i < first + count; ++i) {
                const TriInfo& ti = info[triIndices_[i]];
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
            auto* begin = triIndices_.data() + first;
            auto* split = std::partition(begin, begin + count, [&](uint32_t t) {
                int b = std::min(kBins - 1, static_cast<int>((info[t].centroid[bestAxis] - lo) * scale));
                return b <= bestSplit;
            });
            mid = static_cast<uint32_t>(split - triIndices_.data());
        }
        if (mid == first || mid == first + count) {
            // 全部が同じ位置に重なっているなど、うまく分けられないときは半分に分ける
            mid = first + count / 2;
        }

        uint32_t left = makeNode(first, mid - first);
        makeNode(mid, first + count - mid);
        nodes_[ni].leftOrFirst = left;
        nodes_[ni].count = 0;
        stack.push_back(left);
        stack.push_back(left + 1);
    }
}

float Bvh::intersect(const Scene& scene, Vec3 o, Vec3 d) const {
    if (nodes_.empty()) return -1.0f;
    float inv[3] = {1.0f / d.x, 1.0f / d.y, 1.0f / d.z};
    float org[3] = {o.x, o.y, o.z};
    float best = FLT_MAX;

    auto boxHit = [&](const BvhNode& nd) {
        float tmin = 0.0f, tmax = best;
        for (int k = 0; k < 3; ++k) {
            float t0 = (nd.bmin[k] - org[k]) * inv[k];
            float t1 = (nd.bmax[k] - org[k]) * inv[k];
            if (t0 > t1) std::swap(t0, t1);
            tmin = std::max(tmin, t0);
            tmax = std::min(tmax, t1);
        }
        return tmin <= tmax;
    };

    uint32_t stack[64];
    int sp = 0;
    stack[sp++] = 0;
    while (sp > 0) {
        const BvhNode& nd = nodes_[stack[--sp]];
        if (!boxHit(nd)) continue;
        if (nd.count > 0) {
            for (uint32_t i = nd.leftOrFirst; i < nd.leftOrFirst + nd.count; ++i) {
                uint32_t t = triIndices_[i];
                const Vertex& a = scene.vertices[scene.indices[t * 3]];
                const Vertex& b = scene.vertices[scene.indices[t * 3 + 1]];
                const Vertex& c = scene.vertices[scene.indices[t * 3 + 2]];
                // Möller–Trumbore（光線と三角形の交差）
                float e1[3] = {b.px - a.px, b.py - a.py, b.pz - a.pz};
                float e2[3] = {c.px - a.px, c.py - a.py, c.pz - a.pz};
                float p[3] = {d.y * e2[2] - d.z * e2[1], d.z * e2[0] - d.x * e2[2], d.x * e2[1] - d.y * e2[0]};
                float det = e1[0] * p[0] + e1[1] * p[1] + e1[2] * p[2];
                if (std::fabs(det) < 1e-12f) continue;
                float invDet = 1.0f / det;
                float s[3] = {o.x - a.px, o.y - a.py, o.z - a.pz};
                float u = (s[0] * p[0] + s[1] * p[1] + s[2] * p[2]) * invDet;
                if (u < 0 || u > 1) continue;
                float q[3] = {s[1] * e1[2] - s[2] * e1[1], s[2] * e1[0] - s[0] * e1[2], s[0] * e1[1] - s[1] * e1[0]};
                float v = (d.x * q[0] + d.y * q[1] + d.z * q[2]) * invDet;
                if (v < 0 || u + v > 1) continue;
                float t2 = (e2[0] * q[0] + e2[1] * q[1] + e2[2] * q[2]) * invDet;
                if (t2 > 1e-4f && t2 < best) best = t2;
            }
        } else if (sp + 2 <= 64) {
            stack[sp++] = nd.leftOrFirst;
            stack[sp++] = nd.leftOrFirst + 1;
        }
    }
    return best == FLT_MAX ? -1.0f : best;
}

}  // namespace t28

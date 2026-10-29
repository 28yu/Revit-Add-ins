#pragma once
// BVH（三角形を箱で入れ子に分けた索引）。光線がどの三角形に当たるかを速く調べるために使う。
//  - RT コアの無い GPU では、これを GPU に送ってシェーダーで当たり判定をする（予備の方式）
//  - マウスの指す位置に何があるか（ズーム・回転の中心）を CPU で調べるのにも使う
#include <array>
#include <cstdint>
#include <vector>

#include "scene.h"

namespace t28 {

// GPU 側（shaders/pathtrace_sw.comp）と同じ並び。1つ 32 バイト
struct BvhNode {
    float bmin[3];
    uint32_t leftOrFirst;  // 中間の箱: 左の子の番号（右の子は +1）/ 末端: triIndices の開始位置
    float bmax[3];
    uint32_t count;        // 0 = 中間の箱、1 以上 = 末端に入っている三角形の数
};

// 2 段の BVH: 上の段は「置き場所（インスタンス）」の箱、下の段は「形（メッシュ）」ごとの三角形の箱（形の座標）。
// nodes の 0 番が上の段の根（モデル全体の箱）。items は末端が指す番号の並び（上の段: 置き場所の番号、下の段: 三角形の番号）
class Bvh {
public:
    // scene は ensureInstances / sortGlassLast 済み
    void build(const Scene& scene);

    // 光線が最初に当たる距離を返す（当たらなければ負の値）
    float intersect(const Scene& scene, Vec3 origin, Vec3 dir) const;

    const std::vector<BvhNode>& nodes() const { return nodes_; }
    const std::vector<uint32_t>& triIndices() const { return items_; }
    // 形ごとの下の段の根の番号
    const std::vector<uint32_t>& meshRoots() const { return meshRoots_; }
    Vec3 boundsMin() const { return {nodes_.empty() ? 0 : nodes_[0].bmin[0], nodes_.empty() ? 0 : nodes_[0].bmin[1], nodes_.empty() ? 0 : nodes_[0].bmin[2]}; }
    Vec3 boundsMax() const { return {nodes_.empty() ? 0 : nodes_[0].bmax[0], nodes_.empty() ? 0 : nodes_[0].bmax[1], nodes_.empty() ? 0 : nodes_[0].bmax[2]}; }

private:
    std::vector<BvhNode> nodes_;
    std::vector<uint32_t> items_;
    std::vector<uint32_t> meshRoots_;
    std::vector<std::array<float, 12>> inverse_;  // 置き場所ごとの「世界 → 形の座標」
};

}  // namespace t28

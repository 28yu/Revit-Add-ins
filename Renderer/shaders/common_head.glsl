// 共通部分（前半）: 入出力・乱数・空・当たり判定の結果の形。
// pathtrace_hw.comp / pathtrace_sw.comp から取り込まれる。
// 出力（ノイズ除去 denoise の材料）。色は「最初に当たった面の色」で割った明るさにしておき、
// ノイズ除去のあとで色を掛け戻す（模様や色の境目をぼかさないため）
layout(set = 0, binding = 1, rgba32f) uniform writeonly image2D uNoisy;   // rgb = 今回の明るさ（色で割ったもの）
layout(set = 0, binding = 2, rgba32f) uniform writeonly image2D uGPos;    // xyz = 最初に当たった位置, w = 距離（空なら -1）
layout(set = 0, binding = 7, rgba32f) uniform writeonly image2D uGNormal; // xyz = 面の向き
layout(set = 0, binding = 9, rgba8) uniform writeonly image2D uGAlbedo;   // rgb = 最初に当たった面の色

struct Vertex { vec4 position; vec4 normal; };  // position.w = 面の上の位置 u, normal.w = v（メートル）
layout(set = 0, binding = 3, std430) readonly buffer Vertices { Vertex vertices[]; };
layout(set = 0, binding = 4, std430) readonly buffer Indices { uint indices[]; };
// 三角形ごとのマテリアル番号。ガラスの三角形はガラス用の並びに入っている
layout(set = 0, binding = 5, std430) readonly buffer TriMaterials { uint triMaterials[]; };

struct Material {
    vec4 color;      // rgb = 色（リニア）、a = 不透明度
    vec4 params;     // x = 粗さ, y = 金属らしさ, z = 自己発光の強さ, w = 正面から見たときの反射率
    ivec4 tex;       // x = 色の画像の番号, y = 凹凸の画像の番号（-1 = 無し）, z = 凹凸の種類（0 = 高さ, 1 = 法線）
    vec4 colorMap0;  // 色の画像の貼り方: 画像の横 = dot(xyz, (u, v, 1))、w = 凹凸の深さ
    vec4 colorMap1;  //                   画像の縦 = dot(xyz, (u, v, 1))
    vec4 bumpMap0;   // 凹凸の画像の貼り方
    vec4 bumpMap1;
};
layout(set = 0, binding = 6, std430) readonly buffer Materials { Material materials[]; };

// テクスチャ（Revit の外観の「イメージ」）。画素は 1 つの uint に RGBA（sRGB）を詰めてある。
// 段（ミップマップ）ごとに縦横半分にした画像が、一番大きい画像の後ろに続けて入っている
struct TexInfo { uint offset; uint width; uint height; uint levels; };
layout(set = 0, binding = 10, std430) readonly buffer TexInfos { TexInfo texInfos[]; };
layout(set = 0, binding = 11, std430) readonly buffer Texels { uint texels[]; };

// 照明器具の光源。明るさは「光度（カンデラ）÷ 20000」（太陽 = 5 が 10 万ルクスの単位）
struct Light {
    vec4 positionRadius;  // xyz = 位置, w = 光る部分の半径
    vec4 intensityType;   // rgb = 明るさ, w = 種類（0 = 全方向, 1 = 半球, 2 = スポット）
    vec4 direction;       // xyz = 向き（半球・スポット）, w = cos（明るさが落ち始める角度）
    vec4 cone;            // x = cos（光が届く範囲）
};
layout(set = 0, binding = 12, std430) readonly buffer Lights { Light lights[]; };

layout(push_constant) uniform Params {
    vec4 camPos;      // xyz = 位置, w = 縦の画角の tan(半分)
    vec4 camForward;  // xyz = 前, w = パノラマなら 1
    vec4 camRight;    // xyz = 右, w = 横の倍率（縦横比）
    vec4 camUp;       // xyz = 上
    vec4 sunDir;      // xyz = 太陽のある方向, w = 太陽の明るさ
    vec4 misc;        // x = 照明器具の光源の数, y = 空の明るさ, z = ガラス三角形の開始位置（三角形番号）, w = 1 画素の幅（1 m 先で何 m か）
    uvec4 frame;      // x = フレーム番号, y = 止まってからのフレーム数, z = 最大反射回数, w = 1 フレームで計算する回数
} pc;

const float PI = 3.14159265359;

// --- 乱数（PCG） ---------------------------------------------------------
uint rngState;
uint pcg(uint v) {
    uint state = v * 747796405u + 2891336453u;
    uint word = ((state >> ((state >> 28u) + 4u)) ^ state) * 277803737u;
    return (word >> 22u) ^ word;
}
float rand() { rngState = pcg(rngState); return float(rngState) / 4294967296.0; }

// --- 空 -----------------------------------------------------------------
// 簡単な空のモデル（天頂は青く、地平線は白っぽく、太陽の周りは明るい）。
vec3 skyRadiance(vec3 dir) {
    vec3 sun = normalize(pc.sunDir.xyz);
    float up = dir.z;
    if (up < 0.0) {
        // 地平線より下: 太陽と空に照らされた地面の明るさ（地面の色 × 地面に届く光 ÷ π）
        vec3 groundAlbedo = vec3(0.3, 0.29, 0.27);
        float irradiance = pc.sunDir.w * max(sun.z, 0.0) + PI * 0.6 * pc.misc.y;
        return groundAlbedo * irradiance / PI;
    }
    vec3 zenith = vec3(0.25, 0.45, 0.95);
    vec3 horizon = vec3(0.85, 0.9, 1.0);
    float t = pow(1.0 - up, 4.0);
    vec3 sky = mix(zenith, horizon, t);
    float cosSun = max(dot(dir, sun), 0.0);
    sky += vec3(1.0, 0.9, 0.7) * pow(cosSun, 8.0) * 0.6;   // 太陽の周りの明るさ
    float sunHeight = clamp(sun.z * 4.0, 0.0, 1.0);
    return sky * pc.misc.y * (0.2 + 0.8 * sunHeight);
}

// 太陽そのもの（見える円盤）。直接光はシャドウレイで別に計算するので、カメラから直接見たときだけ使う
vec3 sunDisc(vec3 dir) {
    float cosSun = dot(dir, normalize(pc.sunDir.xyz));
    return cosSun > 0.99996 ? vec3(1.0, 0.95, 0.85) * pc.sunDir.w * 2000.0 : vec3(0.0);
}

// --- 当たり判定 ---------------------------------------------------------
struct Hit {
    bool found;
    float t;
    uint tri;       // 全体での三角形番号
    vec2 bary;
    bool glass;
};

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
    vec4 camUp;       // xyz = 上, w = 地面の高さ（モデルの一番下）
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
// 青い空（天頂は濃く、地平線に向かって白っぽく霞む）＋太陽の周りの明るさ＋雲。
// 雲は向きだけで決まる模様なので、毎フレーム同じ絵になる（動いてもちらつかない）

float hash12(vec2 p) {
    vec3 p3 = fract(vec3(p.xyx) * 0.1031);
    p3 += dot(p3, p3.yzx + 33.33);
    return fract((p3.x + p3.y) * p3.z);
}

float valueNoise(vec2 p) {
    vec2 i = floor(p), f = fract(p);
    vec2 u = f * f * (3.0 - 2.0 * f);
    return mix(mix(hash12(i), hash12(i + vec2(1.0, 0.0)), u.x), mix(hash12(i + vec2(0.0, 1.0)), hash12(i + vec2(1.0, 1.0)), u.x), u.y);
}

// 細かさの違う模様を重ねたもの（雲のもこもこした形）
float fbm(vec2 p) {
    float v = 0.0, a = 0.5;
    for (int i = 0; i < 5; ++i) {
        v += a * valueNoise(p);
        p = p * 2.03 + vec2(17.1, 9.7);
        a *= 0.5;
    }
    return v;
}

// 太陽の色（低いほど夕日のように赤みを帯びる）× 明るさ
vec3 sunColor() {
    float h = normalize(pc.sunDir.xyz).z;
    vec3 c = mix(vec3(1.0, 0.55, 0.3), vec3(1.0, 0.9, 0.78), smoothstep(0.0, 0.45, h));
    return c * pc.sunDir.w;
}

// 雲の無い空（地面を照らす明るさの計算にも使う）
vec3 skyBase(vec3 dir) {
    vec3 sun = normalize(pc.sunDir.xyz);
    float up = max(dir.z, 0.0);
    // 日なたの面（色 × 太陽の明るさ）と同じくらいの明るさに見えるよう、空は明るめにする（暗いと影が青黒くなる）
    vec3 zenith = vec3(0.34, 0.62, 1.30);
    vec3 horizon = vec3(1.45, 1.55, 1.68);
    // 白っぽい霞は地平線のすぐ上だけ（少し上がるとすぐ青くなる）
    vec3 sky = mix(zenith, horizon, pow(1.0 - up, 10.0));
    float cosSun = max(dot(dir, sun), 0.0);
    sky += vec3(1.0, 0.85, 0.65) * (pow(cosSun, 8.0) * 0.45 + pow(cosSun, 64.0) * 0.6);  // 太陽の周りの明るさ
    float sunHeight = clamp(sun.z * 4.0, 0.0, 1.0);
    return sky * pc.misc.y * (0.15 + 0.85 * sunHeight);
}

// 地平線の霞の色（遠くの地面が溶け込む色）
vec3 hazeColor() { return skyBase(normalize(vec3(1.0, 0.0, 0.02))) * 0.95; }

vec3 skyRadiance(vec3 dir) {
    if (dir.z < 0.0) return hazeColor();  // 地面より下は霞の色（地面は別に計算する）
    vec3 sky = skyBase(dir);
    // 雲: 高さ一定の雲の層に、空の向きを写した位置の模様（遠くの雲ほど小さく、地平線に近づくほど密に見える）
    if (dir.z > 0.0) {
        vec2 p = dir.xy / (dir.z + 0.05) * 1.1;
        float d = fbm(p + vec2(3.7, 1.3));
        float cover = smoothstep(0.50, 0.68, d) * smoothstep(0.0, 0.04, dir.z);
        if (cover > 0.0) {
            vec3 sun = normalize(pc.sunDir.xyz);
            float sunHeight = clamp(sun.z * 4.0, 0.0, 1.0);
            // もこもこ感: 太陽の方へ少しずらした所の濃さと比べ、太陽側が薄い（＝光が通る）所を明るく、反対側を灰色にする
            float dTowardSun = fbm(p + vec2(3.7, 1.3) + normalize(sun.xy + vec2(1e-4)) * 0.08);
            float light = clamp(0.65 + (d - dTowardSun) * 6.0, 0.25, 1.15);
            float glow = 1.0 + 0.35 * pow(max(dot(dir, sun), 0.0), 4.0);  // 太陽に近い雲は明るい
            vec3 white = (vec3(0.96, 0.95, 0.93) * pc.misc.y * 2.3 + sunColor() * 0.12) * glow;
            vec3 shade = vec3(0.95, 1.0, 1.12) * pc.misc.y;
            vec3 cloud = mix(shade, white, light / 1.15);
            sky = mix(sky, cloud * (0.15 + 0.85 * sunHeight), cover);
        }
    }
    return sky;
}

// 太陽そのもの（見える円盤）。直接光はシャドウレイで別に計算するので、カメラから直接見たときだけ使う
vec3 sunDisc(vec3 dir) {
    float cosSun = dot(dir, normalize(pc.sunDir.xyz));
    return cosSun > 0.99996 ? sunColor() * 2000.0 : vec3(0.0);
}

// --- 当たり判定 ---------------------------------------------------------
struct Hit {
    bool found;
    float t;
    uint tri;       // 全体での三角形番号
    vec2 bary;
    bool glass;
};

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
    uvec4 look;       // x = 背景色（RGBA8。a が 0 以外なら空・雲を出さずにこの色）, y = 単色（RGBA8。a が 0 以外なら全部この色）,
                      // z = 1 なら地面（と地平線）をまっすぐには見せない, w = 地面の色（RGBA8。a が 0 なら既定の色）
} pc;

const float PI = 3.14159265359;

// 見た目の切り替え（ウィンドウ右上の「表示」の設定）
bool plainBackground() { return (pc.look.x >> 24u) != 0u; }
bool monochrome() { return (pc.look.y >> 24u) != 0u; }
bool hideGround() { return (pc.look.z & 1u) != 0u; }
// 地面の色（look.w。a が 0 なら既定の色）
vec3 groundAlbedo() {
    if ((pc.look.w >> 24u) == 0u) return vec3(0.44, 0.40, 0.37);
    vec3 c = vec3(float(pc.look.w & 255u), float((pc.look.w >> 8u) & 255u), float((pc.look.w >> 16u) & 255u)) / 255.0;
    return pow(c, vec3(2.2));
}
vec3 monoColor() {
    vec3 c = vec3(float(pc.look.y & 255u), float((pc.look.y >> 8u) & 255u), float((pc.look.y >> 16u) & 255u)) / 255.0;
    return pow(c, vec3(2.2));  // 選んだ色（sRGB）を明るさの値に直す
}

// --- 乱数（PCG） ---------------------------------------------------------
uint rngState;
uint pcg(uint v) {
    uint state = v * 747796405u + 2891336453u;
    uint word = ((state >> ((state >> 28u) + 4u)) ^ state) * 277803737u;
    return (word >> 22u) ^ word;
}
float rand() { rngState = pcg(rngState); return float(rngState) / 4294967296.0; }

// --- 空 -----------------------------------------------------------------
// 空と雲は起動時に sky.comp が 1 枚の画像（uSky）に描いておき、ここではそれを読むだけ（速い・毎フレーム同じ絵）。
// 画像: 横 = 方位、縦 = 高度 -2°〜90°（v = sqrt((高度 + 2°) / 92°)、上の行が天頂）
layout(set = 0, binding = 13, rgba16f) uniform readonly image2D uSky;

// 太陽の色（低いほど夕日のように赤みを帯びる）× 明るさ
vec3 sunColor() {
    float h = normalize(pc.sunDir.xyz).z;
    vec3 c = mix(vec3(1.0, 0.55, 0.3), vec3(1.0, 0.9, 0.78), smoothstep(0.0, 0.45, h));
    // 朝夕の低い太陽は、空気の中を長く通るので弱く赤い
    c *= mix(0.3, 1.0, smoothstep(0.0, 0.3, h));
    return c * pc.sunDir.w;
}

// 雲の無い空（地面を照らす明るさの計算に使う。sky.comp の atmosphere と同じ式）
vec3 skyBase(vec3 dir) {
    vec3 sun = normalize(pc.sunDir.xyz);
    float el = max(dir.z, 0.0);
    // 朝夕（太陽が低い）ほど、天頂は紫がかった濃い青、太陽の側の地平線は橙色になる
    float dusk = 1.0 - smoothstep(0.0, 0.35, sun.z);
    vec3 zenith = mix(vec3(0.20, 0.42, 1.05), vec3(0.16, 0.20, 0.55), dusk * 0.8);
    vec3 mid = mix(vec3(0.55, 0.78, 1.35), vec3(0.62, 0.62, 1.05), dusk * 0.7);
    vec3 horizon = vec3(1.45, 1.55, 1.68);
    float towardSun = dot(normalize(dir.xy + vec2(1e-5)), normalize(sun.xy + vec2(1e-5))) * 0.5 + 0.5;
    horizon = mix(horizon, mix(vec3(1.25, 1.0, 1.05), vec3(2.1, 1.05, 0.5), towardSun * towardSun), dusk);
    vec3 sky = mix(mid, zenith, smoothstep(0.0, 1.0, pow(el, 0.6)));
    sky = mix(sky, horizon, exp(-el * mix(14.0, 6.0, dusk)));
    float cosSun = max(dot(dir, sun), 0.0);
    sky += mix(vec3(1.0, 0.85, 0.65), vec3(1.2, 0.6, 0.3), dusk) * (pow(cosSun, 8.0) * 0.45 + pow(cosSun, 64.0) * 0.6);
    // 明るさ: 昼は 1、日の出・日の入りで約 0.4、太陽が地平線の下に沈むほど暗くなる（薄明）
    float t = clamp((sun.z + 0.2) / 0.45, 0.0, 1.0);
    return sky * pc.misc.y * (0.02 + 0.98 * t * t * (3.0 - 2.0 * t));
}

// 地平線の霞の色（遠くの地面が溶け込む色）
vec3 hazeColor() { return skyBase(normalize(vec3(1.0, 0.0, 0.02))) * 0.95; }

vec3 skyTexel(ivec2 p, ivec2 size) {
    p.x = (p.x % size.x + size.x) % size.x;  // 方位は一周つながっている
    p.y = clamp(p.y, 0, size.y - 1);
    return imageLoad(uSky, p).rgb;
}

vec3 skyRadiance(vec3 dir) {
    if (dir.z < -0.03) return hazeColor();  // 地面より下は霞の色（地面は別に計算する）
    ivec2 size = imageSize(uSky);
    const float MIN_EL = -2.0 * PI / 180.0, EL_RANGE = 92.0 * PI / 180.0;
    float el = asin(clamp(dir.z, -1.0, 1.0));
    float v = sqrt(clamp((el - MIN_EL) / EL_RANGE, 0.0, 1.0));
    float u = atan(dir.y, dir.x) / (2.0 * PI) + 0.5;
    vec2 pix = vec2(u * float(size.x), (1.0 - v) * float(size.y)) - 0.5;
    ivec2 b = ivec2(floor(pix));
    vec2 f = pix - vec2(b);
    vec3 a = mix(skyTexel(b, size), skyTexel(b + ivec2(1, 0), size), f.x);
    vec3 c = mix(skyTexel(b + ivec2(0, 1), size), skyTexel(b + ivec2(1, 1), size), f.x);
    return mix(a, c, f.y);
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

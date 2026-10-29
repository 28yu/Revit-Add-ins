// 共通部分（前半）: 入出力・乱数・空・当たり判定の結果の形。
// pathtrace_hw.comp / pathtrace_sw.comp から取り込まれる。
// 出力（ノイズ除去 denoise の材料）。色は「最初に当たった面の色」で割った明るさにしておき、
// ノイズ除去のあとで色を掛け戻す（模様や色の境目をぼかさないため）
layout(set = 0, binding = 1, rgba16f) uniform writeonly image2D uNoisy;   // rgb = 今回の明るさ（色で割ったもの）
layout(set = 0, binding = 2, rgba32f) uniform writeonly image2D uGPos;    // xyz = 最初に当たった位置, w = 距離（空なら -1）
layout(set = 0, binding = 7, rgba16f) uniform writeonly image2D uGNormal; // xyz = 面の向き
layout(set = 0, binding = 9, rgba8) uniform writeonly image2D uGAlbedo;   // rgb = 最初に当たった面の色

// 頂点は 1 つ 6 個の値（24 バイト）: 位置 x, y, z / 面の向き（2 つの 16 ビットに詰めたもの）/ 面の上の位置 u, v（メートル）
// （GPU のメモリを節約するため。src/renderer.cpp の packVertices と同じ並び）
layout(set = 0, binding = 3, std430) readonly buffer Vertices { float vdata[]; };
vec3 vertexPos(uint i) { uint b = i * 6u; return vec3(vdata[b], vdata[b + 1u], vdata[b + 2u]); }
vec2 vertexUv(uint i) { uint b = i * 6u; return vec2(vdata[b + 4u], vdata[b + 5u]); }
// 面の向き: 八面体に写して 2 つの 16 ビットに詰めたもの（x の下位 16 ビットが 0x8000 なら「向き無し」= 0 を返す）
vec3 vertexNormal(uint i) {
    uint packed = floatBitsToUint(vdata[i * 6u + 3u]);
    if ((packed & 0xFFFFu) == 0x8000u) return vec3(0.0);
    vec2 e = vec2(float(int(packed << 16) >> 16), float(int(packed) >> 16)) / 32767.0;
    vec3 n = vec3(e, 1.0 - abs(e.x) - abs(e.y));
    if (n.z < 0.0) n.xy = (1.0 - abs(n.yx)) * vec2(n.x >= 0.0 ? 1.0 : -1.0, n.y >= 0.0 ? 1.0 : -1.0);
    return normalize(n);
}
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

// 置き場所（インスタンス）と形（メッシュ）。頂点は形の座標で持ち、置き場所の行列で世界の座標に動かす
struct Instance {
    vec4 objectToWorld[3];  // 3 行 4 列（行ごと）
    vec4 worldToObject[3];
    vec4 normalMat[3];      // 向き（法線）を動かす 3 行 3 列
    uvec4 meta;             // x = 形の番号
};
layout(set = 0, binding = 14, std430) readonly buffer Instances { Instance instances[]; };
struct Mesh { uint firstTri; uint triCount; uint opaqueCount; uint bvhRoot; };
layout(set = 0, binding = 15, std430) readonly buffer Meshes { Mesh meshes[]; };

vec3 toWorldPoint(uint inst, vec3 p) {
    return vec3(dot(instances[inst].objectToWorld[0], vec4(p, 1.0)), dot(instances[inst].objectToWorld[1], vec4(p, 1.0)),
                dot(instances[inst].objectToWorld[2], vec4(p, 1.0)));
}
vec3 toWorldVector(uint inst, vec3 v) {
    return vec3(dot(instances[inst].objectToWorld[0].xyz, v), dot(instances[inst].objectToWorld[1].xyz, v),
                dot(instances[inst].objectToWorld[2].xyz, v));
}
vec3 toWorldNormal(uint inst, vec3 n) {
    return vec3(dot(instances[inst].normalMat[0].xyz, n), dot(instances[inst].normalMat[1].xyz, n),
                dot(instances[inst].normalMat[2].xyz, n));
}

layout(push_constant) uniform Params {
    vec4 camPos;      // xyz = 位置, w = 縦の画角の tan(半分)
    vec4 camForward;  // xyz = 前, w = パノラマなら 1
    vec4 camRight;    // xyz = 右, w = 横の倍率（縦横比）
    vec4 camUp;       // xyz = 上, w = 地面の高さ（モデルの一番下）
    vec4 sunDir;      // xyz = 太陽のある方向, w = 太陽の明るさ
    vec4 misc;        // x = 照明器具の光源の数, y = 空の明るさ, z = 1 なら空の光を直接狙う, w = 1 画素の幅（1 m 先で何 m か）
    uvec4 frame;      // x = フレーム番号, y = 止まってからのフレーム数, z = 最大反射回数, w = 1 フレームで計算する回数（下位 8 ビット）＋描き始める行（上位 24 ビット）
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

// 空と地面の光を直接狙う（misc.z が 1 のとき）。sky_cdf.comp が作った「明るさの地図」（方位 64 × sin(高度) 64 のマス。
// 上半分が空、下半分が地面）で、明るい所ほど多く狙う。どのマスも広さが同じなので、
// 向きの確率（立体角あたり）= マスの割合 × マスの数 / 4π
layout(set = 0, binding = 16, std430) readonly buffer SkyDist { float skyDist[]; };
const uint SKY_DIST_W = 64u, SKY_DIST_H = 64u;
bool skyLightSampling() { return pc.misc.z > 0.5; }

// 累積の割合の並び（first から count 個）から、x を超える最初の位置を探す
uint findCdf(uint first, uint count, float x) {
    uint lo = 0u, hi = count - 1u;
    while (lo < hi) {
        uint mid = (lo + hi) / 2u;
        if (skyDist[first + mid] <= x) lo = mid + 1u; else hi = mid;
    }
    return lo;
}

float envCellPdf(uint row, uint col) {
    return skyDist[SKY_DIST_H + SKY_DIST_W * SKY_DIST_H + row * SKY_DIST_W + col] * float(SKY_DIST_W * SKY_DIST_H) / (4.0 * PI);
}
uint envRow(float z) { return min(uint((clamp(z, -1.0, 1.0) * 0.5 + 0.5) * float(SKY_DIST_H)), SKY_DIST_H - 1u); }
float envRowCdf(uint rowEnd) { return rowEnd == 0u ? 0.0 : skyDist[rowEnd - 1u]; }  // 行 0〜rowEnd-1 の割合の合計

// 面の向き n から見える高さの範囲（行 rows.x〜rows.y-1）。面の裏側しか見えない行は狙わない（床なら空だけ、天井なら地面だけ）
uvec2 envRowsFor(vec3 n) {
    float side = sqrt(max(1.0 - n.z * n.z, 0.0));
    float zMin = n.z > 0.0 ? -side : -1.0, zMax = n.z < 0.0 ? side : 1.0;
    return uvec2(envRow(zMin), envRow(zMax) + 1u);
}

// 向き dir の確率（立体角あたり）。rows の範囲の中だけで選んだときの値（範囲外は 0）
float skyPdf(vec3 dir, uvec2 rows) {
    uint row = envRow(dir.z);
    if (row < rows.x || row >= rows.y) return 0.0;
    float mass = envRowCdf(rows.y) - envRowCdf(rows.x);
    if (mass <= 0.0) return 0.0;
    float u = atan(dir.y, dir.x) / (2.0 * PI) + 0.5;
    uint col = min(uint(u * float(SKY_DIST_W)), SKY_DIST_W - 1u);
    return envCellPdf(row, col) / mass;
}

// rows の範囲から向きを 1 つ選ぶ（r1〜r4 は 0〜1 の乱数）。戻り値は向きの確率（立体角あたり。0 なら選べなかった）
float sampleSky(uvec2 rows, float r1, float r2, float r3, float r4, out vec3 dir) {
    dir = vec3(0.0, 0.0, 1.0);
    float c0 = envRowCdf(rows.x), c1 = envRowCdf(rows.y);
    if (c1 - c0 <= 0.0) return 0.0;
    uint row = clamp(findCdf(0u, SKY_DIST_H, c0 + r1 * (c1 - c0)), rows.x, rows.y - 1u);
    uint col = findCdf(SKY_DIST_H + row * SKY_DIST_W, SKY_DIST_W, r2);
    float s = clamp((float(row) + r3) / float(SKY_DIST_H) * 2.0 - 1.0, -1.0, 1.0);  // sin（高度）
    float u = (float(col) + r4) / float(SKY_DIST_W);
    float a = (u - 0.5) * 2.0 * PI;
    float c = sqrt(max(1.0 - s * s, 0.0));
    dir = vec3(c * cos(a), c * sin(a), s);
    return envCellPdf(row, col) / (c1 - c0);
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
    uint inst;      // 置き場所の番号
    vec2 bary;
    bool glass;
};

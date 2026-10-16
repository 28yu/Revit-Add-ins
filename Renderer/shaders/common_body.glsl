// 共通部分（後半）: 表面の情報・反射・光の通り道・トーンマッピング・main。
// これより前に traceClosest(origin, dir) と occluded(origin, dir) が定義されている必要がある。
void surfaceAt(Hit h, vec3 rayDir, out vec3 pos, out vec3 normal, out vec3 geomNormal) {
    uint i0 = indices[h.tri * 3u], i1 = indices[h.tri * 3u + 1u], i2 = indices[h.tri * 3u + 2u];
    vec3 p0 = vertices[i0].position.xyz, p1 = vertices[i1].position.xyz, p2 = vertices[i2].position.xyz;
    float b0 = 1.0 - h.bary.x - h.bary.y;
    pos = p0 * b0 + p1 * h.bary.x + p2 * h.bary.y;
    geomNormal = normalize(cross(p1 - p0, p2 - p0));
    vec3 n = vertices[i0].normal.xyz * b0 + vertices[i1].normal.xyz * h.bary.x + vertices[i2].normal.xyz * h.bary.y;
    normal = dot(n, n) > 1e-8 ? normalize(n) : geomNormal;
    // 裏から当たったときは表向きにそろえる（Revit の形は裏表が混ざっていることがある）
    if (dot(geomNormal, rayDir) > 0.0) geomNormal = -geomNormal;
    if (dot(normal, geomNormal) < 0.0) normal = -normal;
}

// --- 反射の向きを選ぶ ---------------------------------------------------
void basis(vec3 n, out vec3 t, out vec3 b) {
    float s = n.z >= 0.0 ? 1.0 : -1.0;
    float a = -1.0 / (s + n.z);
    float c = n.x * n.y * a;
    t = vec3(1.0 + s * n.x * n.x * a, s * c, -s * n.x);
    b = vec3(c, s + n.y * n.y * a, -n.y);
}

vec3 cosineHemisphere(vec3 n) {
    float r1 = rand(), r2 = rand();
    float r = sqrt(r1), phi = 2.0 * PI * r2;
    vec3 t, b;
    basis(n, t, b);
    return normalize(t * (r * cos(phi)) + b * (r * sin(phi)) + n * sqrt(max(0.0, 1.0 - r1)));
}

// GGX（つやのある面の反射）の向きを選ぶ
vec3 ggxReflect(vec3 v, vec3 n, float roughness) {
    float a = max(roughness * roughness, 0.001);
    float r1 = rand(), r2 = rand();
    float phi = 2.0 * PI * r1;
    float cosTheta = sqrt((1.0 - r2) / (1.0 + (a * a - 1.0) * r2));
    float sinTheta = sqrt(max(0.0, 1.0 - cosTheta * cosTheta));
    vec3 t, b;
    basis(n, t, b);
    vec3 h = normalize(t * (sinTheta * cos(phi)) + b * (sinTheta * sin(phi)) + n * cosTheta);
    return reflect(v, h);
}

float fresnelSchlick(float cosTheta, float f0) {
    return f0 + (1.0 - f0) * pow(1.0 - clamp(cosTheta, 0.0, 1.0), 5.0);
}

// --- 1本の光の通り道を追う -----------------------------------------------
vec3 radiance(vec3 origin, vec3 dir) {
    vec3 result = vec3(0.0);
    vec3 throughput = vec3(1.0);
    vec3 sun = normalize(pc.sunDir.xyz);
    vec3 sunColor = vec3(1.0, 0.95, 0.88) * pc.sunDir.w;
    bool specularPath = true;  // カメラから鏡のような反射だけで来た（太陽の円盤を見せてよい）

    for (uint bounce = 0u; bounce <= pc.frame.z; ++bounce) {
        Hit h = traceClosest(origin, dir);
        if (!h.found) {
            result += throughput * (skyRadiance(dir) + (specularPath ? sunDisc(dir) : vec3(0.0)));
            break;
        }
        vec3 pos, n, gn;
        surfaceAt(h, dir, pos, n, gn);
        Material m = materials[triMaterials[h.tri]];
        vec3 albedo = m.color.rgb;
        float roughness = clamp(m.params.x, 0.02, 1.0);
        float metallic = m.params.y;
        result += throughput * albedo * m.params.z;  // 自己発光（照明器具など）

        if (h.glass) {
            // 薄いガラスとして扱う: 表面で反射するか、そのまま通り抜けるか
            float f = fresnelSchlick(abs(dot(-dir, n)), 0.04);
            if (rand() < f) {
                dir = reflect(dir, gn);
            } else {
                throughput *= mix(vec3(1.0), albedo, 0.3);  // ガラスのわずかな色
            }
            origin = pos + dir * 0.001;
            continue;
        }

        vec3 v = dir;
        float cosV = max(dot(-v, n), 0.0);
        // 金属は色付きの反射だけ、金属でない物は「反射（白）」か「拡散（物の色）」を確率で選ぶ
        float specProb = metallic > 0.5 ? 1.0 : fresnelSchlick(cosV, 0.04) * (1.0 - roughness * 0.7) + 0.04;
        specProb = clamp(specProb, 0.0, 1.0);

        // 太陽の直接光（拡散面のみ。影の判定は専用回路で行う）
        float nDotL = dot(n, sun);
        if (nDotL > 0.0 && sun.z > -0.05 && metallic <= 0.5) {
            if (!occluded(pos + gn * 0.002, sun)) {
                result += throughput * (1.0 - specProb) * albedo * sunColor * nDotL;  // 拡散面の明るさ = 色 × 光の強さ × 光の当たる角度
            }
        }

        if (rand() < specProb) {
            vec3 newDir = ggxReflect(v, n, roughness);
            if (dot(newDir, gn) <= 0.0) break;
            throughput *= metallic > 0.5 ? albedo : vec3(1.0);
            dir = newDir;
            // ほぼ鏡のときだけ太陽の円盤を映す（ざらついた面では直接光として計算済み）
            specularPath = specularPath && roughness < 0.1;
        } else {
            dir = cosineHemisphere(n);
            throughput *= albedo;
            specularPath = false;
        }
        origin = pos + gn * 0.002;

        // 光がほとんど残っていなければ打ち切る
        float p = max(throughput.r, max(throughput.g, throughput.b));
        if (bounce >= 2u) {
            if (rand() > p) break;
            throughput /= p;
        }
    }
    return result;
}

// ACES に近いトーンマッピング（明るい所を白飛びさせずに表示する）
vec3 tonemap(vec3 x) {
    const float a = 2.51, b = 0.03, c = 2.43, d = 0.59, e = 0.14;
    return clamp((x * (a * x + b)) / (x * (c * x + d) + e), 0.0, 1.0);
}

// 自動露出用: 8×8 画素（1 つのまとまり）の明るさの平均を共有メモリで求める
shared float sharedLum[64];

void main() {
    ivec2 size = imageSize(uAccum);
    ivec2 pixel = ivec2(gl_GlobalInvocationID.xy);
    // 画面の外の画素も途中で終わらせない（下の barrier() に全員がたどり着く必要があるため）
    bool inside = pixel.x < size.x && pixel.y < size.y;
    float lumForExposure = -1.0;

    if (inside) {
        rngState = pcg(uint(pixel.x) + uint(pixel.y) * 9781u + pc.frame.x * 6271u * 26699u);

        // ピクセル内の位置を毎回少しずらして、ぎざぎざを消す
        vec2 jitter = pc.frame.y == 0u ? vec2(0.5) : vec2(rand(), rand());
        vec2 uv = (vec2(pixel) + jitter) / vec2(size);
        vec3 dir;
        if (pc.camForward.w > 0.5) {
            // 360°パノラマ（正距円筒図法）: 横 = 方位 -180°〜180°、縦 = 上 90°〜下 -90°
            float lon = (uv.x - 0.5) * 2.0 * PI;
            float lat = (0.5 - uv.y) * PI;
            vec3 f = pc.camForward.xyz, r = pc.camRight.xyz, u = pc.camUp.xyz;
            dir = normalize(cos(lat) * (cos(lon) * f + sin(lon) * r) + sin(lat) * u);
        } else {
            vec2 ndc = vec2(uv.x * 2.0 - 1.0, 1.0 - uv.y * 2.0);
            float th = pc.camPos.w;
            dir = normalize(pc.camForward.xyz + pc.camRight.xyz * ndc.x * th * pc.camRight.w + pc.camUp.xyz * ndc.y * th);
        }

        vec3 color = radiance(pc.camPos.xyz, dir);
        // 極端に明るい点（ホタル）を抑える
        float lum = dot(color, vec3(0.2126, 0.7152, 0.0722));
        if (lum > 50.0) color *= 50.0 / lum;

        vec3 accum = color;
        if (pc.frame.y > 0u) accum += imageLoad(uAccum, pixel).rgb;
        imageStore(uAccum, pixel, vec4(accum, 1.0));

        vec3 avg = accum / float(pc.frame.y + 1u);
        vec3 mapped = tonemap(avg * pc.misc.x);
        imageStore(uOutput, pixel, vec4(pow(mapped, vec3(1.0 / 2.2)), 1.0));
        lumForExposure = dot(avg, vec3(0.2126, 0.7152, 0.0722));
    }

    // 自動露出: 1 画素ずつだと、計算回数が少ない間のざらつき（真っ黒な点）で平均が大きく下がり、
    // 明るくしすぎてしまう。8×8 画素の平均をとってから対数にする
    sharedLum[gl_LocalInvocationIndex] = lumForExposure;
    barrier();
    if (gl_LocalInvocationIndex == 0u) {
        float sum = 0.0;
        int n = 0;
        for (int i = 0; i < 64; ++i) {
            if (sharedLum[i] >= 0.0) { sum += sharedLum[i]; n++; }
        }
        if (n > 0) {
            float l = sum / float(n);
            int q = int(clamp(log2(max(l, 1e-5)), -16.0, 16.0) * 64.0);  // 1/64 段刻み（8K でも整数が溢れない）
            atomicAdd(logLumSum, q);
            atomicAdd(logLumCount, 1);
        }
    }
}

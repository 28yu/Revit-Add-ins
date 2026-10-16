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
// 最初に当たった面の情報（ノイズ除去で使う）
struct Primary {
    bool hit;
    vec3 pos;
    vec3 normal;
    vec3 albedo;
    float dist;
};

vec3 radiance(vec3 origin, vec3 dir, out Primary primary) {
    primary.hit = false;
    primary.pos = origin + dir * 1e5;
    primary.normal = -dir;
    primary.albedo = vec3(1.0);
    primary.dist = -1.0;
    vec3 result = vec3(0.0);
    vec3 throughput = vec3(1.0);
    vec3 sun = normalize(pc.sunDir.xyz);
    vec3 sunColor = vec3(1.0, 0.95, 0.88) * pc.sunDir.w;
    bool specularPath = true;  // カメラから鏡のような反射だけで来た（太陽の円盤を見せてよい）
    // ノイズ除去で使う「最初に見えている面」は、ガラスを素通りした先の物にする
    // （ガラスの面を使うと、動いたときに窓の外の景色がにじむため）
    float travelled = 0.0;

    for (uint bounce = 0u; bounce <= pc.frame.z; ++bounce) {
        Hit h = traceClosest(origin, dir);
        if (!h.found) {
            result += throughput * (skyRadiance(dir) + (specularPath ? sunDisc(dir) : vec3(0.0)));
            break;
        }
        vec3 pos, n, gn;
        surfaceAt(h, dir, pos, n, gn);
        Material m = materials[triMaterials[h.tri]];
        if (!primary.hit && !h.glass) {
            primary.hit = true;
            primary.pos = pos;
            primary.normal = n;
            primary.dist = travelled + h.t;
            // 鏡・金属は、映り込みそのものが見えているので色で割らない
            primary.albedo = m.params.y > 0.5 ? vec3(1.0) : max(m.color.rgb, vec3(0.02));
        }
        vec3 albedo = m.color.rgb;
        float roughness = clamp(m.params.x, 0.02, 1.0);
        float metallic = m.params.y;
        result += throughput * albedo * m.params.z;  // 自己発光（照明器具など）

        if (h.glass) {
            // 薄いガラスとして扱う。反射するか通り抜けるかを毎回ランダムに選ぶと、窓に点々のちらつきが出るので、
            // 「必ず通り抜け、表面の映り込みは空の明るさ × 反射率で足す」決まった計算にする
            // （室内側から見た窓への部屋の映り込みは省く）
            float f = fresnelSchlick(abs(dot(-dir, n)), 0.04);
            travelled += h.t;
            result += throughput * f * skyRadiance(reflect(dir, gn));
            throughput *= (1.0 - f) * mix(vec3(1.0), albedo, 0.3);  // ガラスのわずかな色
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

void main() {
    ivec2 size = imageSize(uNoisy);
    ivec2 pixel = ivec2(gl_GlobalInvocationID.xy);
    if (pixel.x >= size.x || pixel.y >= size.y) return;
    rngState = pcg(uint(pixel.x) + uint(pixel.y) * 9781u + pc.frame.x * 6271u * 26699u);

    // 1 フレームで何回か計算して平均する（速い GPU では回数を増やし、動いている間の画質を上げる）
    uint spp = max(pc.frame.w, 1u);
    vec3 color = vec3(0.0);
    Primary primary;
    for (uint s = 0u; s < spp; ++s) {
        // ピクセル内の位置を毎回少しずらして、ぎざぎざを消す（止まって最初の 1 回だけは中心）
        vec2 jitter = (pc.frame.y == 0u && s == 0u) ? vec2(0.5) : vec2(rand(), rand());
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

        Primary pr;
        vec3 c = radiance(pc.camPos.xyz, dir, pr);
        if (s == 0u) primary = pr;  // 面の情報は 1 回目のものを使う
        // 極端に明るい点（ホタル）を抑える
        float lum = dot(c, vec3(0.2126, 0.7152, 0.0722));
        if (lum > 50.0) c *= 50.0 / lum;
        color += c;
    }
    color /= float(spp);

    imageStore(uNoisy, pixel, vec4(color / primary.albedo, 1.0));
    imageStore(uGPos, pixel, vec4(primary.pos, primary.dist));
    imageStore(uGNormal, pixel, vec4(primary.normal, 0.0));
    imageStore(uGAlbedo, pixel, vec4(primary.albedo, 1.0));
}

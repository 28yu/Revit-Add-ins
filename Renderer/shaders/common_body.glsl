// 共通部分（後半）: 表面の情報・反射・光の通り道・トーンマッピング・main。
// これより前に traceClosest(origin, dir) と occluded(origin, dir) が定義されている必要がある。
// uv = テクスチャ上の位置（1 = 画像 1 枚分）, uvPerMeter = 1 m で何枚分進むか（遠くで縮小した段を選ぶのに使う）
void surfaceAt(Hit h, vec3 rayDir, out vec3 pos, out vec3 normal, out vec3 geomNormal, out vec2 uv, out float uvPerMeter) {
    uint i0 = indices[h.tri * 3u], i1 = indices[h.tri * 3u + 1u], i2 = indices[h.tri * 3u + 2u];
    vec4 v0 = vertices[i0].position, v1 = vertices[i1].position, v2 = vertices[i2].position;
    vec3 p0 = v0.xyz, p1 = v1.xyz, p2 = v2.xyz;
    vec2 t0 = vec2(v0.w, vertices[i0].normal.w), t1 = vec2(v1.w, vertices[i1].normal.w), t2 = vec2(v2.w, vertices[i2].normal.w);
    float b0 = 1.0 - h.bary.x - h.bary.y;
    pos = p0 * b0 + p1 * h.bary.x + p2 * h.bary.y;
    uv = t0 * b0 + t1 * h.bary.x + t2 * h.bary.y;
    vec3 cr = cross(p1 - p0, p2 - p0);
    vec2 e1 = t1 - t0, e2 = t2 - t0;
    float uvArea = abs(e1.x * e2.y - e1.y * e2.x);
    uvPerMeter = sqrt(uvArea / max(length(cr), 1e-12));
    geomNormal = normalize(cr);
    vec3 n = vertices[i0].normal.xyz * b0 + vertices[i1].normal.xyz * h.bary.x + vertices[i2].normal.xyz * h.bary.y;
    normal = dot(n, n) > 1e-8 ? normalize(n) : geomNormal;
    // 裏から当たったときは表向きにそろえる（Revit の形は裏表が混ざっていることがある）
    if (dot(geomNormal, rayDir) > 0.0) geomNormal = -geomNormal;
    if (dot(normal, geomNormal) < 0.0) normal = -normal;
}

void basis(vec3 n, out vec3 t, out vec3 b);  // 下で定義

// 三角形の上で、面の上の位置 u, v が増える向き（凹凸の計算に使う）。求められなければ面の向きから適当に決める
void tangentFrame(Hit h, vec3 n, out vec3 tu, out vec3 tv) {
    uint i0 = indices[h.tri * 3u], i1 = indices[h.tri * 3u + 1u], i2 = indices[h.tri * 3u + 2u];
    vec4 v0 = vertices[i0].position, v1 = vertices[i1].position, v2 = vertices[i2].position;
    vec3 dp1 = v1.xyz - v0.xyz, dp2 = v2.xyz - v0.xyz;
    vec2 d1 = vec2(v1.w, vertices[i1].normal.w) - vec2(v0.w, vertices[i0].normal.w);
    vec2 d2 = vec2(v2.w, vertices[i2].normal.w) - vec2(v0.w, vertices[i0].normal.w);
    float det = d1.x * d2.y - d1.y * d2.x;
    if (abs(det) < 1e-12) {
        basis(n, tu, tv);
        return;
    }
    tu = (dp1 * d2.y - dp2 * d1.y) / det;
    tv = (dp2 * d1.x - dp1 * d2.x) / det;
    // 面の向きに直角にそろえる
    tu = normalize(tu - n * dot(n, tu));
    tv = normalize(tv - n * dot(n, tv) - tu * dot(tu, tv));
}

// --- テクスチャ ---------------------------------------------------------
// srgb: 色の画像なら true（sRGB → リニアに直す）、凹凸の画像なら false（そのままの値）
vec4 texel(TexInfo ti, uint offset, uint w, uint h, ivec2 p, bool srgb) {
    // 繰り返し貼る（画像の外は反対側から続ける）
    uint x = uint(p.x - int(w) * int(floor(float(p.x) / float(w))));
    uint y = uint(p.y - int(h) * int(floor(float(p.y) / float(h))));
    uint c = texels[offset + min(y, h - 1u) * w + min(x, w - 1u)];
    vec4 s = vec4(float(c & 255u), float((c >> 8) & 255u), float((c >> 16) & 255u), float(c >> 24)) / 255.0;
    if (!srgb) return s;
    // sRGB → リニア
    vec3 lin = mix(s.rgb / 12.92, pow((s.rgb + 0.055) / 1.055, vec3(2.4)), step(vec3(0.04045), s.rgb));
    return vec4(lin, s.a);
}

// 段 level の画像から、周りの 4 画素を距離に応じて混ぜて取り出す（バイリニア）
vec4 sampleLevel(TexInfo ti, int level, vec2 uv, bool srgb) {
    uint offset = ti.offset, w = ti.width, h = ti.height;
    for (int i = 0; i < level; ++i) {
        offset += w * h;
        w = max(w / 2u, 1u);
        h = max(h / 2u, 1u);
    }
    // 画像の上が v = 1（Revit と同じ向き）
    vec2 p = vec2(uv.x, 1.0 - uv.y) * vec2(w, h) - 0.5;
    vec2 base = floor(p);
    vec2 f = p - base;
    ivec2 b = ivec2(base);
    vec4 a = mix(texel(ti, offset, w, h, b, srgb), texel(ti, offset, w, h, b + ivec2(1, 0), srgb), f.x);
    vec4 c = mix(texel(ti, offset, w, h, b + ivec2(0, 1), srgb), texel(ti, offset, w, h, b + ivec2(1, 1), srgb), f.x);
    return mix(a, c, f.y);
}

// 縮小した段のどれを使うか。footprint = この点で 1 画素（または光の広がり）が覆うテクスチャの幅（画像 1 枚 = 1）
float textureLod(int index, float footprint) {
    TexInfo ti = texInfos[index];
    float lod = log2(max(footprint * float(max(ti.width, ti.height)), 1e-8));
    return clamp(lod, 0.0, float(ti.levels - 1u));
}

vec4 sampleTextureLod(int index, vec2 uv, float lod, bool srgb) {
    TexInfo ti = texInfos[index];
    int l0 = int(floor(lod));
    int l1 = min(l0 + 1, int(ti.levels) - 1);
    return mix(sampleLevel(ti, l0, uv, srgb), sampleLevel(ti, l1, uv, srgb), lod - float(l0));
}

// 貼り方（2 行 3 列）で、面の上の位置（メートル）→ 画像上の位置にする
vec2 mapUv(vec4 row0, vec4 row1, vec2 uv) { return vec2(dot(row0.xyz, vec3(uv, 1.0)), dot(row1.xyz, vec3(uv, 1.0))); }
// 貼り方の拡大率（1 m が画像の何枚分か）
float mapScale(vec4 row0, vec4 row1) { return sqrt(abs(row0.x * row1.y - row0.y * row1.x)); }

// 凹凸（レリーフ パターン）で面の向きを傾ける。footprint はメートル単位の光の筋の太さ
vec3 bumpNormal(Material m, Hit h, vec3 n, vec2 uv, float footprintM) {
    vec3 tu, tv;
    tangentFrame(h, n, tu, tv);
    vec2 uvB = mapUv(m.bumpMap0, m.bumpMap1, uv);
    float lod = textureLod(m.tex.y, footprintM * mapScale(m.bumpMap0, m.bumpMap1));
    float strength = m.colorMap0.w;
    if (m.tex.z == 1) {
        // 法線の画像: 色 (r, g, b) がそのまま向き（画像の横・縦・面の外向き）
        vec3 t = sampleTextureLod(m.tex.y, uvB, lod, false).rgb * 2.0 - 1.0;
        t.xy *= strength;
        vec3 nb = normalize(tu * t.x * sign(m.bumpMap0.x) + tv * t.y * sign(m.bumpMap1.y) + n * max(t.z, 0.05));
        return nb;
    }
    // 高さの画像: 隣の位置との高さの差から傾きを求める（差を取る間隔は使っている段の 1 画素）
    TexInfo ti = texInfos[m.tex.y];
    float du = exp2(floor(lod)) / float(ti.width), dv = exp2(floor(lod)) / float(ti.height);
    float h0 = sampleTextureLod(m.tex.y, uvB, lod, false).r;
    float hu = sampleTextureLod(m.tex.y, uvB + vec2(du, 0.0), lod, false).r;
    float hv = sampleTextureLod(m.tex.y, uvB + vec2(0.0, dv), lod, false).r;
    vec2 g = vec2((hu - h0) / du, (hv - h0) / dv) * strength;  // 画像 1 枚あたりの高さの変化（メートル）
    // 画像上の向き → 面の上の向き（メートルあたり）に直す
    float dhdu = g.x * m.bumpMap0.x + g.y * m.bumpMap1.x;
    float dhdv = g.x * m.bumpMap0.y + g.y * m.bumpMap1.y;
    return normalize(n - tu * dhdu - tv * dhdv);
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

// つるつるの鏡（金属で粗さが小さい）
bool isMirror(Material m) { return m.params.y > 0.5 && m.params.x < 0.1; }

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
    // 光の筋の太さ（m）と広がり方。遠く・ざらついた反射の先ほど太くなり、テクスチャの縮小した段を使う
    float coneWidth = 0.0;
    float coneSpread = pc.misc.w;
    // 鏡: 映っている物を「鏡の向こうにある物（虚像）」として最初に見えている面にする。
    // 動いたときに前のフレームと位置を合わせられるので、鏡の映り込みが残像にならない
    vec3 camOrigin = origin, camDir = dir;
    bool viaMirror = false;
    vec3 mirrorN = vec3(0.0), mirrorTint = vec3(1.0);

    for (uint bounce = 0u; bounce <= pc.frame.z; ++bounce) {
        Hit h = traceClosest(origin, dir);
        if (!h.found) {
            result += throughput * (skyRadiance(dir) + (specularPath ? sunDisc(dir) : vec3(0.0)));
            break;
        }
        vec3 pos, n, gn;
        vec2 uv;
        float uvPerMeter;
        surfaceAt(h, dir, pos, n, gn, uv, uvPerMeter);
        coneWidth += h.t * coneSpread;
        Material m = materials[triMaterials[h.tri]];
        vec3 albedo = m.color.rgb;
        float slant = max(abs(dot(dir, gn)), 0.2);  // 斜めに見ると引き伸ばされる
        float footprintM = coneWidth / slant * uvPerMeter;  // 光の筋が面の上で覆う幅（面の上の位置の単位 = メートル）
        if (m.tex.x >= 0) {
            vec2 uvC = mapUv(m.colorMap0, m.colorMap1, uv);
            float lod = textureLod(m.tex.x, footprintM * mapScale(m.colorMap0, m.colorMap1));
            albedo *= sampleTextureLod(m.tex.x, uvC, lod, true).rgb;
        }
        if (m.tex.y >= 0 && !h.glass) {
            vec3 nb = bumpNormal(m, h, n, uv, footprintM);
            if (dot(nb, gn) > 0.05) n = nb;  // 裏を向くほど傾いたときは使わない
        }
        if (!primary.hit && !h.glass) {
            if (isMirror(m) && !viaMirror) {
                // 鏡（1 枚目だけ）: この先に見える物を最初に見えている面にする
                viaMirror = true;
                mirrorN = n;
                mirrorTint = albedo;
                travelled += h.t;
            } else {
                primary.hit = true;
                primary.dist = travelled + h.t;
                // 鏡・金属は、映り込みそのものが見えているので色で割らない
                vec3 a = m.params.y > 0.5 ? vec3(1.0) : max(albedo, vec3(0.02));
                if (viaMirror) {
                    // 虚像: カメラから見て、鏡までの距離＋鏡からの距離だけ奥にある。向きも鏡で折り返す
                    primary.pos = camOrigin + camDir * primary.dist;
                    primary.normal = reflect(n, mirrorN);
                    primary.albedo = max(mirrorTint * a, vec3(0.02));
                } else {
                    primary.pos = pos;
                    primary.normal = n;
                    primary.albedo = a;
                }
            }
        }
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
        float specProb = metallic > 0.5 ? 1.0 : fresnelSchlick(cosV, clamp(m.params.w, 0.0, 1.0)) * (1.0 - roughness * 0.7) + 0.04;
        specProb = clamp(specProb, 0.0, 1.0);

        // 太陽の直接光（拡散面のみ。影の判定は専用回路で行う）
        float nDotL = dot(n, sun);
        if (nDotL > 0.0 && sun.z > -0.05 && metallic <= 0.5) {
            if (!occluded(pos + gn * 0.002, sun)) {
                result += throughput * (1.0 - specProb) * albedo * sunColor * nDotL;  // 拡散面の明るさ = 色 × 光の強さ × 光の当たる角度
            }
        }

        // 照明器具の光（拡散面のみ）。毎回 1 つの光源を選んで、その光が届くか影の判定をする
        uint lightCount = uint(pc.misc.x);
        if (lightCount > 0u && metallic <= 0.5) {
            uint li = min(uint(rand() * float(lightCount)), lightCount - 1u);
            Light L = lights[li];
            float radius = L.positionRadius.w;
            // 光る部分（球）の中の 1 点を狙う（影の縁がぼける）
            vec3 jitter = (vec3(rand(), rand(), rand()) * 2.0 - 1.0) * radius;
            vec3 toL = L.positionRadius.xyz + jitter - pos;
            float d2 = max(dot(toL, toL), radius * radius + 1e-4);
            float d = sqrt(d2);
            vec3 l = toL / d;
            float nDotLl = dot(n, l);
            if (nDotLl > 0.0 && dot(gn, l) > 0.0) {
                float k = 1.0;
                float c = dot(-l, L.direction.xyz);
                if (L.intensityType.w > 1.5) k = smoothstep(L.cone.x, L.direction.w, c);  // スポット
                else if (L.intensityType.w > 0.5) k = c > 0.0 ? 1.0 : 0.0;           // 半球（下向きなど）
                // 照明器具そのもの（光源のすぐ近くのかさ・電球の形）は影にしない
                float reach = d - radius * 1.5 - 0.05;
                if (k > 0.0 && (reach <= 0.0 || !occludedDist(pos + gn * 0.002, l, reach))) {
                    result += throughput * (1.0 - specProb) * albedo * L.intensityType.rgb * (k * nDotLl / d2 * float(lightCount));
                }
            }
        }

        if (rand() < specProb) {
            vec3 newDir = ggxReflect(v, n, roughness);
            if (dot(newDir, gn) <= 0.0) break;
            throughput *= metallic > 0.5 ? albedo : vec3(1.0);
            dir = newDir;
            coneSpread += roughness * roughness * 0.5;
            // ほぼ鏡のときだけ太陽の円盤を映す（ざらついた面では直接光として計算済み）
            specularPath = specularPath && roughness < 0.1;
        } else {
            dir = cosineHemisphere(n);
            throughput *= albedo;
            coneSpread = max(coneSpread, 0.5);
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

// 画素の中心を通る光線で、最初に見えている面の位置・向きを求める（ガラスは素通りする）。
// ノイズ除去で「同じ面か」を比べるのに使うので、毎フレーム同じ結果になるよう、狙う位置をずらさない
// mirror: ざらついた金属のように、見る位置で映り込みが変わる度合い（0〜1）。動いている間の残像を防ぐのに使う。
// つるつるの鏡は、映っている物の虚像（鏡の向こうにあるとみなした位置・向き）を返すので 0
void primaryGeometry(vec3 origin, vec3 dir, out vec3 pos, out vec3 normal, out float dist, out float mirror) {
    pos = origin + dir * 1e5;
    normal = -dir;
    dist = -1.0;
    mirror = 0.0;
    vec3 camOrigin = origin, camDir = dir;
    bool viaMirror = false;
    vec3 mirrorN = vec3(0.0);
    float travelled = 0.0;
    for (int i = 0; i < 8; ++i) {
        Hit h = traceClosest(origin, dir);
        if (!h.found) return;
        vec3 p, n, gn;
        vec2 uv;
        float uvPerMeter;
        surfaceAt(h, dir, p, n, gn, uv, uvPerMeter);
        travelled += h.t;
        if (!h.glass) {
            Material m = materials[triMaterials[h.tri]];
            if (isMirror(m) && !viaMirror) {
                viaMirror = true;
                mirrorN = n;
                dir = reflect(dir, n);
                origin = p + gn * 0.002;
                continue;
            }
            pos = viaMirror ? camOrigin + camDir * travelled : p;
            normal = viaMirror ? reflect(n, mirrorN) : n;
            dist = travelled;
            if (m.params.y > 0.5) mirror = clamp(1.0 - m.params.x * 2.0, 0.0, 1.0);
            return;
        }
        origin = p + dir * 0.001;
    }
}

vec3 cameraRay(vec2 uv) {
    if (pc.camForward.w > 0.5) {
        // 360°パノラマ（正距円筒図法）: 横 = 方位 -180°〜180°、縦 = 上 90°〜下 -90°
        float lon = (uv.x - 0.5) * 2.0 * PI;
        float lat = (0.5 - uv.y) * PI;
        vec3 f = pc.camForward.xyz, r = pc.camRight.xyz, u = pc.camUp.xyz;
        return normalize(cos(lat) * (cos(lon) * f + sin(lon) * r) + sin(lat) * u);
    }
    vec2 ndc = vec2(uv.x * 2.0 - 1.0, 1.0 - uv.y * 2.0);
    float th = pc.camPos.w;
    return normalize(pc.camForward.xyz + pc.camRight.xyz * ndc.x * th * pc.camRight.w + pc.camUp.xyz * ndc.y * th);
}

void main() {
    ivec2 size = imageSize(uNoisy);
    ivec2 pixel = ivec2(gl_GlobalInvocationID.xy);
    if (pixel.x >= size.x || pixel.y >= size.y) return;
    rngState = pcg(uint(pixel.x) + uint(pixel.y) * 9781u + pc.frame.x * 6271u * 26699u);

    // 1 フレームで何回か計算して平均する（速い GPU では回数を増やし、動いている間の画質を上げる）
    uint spp = max(pc.frame.w, 1u);
    vec3 demod = vec3(0.0);   // 色で割った明るさ（1 回ずつ、その回の面の色で割ってから平均する）
    vec3 albedo = vec3(0.0);  // 面の色の平均
    for (uint s = 0u; s < spp; ++s) {
        // ピクセル内の位置を毎回少しずらして、ぎざぎざを消す（止まって最初の 1 回だけは中心）
        vec2 jitter = (pc.frame.y == 0u && s == 0u) ? vec2(0.5) : vec2(rand(), rand());
        vec3 dir = cameraRay((vec2(pixel) + jitter) / vec2(size));

        Primary pr;
        vec3 c = radiance(pc.camPos.xyz, dir, pr);
        // 極端に明るい点（ホタル）を抑える
        float lum = dot(c, vec3(0.2126, 0.7152, 0.0722));
        if (lum > 50.0) c *= 50.0 / lum;
        demod += c / pr.albedo;
        albedo += pr.albedo;
    }
    demod /= float(spp);
    albedo /= float(spp);

    imageStore(uNoisy, pixel, vec4(demod, 1.0));
    // 位置・向きは画素の中心で求める（輪郭で毎フレーム手前・奥が入れ替わらないように）
    vec3 gpos, gnormal;
    float gdist, gmirror;
    primaryGeometry(pc.camPos.xyz, cameraRay((vec2(pixel) + 0.5) / vec2(size)), gpos, gnormal, gdist, gmirror);
    imageStore(uGPos, pixel, vec4(gpos, gdist));
    imageStore(uGNormal, pixel, vec4(gnormal, gmirror));
    imageStore(uGAlbedo, pixel, vec4(albedo, 1.0));
}

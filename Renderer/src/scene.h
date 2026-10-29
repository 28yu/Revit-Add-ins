#pragma once
// 描画するモデルのデータ。座標はメートル、Z が上（Revit と同じ）。
#include <cstdint>
#include <string>
#include <vector>

namespace t28 {

struct Vec3 {
    float x = 0, y = 0, z = 0;
};

struct Vertex {
    float px, py, pz, u;  // 位置と、面の上の位置（横。メートル）
    float nx, ny, nz, v;  // 法線（面の向き）と、面の上の位置（縦。メートル）
};

// テクスチャの貼り方。面の上の位置 (u, v)（メートル）→ 画像上の位置（1 = 画像 1 枚分）:
//   画像の横 = xform[0] * u + xform[1] * v + xform[2]、画像の縦 = xform[3] * u + xform[4] * v + xform[5]
// （Revit の外観の「サイズ」「オフセット」「回転」をまとめたもの）
struct TextureMap {
    int32_t texture = -1;  // テクスチャ画像の番号（-1 = 無し）
    float xform[6] = {1, 0, 0, 0, 1, 0};
};

enum class BumpType : int32_t { Height = 0, Normal = 1 };

struct Material {
    float r = 0.8f, g = 0.8f, b = 0.8f;  // 色（リニア）
    float opacity = 1.0f;                // 1 = 不透明、それ未満はガラスとして扱う
    float roughness = 0.5f;              // 0 = 鏡のようにつるつる、1 = ざらざら
    float metallic = 0.0f;
    float emissive = 0.0f;               // 自己発光の強さ
    float f0 = 0.04f;                    // 正面から見たときの反射率（Revit の外観の「反射率」）
    TextureMap color;                    // 色の画像（外観の「イメージ」）
    TextureMap bump;                     // 凹凸の画像（外観の「レリーフ パターン」）
    BumpType bumpType = BumpType::Height;
    float bumpDepth = 0.0f;              // 高さの画像の白と黒の差（メートル）。法線の画像では強さ（1 = そのまま）
    std::string name;
};

// テクスチャ画像。1 画素 4 バイト。色の画像は sRGB、凹凸の画像はそのままの値（linear = true）
struct Texture {
    uint32_t width = 0, height = 0;
    bool linear = false;
    std::vector<uint8_t> rgba;
};

// 照明器具の光源。明るさは「光度（カンデラ）÷ 20000」（太陽 5 = 10 万ルクスに合わせた単位）
enum class LightType : int32_t { Omni = 0, Hemisphere = 1, Spot = 2 };
struct Light {
    Vec3 position;
    float radius = 0.05f;      // 光る部分の大きさ（影のぼけ方）
    Vec3 intensity;            // 色 × 明るさ
    LightType type = LightType::Omni;
    Vec3 direction{0, 0, -1};  // 半球・スポットの向き
    float cosInner = 1.0f, cosOuter = 0.0f;  // スポット: 明るさが落ち始める角度・届く角度（cos）
    int64_t element = -1;      // この光源を持つ Revit の要素の ID（ライブ同期の差分で入れ替える。-1 = 不明）
};

struct Camera {
    Vec3 eye{0, -5, 1.6f};
    Vec3 forward{0, 1, 0};
    Vec3 up{0, 0, 1};
    float fovY = 50.0f;  // 縦の画角（度）
};

// 3D ビューの「レンダリング設定」（描画ウィンドウを開いたときの最初の見た目にする。ウィンドウで変えればそちらが優先）
struct RevitLook {
    bool sun = true;              // 照明のスキームで太陽を使う
    bool artificial = true;       // 照明のスキームで人工照明（照明器具）を使う
    int background = 0;           // 0 = 空, 1 = 色（グラデーションも）, 2 = 透明
    uint32_t backgroundColor = 0xFFFFFF;  // 0xBBGGRR
    uint32_t groundColor = 0;     // 0xFFBBGGRR（グラデーションの地面の色。a が 0 なら指定なし）
    float cloudAmount = 0.5f;     // 雲の量（0 = 雲なし 〜 1 = 曇り。0.5 がこれまでの見た目）
    bool operator==(const RevitLook& o) const {
        return sun == o.sun && artificial == o.artificial && background == o.background &&
               backgroundColor == o.backgroundColor && groundColor == o.groundColor && cloudAmount == o.cloudAmount;
    }
    bool operator!=(const RevitLook& o) const { return !(*this == o); }
};

// 形（メッシュ）: 三角形の並びのうち、連続した一続き。座標は「置き場所」から見た座標（インスタンスの座標系）
struct MeshRange {
    uint32_t firstTri = 0, triCount = 0;
    uint32_t opaqueCount = 0;  // 不透明な三角形の数（sortGlassLast がメッシュの中で不透明→ガラスの順に並べる）
};
// 置き場所（インスタンス）: どの形を、どの位置・向きに置くか。
// m は 3 行 4 列（行ごと）: 世界の座標 = m の左 3 列 × 形の座標 + m の右 1 列
struct Instance {
    uint32_t mesh = 0;
    float m[12] = {1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0};
    int64_t element = -1;  // この置き場所を描いている Revit の要素の ID（ライブ同期の差分で入れ替える。-1 = 不明）
};

struct Scene {
    // 形のデータ。同じ形（同じファミリなど）は 1 回だけ持ち、instances で何度も置く（GPU のメモリ・ファイルの大きさを抑える）。
    // meshes が空のときは「全部の三角形で 1 つの形を、そのままの位置に 1 回置く」とみなす（版 5 までのファイル）
    std::vector<Vertex> vertices;
    std::vector<uint32_t> indices;       // 3つで1つの三角形
    std::vector<uint32_t> triMaterials;  // 三角形ごとのマテリアル番号
    std::vector<MeshRange> meshes;
    std::vector<Instance> instances;
    std::vector<Material> materials;
    std::vector<Texture> textures;
    std::vector<Light> lights;
    Camera camera;
    Vec3 sunDir{0.5f, -0.6f, 0.62f};     // 太陽のある方向
    float sunIntensity = 5.0f;
    float skyIntensity = 1.0f;
    // 時間帯を変えるとき（太陽の通り道を求める）に使う敷地の情報。Revit の「場所」と「太陽の設定」から
    float latitude = 35.7f;    // 緯度（度。北が正）
    float trueNorth = 0.0f;    // 真北の向き（ラジアン。Revit のプロジェクトの位置の角度）
    float dayOfYear = 80.0f;   // 1 月 1 日 = 1（既定は春分のころ）
    RevitLook revitLook;
    uint32_t syncSeq = 0;      // ライブ同期の通し番号（このシーンがどこまでの変更を含むか。版 7 から）

    // 形ごとに、不透明な三角形を前、ガラスの三角形を後ろに並べ替える（MeshRange::opaqueCount を決める）。
    // 戻り値: 不透明な三角形の数（全部の形の合計。記録用）
    uint32_t sortGlassLast();
    // meshes / instances が空なら「全部で 1 つの形・そのままの位置に 1 回」を入れる
    void ensureInstances();
    // useLimit 回以下しか置かれていない形を、置き場所の位置に動かして 1 つの形（そのままの位置に 1 回）にまとめる。
    // 1 回しか使わない形（壁・床など）を 1 つずつ別に持つと、加速構造の数が増えて遅くなるため。
    // useLimit を大きくすると全部まとめる（インスタンス化しない。比較用）
    void mergeRarelyUsedMeshes(uint32_t useLimit = 1);

    size_t triangleCount() const { return indices.size() / 3; }
    // 置き場所の分も数えた三角形の数（画面に描かれる数）
    size_t instancedTriangleCount() const;
};

// 3 行 4 列の行列（Instance::m と同じ形）の逆行列。逆にできなければ false（out は単位行列）
bool invertAffine(const float m[12], float out[12]);

// Revit 側（Tools28）が書き出すシーンファイル（.t28s）を読む。
// offset: ファイルの途中から読む（ビューアの exe に埋め込んだシーンを読むとき）
// allowEmpty: 三角形が無くてもよい（ライブ同期の差分で、消すだけのとき）
bool loadScene(const std::string& path, Scene& scene, std::string& error, long offset = 0, bool allowEmpty = false);
// ライブ同期の差分（Revit 側 SceneWriter.WriteDelta）を読む。removed = 消す要素の ID、seq = 通し番号
bool loadSceneDelta(const std::string& path, Scene& delta, std::vector<int64_t>& removed, uint32_t& seq, std::string& error);
// 差分を足す: removed の要素の置き場所・光源を消してから、delta の形・置き場所・光源・マテリアル・テクスチャを足す。
// 使われなくなった形が多くなったら詰め直す。scene は prepareScene する前のもの（要素 ID が残っているもの）であること
void applySceneDelta(Scene& scene, const Scene& delta, const std::vector<int64_t>& removed);
// シーンを同じ形式で書き出す（ビューアの exe に埋め込むときに使う）
void saveScene(const Scene& scene, std::vector<uint8_t>& out);
// 時刻（時）の太陽の向き。敷地の緯度・真北・日付から求める（時刻は太陽の動きに合わせた時刻で、正午に真南）
Vec3 sunDirectionAt(const Scene& scene, float hour);
// 太陽の向きが何時ごろか（時間帯のつまみの最初の位置に使う）
float hourOfSun(const Scene& scene, Vec3 dir);
// 動作確認用の部屋（窓・柱・ガラス・鏡）
Scene makeTestScene();
// 性能測定用のモデル（立体駐車場）。preset = small / large / huge。知らない名前なら false
bool makeBenchScene(const std::string& preset, Scene& scene);

}  // namespace t28

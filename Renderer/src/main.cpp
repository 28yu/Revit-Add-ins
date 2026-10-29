// Tools28 リアルタイムレンダリング（描画プログラム本体）
//
// 使い方:
//   Tools28Render --scene model.t28s [--lang JP]          ウィンドウで表示（Revit から起動される）
//   Tools28Render --test-scene                             確認用の部屋をウィンドウで表示
//   Tools28Render --scene model.t28s --output out.png      画像だけ作って終わる
//       [--width 1280 --height 720 --samples 256 --panorama]
// 見た目: --background RRGGBB / --mono RRGGBB / --no-ground / --transparent / --time 時 /
//         --lighting sun|artificial|both / --clouds 0〜1（画像だけ作るとき）
// その他: --log ファイル / --gpu 番号 / --bounces 回数 / --exposure EV / --no-auto-exposure / --render-scale 倍率
//         --software（RT コアを使わない予備の方式で描く。動作確認用）
// 性能測定: --bench-scene small|large|huge（測定用の立体駐車場）/ --bench（画像だけ作るときに時間・メモリ・大きさを測る）/
//         --bench-frames 回数 / --save-scene 保存先.t28s
// ローカルトーンマッピングの強さ: --local-tonemap 0〜1（0 = 使わない。比較用）/ カメラ: --camera x,y,z,前x,前y,前z
// 同じ形を何度も持つ（インスタンス化しない。比較用）: --no-instancing
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <string>
#include <vector>

#include "app.h"
#include "export.h"
#include "log.h"
#include "i18n.h"

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shellapi.h>
#endif

using namespace t28;

static std::string defaultLogPath() {
#ifdef _WIN32
    const char* base = std::getenv("LOCALAPPDATA");
    if (base) {
        std::string dir = std::string(base) + "\\Tools28\\Render";
        CreateDirectoryA((std::string(base) + "\\Tools28").c_str(), nullptr);
        CreateDirectoryA(dir.c_str(), nullptr);
        return dir + "\\render.log";
    }
#endif
    return "";
}

static int run(int argc, char** argv) {
    Options opt;
    bool testScene = false;
    std::string benchScene;  // 性能測定用のモデル（small / large / huge）
    bool cameraOverride = false;
    float cameraValues[6] = {};
    const char* lang = nullptr;
    for (int i = 1; i < argc; ++i) {
        std::string a = argv[i];
        auto next = [&]() -> const char* { return i + 1 < argc ? argv[++i] : ""; };
        if (a == "--scene") opt.scenePath = next();
        else if (a == "--test-scene") testScene = true;
        else if (a == "--bench-scene") benchScene = next();
        else if (a == "--bench") opt.bench = true;
        else if (a == "--no-instancing") opt.instancing = false;
        else if (a == "--local-tonemap") opt.localToneMapping = static_cast<float>(std::atof(next()));
        else if (a == "--camera") {
            // 確認用: カメラの位置と向き「x,y,z,前x,前y,前z」（メートル）
            float c[6] = {0, 0, 0, 0, 1, 0};
            if (std::sscanf(next(), "%f,%f,%f,%f,%f,%f", &c[0], &c[1], &c[2], &c[3], &c[4], &c[5]) == 6) {
                cameraOverride = true;
                std::copy(c, c + 6, cameraValues);
            }
        }
        else if (a == "--bench-frames") opt.benchFrames = std::max(1, std::atoi(next()));
        else if (a == "--save-scene") opt.saveScenePath = next();
        else if (a == "--output") opt.outputPath = next();
        else if (a == "--log") opt.logPath = next();
        else if (a == "--title") opt.title = next();
        else if (a == "--view") opt.viewName = next();
        else if (a == "--live") opt.liveSync = true;
        else if (a == "--lang") lang = next();
        else if (a == "--width") opt.width = std::atoi(next());
        else if (a == "--height") opt.height = std::atoi(next());
        else if (a == "--samples") opt.samples = std::atoi(next());
        else if (a == "--panorama") opt.panorama = true;
        else if (a == "--gpu") opt.gpu = std::atoi(next());
        else if (a == "--bounces") opt.bounces = static_cast<uint32_t>(std::atoi(next()));
        else if (a == "--exposure") opt.exposureEv = static_cast<float>(std::atof(next()));
        else if (a == "--no-auto-exposure") opt.autoExposure = false;
        else if (a == "--render-scale") opt.renderScale = static_cast<float>(std::atof(next()));
        else if (a == "--exit-after") opt.exitAfterFrames = std::atoi(next());
        else if (a == "--export-samples") opt.exportSamples = std::atoi(next());
        else if (a == "--sky-size") opt.skySize = std::max(256, std::min(8192, std::atoi(next())));
        else if (a == "--motion-test") opt.motionTest = std::atoi(next());
        else if (a == "--motion-reference") opt.motionReference = std::atoi(next());
        else if (a == "--software") opt.forceSoftware = true;
        // 見た目の確認用: --background RRGGBB（空・雲の代わりの背景色）/ --mono RRGGBB（物を全部この色に）
        else if (a == "--no-ground") opt.hideGround = true;
        else if (a == "--transparent") opt.transparent = true;
        else if (a == "--time") opt.timeHour = static_cast<float>(std::atof(next()));
        else if (a == "--lighting") opt.lighting = next();
        else if (a == "--clouds") opt.clouds = static_cast<float>(std::atof(next()));
        else if (a == "--background" || a == "--mono") {
            unsigned rgb = static_cast<unsigned>(std::strtoul(next(), nullptr, 16));
            uint32_t bgr = 0xFF000000u | ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF);
            (a == "--background" ? opt.lookBackground : opt.lookMono) = bgr;
        }
    }
    // シーンの指定が無ければ、自分の exe に埋め込まれたシーンを探す（書き出したビューアとして開かれたとき）
    std::string embeddedExe;
    long embeddedOffset = 0;
    bool embedded = opt.scenePath.empty() && !testScene && benchScene.empty() &&
                    findEmbeddedScene(embeddedExe, embeddedOffset, opt.title);
    // ビュー名: 指定が無ければ題名（モデル名 - ビュー名）の後ろの部分（書き出したビューアの exe でも使えるように）
    if (opt.viewName.empty()) {
        size_t sep = opt.title.rfind(" - ");
        opt.viewName = sep == std::string::npos ? opt.title : opt.title.substr(sep + 3);
    }
    // 言語: 指定が無ければ Windows の表示言語（ビューアを受け取った人の PC に合わせる）
    setLanguage(lang ? lang : systemLanguage());
    if (opt.logPath.empty() && opt.outputPath.empty()) opt.logPath = defaultLogPath();
    if (!opt.logPath.empty()) logOpen(opt.logPath);
    installCrashHandler();
    logInfo("Tools28Render start (scene=%s, output=%s, embedded=%d)", opt.scenePath.c_str(), opt.outputPath.c_str(),
            embedded ? 1 : 0);

    Scene scene;
    auto loadStart = std::chrono::steady_clock::now();
    if (!benchScene.empty()) {
        if (!makeBenchScene(benchScene, scene)) fatal("unknown bench scene: %s (small / large / huge)", benchScene.c_str());
        logInfo("Bench scene %s: %zu triangles stored (%zu drawn), %zu vertices, %zu shapes, %zu placements, %zu textures",
                benchScene.c_str(), scene.triangleCount(), scene.instancedTriangleCount(), scene.vertices.size(),
                scene.meshes.size(), scene.instances.size(), scene.textures.size());
    } else if (embedded) {
        std::string error;
        if (!loadScene(embeddedExe, scene, error, embeddedOffset)) fatal("%s\n%s", tr("SceneLoadFailed"), error.c_str());
    } else if (!opt.scenePath.empty() && !testScene) {
        std::string error;
        if (!loadScene(opt.scenePath, scene, error)) fatal("%s\n%s", tr("SceneLoadFailed"), error.c_str());
    } else {
        scene = makeTestScene();
    }
    opt.sceneLoadSeconds = std::chrono::duration<double>(std::chrono::steady_clock::now() - loadStart).count();
    if (cameraOverride) {
        scene.camera.eye = {cameraValues[0], cameraValues[1], cameraValues[2]};
        scene.camera.forward = {cameraValues[3], cameraValues[4], cameraValues[5]};
        scene.camera.up = {0, 0, 1};
    }
    // 確認用: レンダリング設定の照明・雲の量を置き換える
    if (!opt.lighting.empty()) {
        scene.revitLook.sun = opt.lighting != "artificial";
        scene.revitLook.artificial = opt.lighting != "sun";
    }
    if (opt.clouds >= 0) scene.revitLook.cloudAmount = std::min(1.0f, opt.clouds);
    if (opt.width < 16) opt.width = 16;
    if (opt.height < 16) opt.height = 16;
    if (opt.samples < 1) opt.samples = 1;
    if (opt.renderScale != 0.0f && (opt.renderScale < 0.25f || opt.renderScale > 2.0f)) opt.renderScale = 0.0f;

    return opt.outputPath.empty() ? runWindow(opt, scene) : runOffscreen(opt, scene);
}

#ifdef _WIN32
// Windows では引数を UTF-16 で受け取り、日本語のファイル名も扱えるよう UTF-8 に変換する
int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR, int) {
    int argc = 0;
    wchar_t** wargv = CommandLineToArgvW(GetCommandLineW(), &argc);
    std::vector<std::string> args(argc);
    std::vector<char*> argv(argc);
    for (int i = 0; i < argc; ++i) {
        int len = WideCharToMultiByte(CP_UTF8, 0, wargv[i], -1, nullptr, 0, nullptr, nullptr);
        args[i].resize(len > 0 ? len - 1 : 0);
        WideCharToMultiByte(CP_UTF8, 0, wargv[i], -1, &args[i][0], len, nullptr, nullptr);
        argv[i] = &args[i][0];
    }
    LocalFree(wargv);
    return run(argc, argv.data());
}
#else
int main(int argc, char** argv) { return run(argc, argv); }
#endif

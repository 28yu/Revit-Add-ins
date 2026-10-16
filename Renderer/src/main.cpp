// Tools28 リアルタイムレンダリング（描画プログラム本体）
//
// 使い方:
//   Tools28Render --scene model.t28s [--lang JP]          ウィンドウで表示（Revit から起動される）
//   Tools28Render --test-scene                             確認用の部屋をウィンドウで表示
//   Tools28Render --scene model.t28s --output out.png      画像だけ作って終わる
//       [--width 1280 --height 720 --samples 256 --panorama]
// その他: --log ファイル / --gpu 番号 / --bounces 回数 / --exposure EV / --no-auto-exposure / --render-scale 倍率
//         --software（RT コアを使わない予備の方式で描く。動作確認用）
#include <algorithm>
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
    const char* lang = nullptr;
    for (int i = 1; i < argc; ++i) {
        std::string a = argv[i];
        auto next = [&]() -> const char* { return i + 1 < argc ? argv[++i] : ""; };
        if (a == "--scene") opt.scenePath = next();
        else if (a == "--test-scene") testScene = true;
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
    }
    // シーンの指定が無ければ、自分の exe に埋め込まれたシーンを探す（書き出したビューアとして開かれたとき）
    std::string embeddedExe;
    long embeddedOffset = 0;
    bool embedded = opt.scenePath.empty() && !testScene &&
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
    logInfo("Tools28Render start (scene=%s, output=%s, embedded=%d)", opt.scenePath.c_str(), opt.outputPath.c_str(),
            embedded ? 1 : 0);

    Scene scene;
    if (embedded) {
        std::string error;
        if (!loadScene(embeddedExe, scene, error, embeddedOffset)) fatal("%s\n%s", tr("SceneLoadFailed"), error.c_str());
    } else if (!opt.scenePath.empty() && !testScene) {
        std::string error;
        if (!loadScene(opt.scenePath, scene, error)) fatal("%s\n%s", tr("SceneLoadFailed"), error.c_str());
    } else {
        scene = makeTestScene();
    }
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

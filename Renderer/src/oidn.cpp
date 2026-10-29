#include "oidn.h"

#include <algorithm>
#include <cstddef>
#include <thread>
#include <type_traits>

#include "log.h"

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#else
#include <dlfcn.h>
#include <unistd.h>
#endif

namespace t28 {

// OIDN の C の関数（oidn.h と同じ形。ヘッダーを入れずに、使う分だけ自分で書く）
namespace {
using OIDNDevice = void*;
using OIDNFilter = void*;
const int kDeviceTypeCpu = 1;  // OIDN_DEVICE_TYPE_CPU
const int kFormatFloat3 = 3;   // OIDN_FORMAT_FLOAT3
const int kQualityHigh = 6;    // OIDN_QUALITY_HIGH
const int kErrorNone = 0;      // OIDN_ERROR_NONE
}  // namespace

struct OidnDenoiser::Api {
    OIDNDevice (*newDevice)(int);
    void (*setDeviceInt)(OIDNDevice, const char*, int);
    void (*commitDevice)(OIDNDevice);
    int (*getDeviceError)(OIDNDevice, const char**);
    void (*releaseDevice)(OIDNDevice);
    OIDNFilter (*newFilter)(OIDNDevice, const char*);
    void (*setSharedFilterImage)(OIDNFilter, const char*, void*, int, size_t, size_t, size_t, size_t, size_t);
    void (*setFilterBool)(OIDNFilter, const char*, bool);
    void (*setFilterInt)(OIDNFilter, const char*, int);
    void (*commitFilter)(OIDNFilter);
    void (*executeFilter)(OIDNFilter);
    void (*releaseFilter)(OIDNFilter);
};

namespace {

// exe のあるフォルダ（最後に区切り文字を付ける）
std::string exeDirectory() {
#ifdef _WIN32
    std::wstring buf(32768, L'\0');
    DWORD n = GetModuleFileNameW(nullptr, &buf[0], static_cast<DWORD>(buf.size()));
    buf.resize(n);
    size_t slash = buf.find_last_of(L"\\/");
    std::wstring dir = slash == std::wstring::npos ? std::wstring() : buf.substr(0, slash + 1);
    int len = WideCharToMultiByte(CP_UTF8, 0, dir.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string s(len > 0 ? len - 1 : 0, '\0');
    if (len > 1) WideCharToMultiByte(CP_UTF8, 0, dir.c_str(), -1, &s[0], len, nullptr, nullptr);
    return s;
#else
    char buf[4096];
    ssize_t n = readlink("/proc/self/exe", buf, sizeof(buf) - 1);
    if (n <= 0) return "";
    buf[n] = '\0';
    std::string p(buf);
    size_t slash = p.find_last_of('/');
    return slash == std::string::npos ? std::string() : p.substr(0, slash + 1);
#endif
}

void* openLibrary(const std::string& path) {
#ifdef _WIN32
    int n = MultiByteToWideChar(CP_UTF8, 0, path.c_str(), -1, nullptr, 0);
    std::wstring w(n > 0 ? n - 1 : 0, L'\0');
    if (n > 1) MultiByteToWideChar(CP_UTF8, 0, path.c_str(), -1, &w[0], n);
    // 依存する DLL（OpenImageDenoise_core.dll・tbb12.dll）も同じフォルダから探す
    return reinterpret_cast<void*>(LoadLibraryExW(w.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
#else
    return dlopen(path.c_str(), RTLD_NOW | RTLD_LOCAL);
#endif
}

void* findSymbol(void* lib, const char* name) {
#ifdef _WIN32
    return reinterpret_cast<void*>(GetProcAddress(static_cast<HMODULE>(lib), name));
#else
    return dlsym(lib, name);
#endif
}

void closeLibrary(void* lib) {
#ifdef _WIN32
    FreeLibrary(static_cast<HMODULE>(lib));
#else
    dlclose(lib);
#endif
}

}  // namespace

OidnDenoiser::~OidnDenoiser() {
    if (device_ && api_) api_->releaseDevice(device_);
    delete api_;
    // DLL は閉じない（TBB のスレッドが残っていることがあり、閉じると終了時に落ちることがあるため。プログラムの終了で解放される）
}

bool OidnDenoiser::load() {
    if (device_) return true;
#ifdef _WIN32
    // tbb12.dll は CPU 用の部品（OpenImageDenoise_device_cpu.dll）が使う。部品は OIDN が後から読み込むが、そのとき
    // oidn フォルダは探されない（exe のフォルダと Windows のフォルダだけ）ので、先に場所を指定して読んでおく
    openLibrary(exeDirectory() + "oidn\\tbb12.dll");
    const std::string path = exeDirectory() + "oidn\\OpenImageDenoise.dll";
#else
    // Linux（動作確認用）: 配布物の RPATH（$ORIGIN/../lib）では見つからないので、依存するものを先に読んでおく
    const std::string dir = exeDirectory() + "liboidn/";
    for (const char* dep : {"libtbb.so.12", "libOpenImageDenoise_core.so.2.3.3"}) dlopen((dir + dep).c_str(), RTLD_NOW | RTLD_GLOBAL);
    const std::string path = dir + "libOpenImageDenoise.so.2";
#endif
    lib_ = openLibrary(path);
    if (!lib_) {
#ifdef _WIN32
        logInfo("Open Image Denoise not found (%s, error %lu): using the built-in denoiser only", path.c_str(), GetLastError());
#else
        logInfo("Open Image Denoise not found (%s: %s): using the built-in denoiser only", path.c_str(), dlerror());
#endif
        return false;
    }
    Api api{};
    bool ok = true;
    auto get = [&](auto& fn, const char* name) {
        fn = reinterpret_cast<std::remove_reference_t<decltype(fn)>>(findSymbol(lib_, name));
        if (!fn) ok = false;
    };
    get(api.newDevice, "oidnNewDevice");
    get(api.setDeviceInt, "oidnSetDeviceInt");
    get(api.commitDevice, "oidnCommitDevice");
    get(api.getDeviceError, "oidnGetDeviceError");
    get(api.releaseDevice, "oidnReleaseDevice");
    get(api.newFilter, "oidnNewFilter");
    get(api.setSharedFilterImage, "oidnSetSharedFilterImage");
    get(api.setFilterBool, "oidnSetFilterBool");
    get(api.setFilterInt, "oidnSetFilterInt");
    get(api.commitFilter, "oidnCommitFilter");
    get(api.executeFilter, "oidnExecuteFilter");
    get(api.releaseFilter, "oidnReleaseFilter");
    if (!ok) {
        logError("Open Image Denoise: missing functions in %s", path.c_str());
        closeLibrary(lib_);
        lib_ = nullptr;
        return false;
    }
    OIDNDevice dev = api.newDevice(kDeviceTypeCpu);
    if (!dev) {
        const char* message = nullptr;
        api.getDeviceError(nullptr, &message);  // 装置を作れなかったときの理由（スレッドごとの最後のエラー）
        logError("Open Image Denoise: could not create the CPU device (%s)", message ? message : "unknown");
        return false;
    }
    // Revit や描画ウィンドウが固まらないよう、CPU のコアを 1 つ空けておく
    unsigned cores = std::thread::hardware_concurrency();
    if (cores > 2) api.setDeviceInt(dev, "numThreads", static_cast<int>(cores - 1));
    api.commitDevice(dev);
    const char* message = nullptr;
    if (api.getDeviceError(dev, &message) != kErrorNone) {
        logError("Open Image Denoise: %s", message ? message : "device error");
        api.releaseDevice(dev);
        return false;
    }
    api_ = new Api(api);
    device_ = dev;
    logInfo("Open Image Denoise loaded (%s, %u threads)", path.c_str(), cores > 2 ? cores - 1 : cores);
    return true;
}

bool OidnDenoiser::denoise(const float* color, const float* albedo, const float* normal, float* output, int width, int height) {
    if (!device_) return false;
    const Api& a = *api_;
    OIDNFilter filter = a.newFilter(device_, "RT");  // レイトレーシングの画像用
    if (!filter) return false;
    const size_t w = static_cast<size_t>(width), h = static_cast<size_t>(height);
    a.setSharedFilterImage(filter, "color", const_cast<float*>(color), kFormatFloat3, w, h, 0, 0, 0);
    a.setSharedFilterImage(filter, "albedo", const_cast<float*>(albedo), kFormatFloat3, w, h, 0, 0, 0);
    a.setSharedFilterImage(filter, "normal", const_cast<float*>(normal), kFormatFloat3, w, h, 0, 0, 0);
    a.setSharedFilterImage(filter, "output", output, kFormatFloat3, w, h, 0, 0, 0);
    a.setFilterBool(filter, "hdr", true);       // 明るさの範囲が広い画像（1 を超える）
    a.setFilterBool(filter, "cleanAux", true);  // 面の色・向きにはノイズが無い（細部をより残せる）
    a.setFilterInt(filter, "quality", kQualityHigh);
    a.commitFilter(filter);
    a.executeFilter(filter);
    const char* message = nullptr;
    bool ok = a.getDeviceError(device_, &message) == kErrorNone;
    if (!ok) logError("Open Image Denoise: %s", message ? message : "error");
    a.releaseFilter(filter);
    return ok;
}

}  // namespace t28

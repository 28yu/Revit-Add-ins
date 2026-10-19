#include "log.h"

#include <cstdarg>
#include <cstdio>
#include <cstdlib>
#include <ctime>
#include <exception>
#include <mutex>
#include <new>

#include "i18n.h"

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#endif

namespace t28 {

static FILE* g_logFile = nullptr;
static std::mutex g_logMutex;

FILE* openFile(const std::string& utf8Path, const char* mode) {
#ifdef _WIN32
    wchar_t wpath[1024], wmode[16];
    MultiByteToWideChar(CP_UTF8, 0, utf8Path.c_str(), -1, wpath, 1024);
    MultiByteToWideChar(CP_UTF8, 0, mode, -1, wmode, 16);
    return _wfopen(wpath, wmode);
#else
    return fopen(utf8Path.c_str(), mode);
#endif
}

void logOpen(const std::string& path) {
    std::lock_guard<std::mutex> lock(g_logMutex);
    if (g_logFile) fclose(g_logFile);
    // 前回のログは render.prev.log として 1 つだけ残す（落ちたあとに開き直しても、落ちたときの記録を調べられるように）
    std::string prev = path;
    size_t dot = prev.rfind('.');
    prev = (dot == std::string::npos ? prev : prev.substr(0, dot)) + ".prev.log";
#ifdef _WIN32
    wchar_t wfrom[1024], wto[1024];
    MultiByteToWideChar(CP_UTF8, 0, path.c_str(), -1, wfrom, 1024);
    MultiByteToWideChar(CP_UTF8, 0, prev.c_str(), -1, wto, 1024);
    MoveFileExW(wfrom, wto, MOVEFILE_REPLACE_EXISTING);
#else
    std::rename(path.c_str(), prev.c_str());
#endif
    g_logFile = openFile(path, "w");
}

static void writeLine(const char* level, const char* fmt, va_list args) {
    char message[2048];
    vsnprintf(message, sizeof(message), fmt, args);
    time_t now = time(nullptr);
    char stamp[32];
    strftime(stamp, sizeof(stamp), "%H:%M:%S", localtime(&now));

    std::lock_guard<std::mutex> lock(g_logMutex);
    fprintf(stderr, "[%s] %s %s\n", stamp, level, message);
    if (g_logFile) {
        fprintf(g_logFile, "[%s] %s %s\n", stamp, level, message);
        fflush(g_logFile);
    }
}

void logInfo(const char* fmt, ...) {
    va_list args;
    va_start(args, fmt);
    writeLine("INFO ", fmt, args);
    va_end(args);
}

void logError(const char* fmt, ...) {
    va_list args;
    va_start(args, fmt);
    writeLine("ERROR", fmt, args);
    va_end(args);
}

#ifdef _WIN32
// 想定外の異常終了（メモリの読み書き違反など）を記録して、何も出さずに消えないようにする
static LONG WINAPI onCrash(EXCEPTION_POINTERS* info) {
    unsigned long code = info && info->ExceptionRecord ? info->ExceptionRecord->ExceptionCode : 0;
    void* address = info && info->ExceptionRecord ? info->ExceptionRecord->ExceptionAddress : nullptr;
    char module[MAX_PATH] = "?";
    HMODULE mod = nullptr;
    if (address && GetModuleHandleExA(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                      static_cast<LPCSTR>(address), &mod))
        GetModuleFileNameA(mod, module, MAX_PATH);
    logError("CRASH: exception 0x%08lX at %p (%s)", code, address, module);
    const char* key = code == EXCEPTION_STACK_OVERFLOW ? "CrashedStack" : "Crashed";
    wchar_t wide[2048];
    MultiByteToWideChar(CP_UTF8, 0, tr(key), -1, wide, 2048);
    MessageBoxW(nullptr, wide, L"Tools28 Render", MB_OK | MB_ICONERROR);
    if (g_logFile) fclose(g_logFile);
    return EXCEPTION_EXECUTE_HANDLER;
}
#endif

static void onTerminate() {
    // メモリが足りない（std::bad_alloc）などの例外が受け止められずに終わるとき
    const char* what = "unknown";
    bool outOfMemory = false;
    if (std::exception_ptr e = std::current_exception()) {
        try { std::rethrow_exception(e); }
        catch (const std::bad_alloc&) { what = "std::bad_alloc"; outOfMemory = true; }
        catch (const std::exception& ex) { what = ex.what(); }
        catch (...) {}
    }
    fatal("%s\n(%s)", tr(outOfMemory ? "OutOfMemory" : "Crashed"), what);
}

void installCrashHandler() {
#ifdef _WIN32
    SetUnhandledExceptionFilter(onCrash);
    // スタックが溢れたときも、記録を書くための余裕を残しておく
    ULONG guarantee = 64 * 1024;
    SetThreadStackGuarantee(&guarantee);
#endif
    std::set_terminate(onTerminate);
}

void fatal(const char* fmt, ...) {
    char message[2048];
    va_list args;
    va_start(args, fmt);
    vsnprintf(message, sizeof(message), fmt, args);
    va_end(args);
    logError("%s", message);
#ifdef _WIN32
    // 文字列は UTF-8 なので、画面表示用に UTF-16 へ変換する
    wchar_t wide[2048];
    MultiByteToWideChar(CP_UTF8, 0, message, -1, wide, 2048);
    MessageBoxW(nullptr, wide, L"Tools28 Render", MB_OK | MB_ICONERROR);
#endif
    if (g_logFile) fclose(g_logFile);
    std::exit(1);
}

}  // namespace t28

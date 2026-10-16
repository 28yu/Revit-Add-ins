#include "log.h"

#include <cstdarg>
#include <cstdio>
#include <cstdlib>
#include <ctime>
#include <mutex>

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

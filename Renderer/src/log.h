#pragma once
// 動作記録（ログ）。不具合のときにユーザーから送ってもらい、原因を調べるために使う。
#include <cstdio>
#include <string>

namespace t28 {

// UTF-8 のファイル名でファイルを開く（Windows でも日本語のフォルダ名を扱えるようにする）
FILE* openFile(const std::string& utf8Path, const char* mode);

void logOpen(const std::string& path);
void logInfo(const char* fmt, ...);
void logError(const char* fmt, ...);
// 致命的なエラー: 記録して、Windows ではメッセージ画面も出し、終了する
[[noreturn]] void fatal(const char* fmt, ...);

}  // namespace t28

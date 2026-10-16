#pragma once
// 画面に出す文字（日本語・英語・中国語）。言語は Revit 側（Tools28 の言語設定）から --lang で受け取る。

namespace t28 {

enum class Lang { JP, US, CN };

void setLanguage(const char* code);  // "JP" / "US" / "CN"
const char* tr(const char* key);

}  // namespace t28

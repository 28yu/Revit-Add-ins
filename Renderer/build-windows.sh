#!/bin/bash
# Windows 用の Tools28Render.exe を Linux 上で作り、prebuilt/win-x64 に置く。
# 必要なもの: cmake, mingw-w64, glslc（apt install cmake mingw-w64 glslc）
set -euo pipefail
cd "$(dirname "$0")"
BUILD_DIR="${1:-build-win}"
cmake -S . -B "$BUILD_DIR" -DCMAKE_TOOLCHAIN_FILE=cmake/mingw-w64.cmake -DCMAKE_BUILD_TYPE=Release
cmake --build "$BUILD_DIR" -j"$(nproc)"
mkdir -p prebuilt/win-x64
x86_64-w64-mingw32-strip -o prebuilt/win-x64/Tools28Render.exe "$BUILD_DIR/Tools28Render.exe"
ls -la prebuilt/win-x64/Tools28Render.exe

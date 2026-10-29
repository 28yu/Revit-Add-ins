#pragma once
// ビューアの exe に入れるシーンを小さくする（「詰めたシーン」。.t28s とは別の形式で、ビューアの中だけで使う）。
//
// 形のデータは、そのままだと 1 頂点 32 バイト・1 三角形 16 バイトで、大きなモデルでは 100MB を超える。
//   - 頂点: 位置・面の上の位置は「前の頂点との差」（数値のビットの差なので元に戻せる）、面の向きは 16 ビット×2（GPU と同じ精度）
//   - 三角形: 頂点番号は「前の番号との差」、マテリアル番号はそのまま
//   - どれも 4 バイトの数を「1 バイト目だけ・2 バイト目だけ…」に並べ替えてから、全体を deflate（zip と同じ方式）で圧縮する
//     （似た値が続くので、並べ替えると圧縮しやすくなる）
//   - 色の画像（透明な部分が無いもの）は JPEG にする
// 形式: "T28P" | uint32 版 | uint64 元の大きさ | 圧縮したデータ（中身は scene_pack.cpp の packScene の順）
#include <cstdint>
#include <string>
#include <vector>

#include "scene.h"

namespace t28 {

// scene を詰めて out の後ろに足す。camera は最初の視点（ビューアを開いたときの視点）
void packScene(const Scene& scene, const Camera& camera, std::vector<uint8_t>& out);
// 詰めたシーンを読む（data は "T28P" から始まる）
bool unpackScene(const uint8_t* data, size_t size, Scene& scene, std::string& error);
// data が詰めたシーンか
bool isPackedScene(const uint8_t* data, size_t size);

}  // namespace t28

#!/usr/bin/env bash
# 描画プログラムの性能測定（段階0: 改善の前後を数値で比べるための土台）。
#
# 使い方:
#   tools/bench/run.sh [small|large|huge ...]      速さ・メモリ・ファイルの大きさを測る（既定: small large）
#   tools/bench/run.sh --quality                    画質を測る（small を正解の画像と比べる。明るさは決め打ち）
# 環境変数: OUT（結果を置くフォルダ。既定は build-linux/bench）、FRAMES（測るフレーム数。既定 8）、
#           SIZE（画像の大きさ。既定 640x360）、EXTRA（描画プログラムに足す引数）
#
# GPU の無い Linux では Mesa の lavapipe（CPU で Vulkan を再現）で動く。速さの絶対値は実機と大きく違うので、
# 「同じ環境で、変更の前と後を比べる」ためにだけ使う（GPU メモリ・ファイルの大きさ・画質は実機とほぼ同じ傾向になる）。
set -euo pipefail
cd "$(dirname "$0")/../.."
BIN=build-linux/Tools28Render
OUT=${OUT:-build-linux/bench}
FRAMES=${FRAMES:-8}
SIZE=${SIZE:-640x360}
W=${SIZE%x*}
H=${SIZE#*x}
EXTRA=${EXTRA:-}
mkdir -p "$OUT"
if [ ! -x "$BIN" ]; then
  cmake -S . -B build-linux -DT28_SYSTEM_GLFW=ON -DCMAKE_BUILD_TYPE=Release >/dev/null
fi
cmake --build build-linux -j"$(nproc)" >/dev/null
if [ -z "${VK_ICD_FILENAMES:-}" ] && [ -f /usr/share/vulkan/icd.d/lvp_icd.json ]; then
  export VK_ICD_FILENAMES=/usr/share/vulkan/icd.d/lvp_icd.json
fi

bench() {  # $1 = モデル, 残り = 足す引数
  local preset=$1; shift
  "$BIN" --bench-scene "$preset" --bench --bench-frames "$FRAMES" --output "$OUT/$preset.png" \
    --width "$W" --height "$H" --sky-size 1024 --log "$OUT/$preset.log" "$@" $EXTRA >/dev/null 2>&1 || true
  grep -o 'BENCH {.*}' "$OUT/$preset.log" | sed 's/^BENCH //' || echo '{"error":"see log"}'
}

if [ "${1:-}" = "--quality" ]; then
  # 明るさは決め打ちにして比べる（自動露出・ローカルトーンマッピングは絵ごとに少し変わり、ノイズの比較を邪魔するため）
  FIXED="--no-auto-exposure --exposure 2 --local-tonemap 0"
  # 正解: 512 回足し合わせた画像（Open Image Denoise を使わない。一度作れば使い回す）
  if [ ! -f "$OUT/small_reference.png" ]; then
    "$BIN" --bench-scene small --output "$OUT/small_reference.png" --width "$W" --height "$H" --sky-size 1024 \
      --samples 512 --final-denoise 0 $FIXED --log "$OUT/small_reference.log" $EXTRA >/dev/null 2>&1
  fi
  echo "frames=$FRAMES $(bench small $FIXED)"
  python3 tools/bench/compare.py "$OUT/small.png" "$OUT/small_reference.png"
  exit 0
fi

presets=("$@")
[ ${#presets[@]} -eq 0 ] && presets=(small large)
printf '| モデル | 三角形 | シーン(MB) | ビューア(MB) | 準備(s) | GPU最大(MB) | 止まっている(ms) | 動いている(ms) | 最大(ms) | PNG(KB) | 白とび(%%) |\n'
printf '|--|--|--|--|--|--|--|--|--|--|--|\n'
for p in "${presets[@]}"; do
  j=$(bench "$p")
  echo "$j" > "$OUT/$p.json"
  python3 - "$p" "$j" <<'PY'
import json, sys
p, j = sys.argv[1], json.loads(sys.argv[2])
if "error" in j:
    print(f"| {p} | 失敗（ログを確認） |"); sys.exit()
print(f"| {p} | {j['triangles']:,} | {j['scene_mb']} | {j.get('viewer_mb', '-')} | {j['prepare_s'] + j['init_s']:.1f} | {j['gpu_peak_mb']:.0f} | "
      f"{j['still_ms']:.0f} | {j['moving_ms']:.0f} | {max(j['still_max_ms'], j['moving_max_ms']):.0f} | {j['png_kb']:.0f} | {j['blown_pct']} |")
PY
done

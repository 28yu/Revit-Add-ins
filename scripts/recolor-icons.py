#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ボタンアイコンの青系の色を、指定した色に合わせて一括変換する。

アンカー色（変換前の代表色）と目標色の関係を HSV で求め、青系ピクセル全体に
同じ関係を適用する。陰影・グラデーション・アンチエイリアスはそのまま残る。

  H: 目標との差だけ回転
  S: 比率で縮小/拡大
  V: 白へ寄せる曲線 V' = 1 - (1-V)*k （明るい色でクランプしないため）

⚠️ この変換は「一度だけ」実行するもの。変換後のアイコンにもう一度かけると
   色がさらにずれる（冪等ではない）。実行前に必ず --dry-run で確認し、
   git の差分で意図どおりか見てからコミットすること。

使い方:
    python3 scripts/recolor-icons.py --dry-run
    python3 scripts/recolor-icons.py
    python3 scripts/recolor-icons.py --from 0066CC --to 67B1E6

適用履歴:
    2026-08-28  全アイコンに #0066CC → #67B1E6 を適用したが、実機で確認したところ
                元の色のほうが良いものが多かったため、下記3機能を除いて変換前に戻した。
                現在 Revit の青が適用されているのは次の3機能のみ:

                  filter_management     #0066CC → #67B1E6
                  dwg_layer_transfer    #0066CC → #67B1E6
                  view_template         #0066CC → #1D7CBF
                    （白抜きの横線があるため、色相 205° は揃えたまま
                      彩度・明度を戻してコントラストを確保）

    ⚠️ 他のアイコンへ広げる場合は、必ず --only で対象を絞り、--dry-run で
       確認してから適用すること。全体に一括適用すると、凡例・多色パレット・
       白抜き文字を持つアイコンで見え方が崩れる。
"""
import argparse
import colorsys
import glob
import os
import sys

try:
    from PIL import Image
except ImportError:
    sys.exit("Pillow が必要です:  pip install Pillow")

# 青とみなす範囲。低彩度のグレーを除外して、構造物のグレー等を壊さない。
HUE_MIN, HUE_MAX = 185.0, 260.0
SAT_MIN = 0.15

# 変更しないファイル
EXCLUDE = {
    # 国旗（国旗の色は変えてはいけない）
    'flag_us_32.png', 'flag_us_16.png',
    'flag_jp_32.png', 'flag_jp_16.png',
    'flag_cn_32.png', 'flag_cn_16.png',
    # 部屋3D色分け（青/緑/橙/紫の4色パレット。青だけ変えると配色が壊れる）
    'room_3d_color.png',
    # 型枠数量算出（ブランド青ではなく構造物のグレー #A2A7B0 / #8094A5）
    'formwork.png',
    # 梁天端/梁下端レベル色分け（青・黄・桃の3色凡例。青だけ淡くすると凡例の統一感が崩れる）
    'beam_top_level.png', 'beam_under_level.png',
}

TARGET_GLOBS = ('Resources/Icons/*.png', 'Docs/icons/features/*.png')


def parse_hex(text):
    text = text.lstrip('#')
    if len(text) != 6:
        raise argparse.ArgumentTypeError(f"6桁の16進数で指定してください: {text}")
    return tuple(int(text[i:i + 2], 16) for i in (0, 2, 4))


def build_transform(src, dst):
    sh, ss, sv = colorsys.rgb_to_hsv(*[c / 255 for c in src])
    dh, ds, dv = colorsys.rgb_to_hsv(*[c / 255 for c in dst])
    if ss == 0 or sv == 1:
        sys.exit("アンカー色には彩度があり、かつ真っ白でない色を指定してください")
    return (dh - sh) * 360.0, ds / ss, (1 - dv) / (1 - sv)


def make_converter(hue_shift, sat_scale, val_k):
    def convert(r, g, b):
        h, s, v = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)
        deg = h * 360.0
        if not (HUE_MIN <= deg <= HUE_MAX) or s < SAT_MIN:
            return None
        nh = ((deg + hue_shift) % 360.0) / 360.0
        ns = min(1.0, s * sat_scale)
        nv = min(1.0, 1.0 - (1.0 - v) * val_k)
        nr, ng, nb = colorsys.hsv_to_rgb(nh, ns, nv)
        return round(nr * 255), round(ng * 255), round(nb * 255)
    return convert


def process(path, convert, dry_run):
    im = Image.open(path).convert('RGBA')
    px = im.load()
    w, h = im.size
    cache = {}
    changed = 0
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if a == 0:
                continue
            key = (r, g, b)
            if key not in cache:
                cache[key] = convert(r, g, b)
            new = cache[key]
            if new is not None:
                if not dry_run:
                    px[x, y] = (new[0], new[1], new[2], a)
                changed += 1
    if changed and not dry_run:
        im.save(path)
    return changed, w * h


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--from', dest='src', type=parse_hex, default=parse_hex('0066CC'),
                        help='変換前の代表色（アンカー）。既定 0066CC')
    parser.add_argument('--to', dest='dst', type=parse_hex, default=parse_hex('67B1E6'),
                        help='目標色。既定 67B1E6（Revit の青）')
    parser.add_argument('--only', nargs='+', metavar='NAME', default=None,
                        help='指定したファイル名だけを対象にする（例: --only manual_16.png ver_16.png）')
    parser.add_argument('--dry-run', action='store_true', help='書き換えずに件数だけ表示する')
    args = parser.parse_args()

    hue_shift, sat_scale, val_k = build_transform(args.src, args.dst)
    print("#{:02X}{:02X}{:02X} -> #{:02X}{:02X}{:02X}".format(*args.src, *args.dst))
    print(f"H{hue_shift:+.1f}deg  S x{sat_scale:.3f}  V'=1-(1-V)x{val_k:.3f}")
    if not args.dry_run:
        print("!! この変換は冪等ではありません。二重適用しないよう注意してください。")
    print()

    convert = make_converter(hue_shift, sat_scale, val_k)
    targets = []
    for pattern in TARGET_GLOBS:
        targets.extend(sorted(glob.glob(pattern)))

    total = 0
    for path in targets:
        name = os.path.basename(path)
        if args.only is not None and name not in args.only:
            continue
        if name in EXCLUDE:
            print(f"{path:42} skip")
            continue
        changed, _ = process(path, convert, args.dry_run)
        if changed:
            total += 1
            print(f"{path:42} {changed:>7} px")

    print(f"\n{'(dry-run) ' if args.dry_run else ''}changed files: {total}")
    return 0


if __name__ == '__main__':
    sys.exit(main())

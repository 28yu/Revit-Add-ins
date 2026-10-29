#!/usr/bin/env python3
"""画像の比べ方（性能測定の画質の指標）。

使い方: compare.py 比べる画像.png 正解の画像.png
  - RMSE / PSNR: 正解（十分に足し合わせた画像）とのずれ。小さい RMSE・大きい PSNR ほどノイズが少ない
  - 白とび: RGB すべてが 250 以上の画素の割合（％）
  - まだら: 正解との差を 8×8 画素でならした値の標準偏差。ノイズ除去が残した大きな「むら」の強さ
結果は JSON の 1 行で出す。
"""
import json
import sys

import numpy as np
from PIL import Image


def load(path):
    return np.asarray(Image.open(path).convert("RGB"), dtype=np.float64)


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        return 2
    a, ref = load(sys.argv[1]), load(sys.argv[2])
    if a.shape != ref.shape:
        print(f"size differs: {a.shape} vs {ref.shape}", file=sys.stderr)
        return 1
    diff = a - ref
    rmse = float(np.sqrt(np.mean(diff ** 2)))
    psnr = float(20 * np.log10(255.0 / rmse)) if rmse > 0 else 99.0
    blown = float(np.mean(np.all(a >= 250, axis=2)) * 100.0)
    lum = diff @ np.array([0.2126, 0.7152, 0.0722])
    h, w = (lum.shape[0] // 8) * 8, (lum.shape[1] // 8) * 8
    blocks = lum[:h, :w].reshape(h // 8, 8, w // 8, 8).mean(axis=(1, 3))
    print(json.dumps({"rmse": round(rmse, 2), "psnr": round(psnr, 2), "blown_pct": round(blown, 2),
                      "blotch": round(float(blocks.std()), 2)}))
    return 0


if __name__ == "__main__":
    sys.exit(main())

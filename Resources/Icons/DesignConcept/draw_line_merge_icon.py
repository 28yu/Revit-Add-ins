# 直線まとめ（LineMerge）のアイコンを描くスクリプト
# 使い方: python3 draw_line_merge_icon.py <出力フォルダ>
# 上: 同じ直線上に並んだバラバラの線（3本）/ 下: まとめた1本の線。間に下向きの矢印。
# 96px は 288DPI（pHYs）で保存する（無いと WPF でリボンからはみ出す。Docs/DEVLOG.md 参照）
import sys, os
from PIL import Image, ImageDraw
OUT = sys.argv[1]; os.makedirs(OUT, exist_ok=True)
S = 24  # 32単位 → 768px で描いて縮小（アンチエイリアス）
def p(v): return round(v * S)
im = Image.new('RGBA', (32 * S, 32 * S), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
def seg(x1, x2, y, color, w=2.6):
    d.line([(p(x1), p(y)), (p(x2), p(y))], fill=color, width=p(w))
    for x in (x1, x2):   # 端点の丸
        r = 1.9
        d.ellipse([p(x - r), p(y - r), p(x + r), p(y + r)], fill='#ffffff', outline=color, width=p(1.1))
# バラバラの線（薄い青）
seg(3, 9.5, 8, '#67b1e6'); seg(12.5, 18.5, 8, '#67b1e6'); seg(21.5, 29, 8, '#67b1e6')
# 下向きの矢印
d.line([(p(16), p(12.2)), (p(16), p(18.6))], fill='#5c5c5c', width=p(1.8))
d.polygon([(p(12.4), p(17.2)), (p(19.6), p(17.2)), (p(16), p(21.2))], fill='#5c5c5c')
# まとめた1本（濃い青）
seg(3, 29, 25.5, '#0066cc', 3.0)
im.resize((96, 96), Image.LANCZOS).save(f'{OUT}/line_merge.png', dpi=(288, 288))

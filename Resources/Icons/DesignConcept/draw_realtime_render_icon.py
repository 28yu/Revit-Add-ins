# リアルタイムレンダリングのアイコン（96x96 / 288DPI と 16x16）を描くスクリプト
# 使い方: python3 draw_realtime_render_icon.py <出力フォルダ>
# 描き方の約束は draw_v26_icons_96.py と同じ（32単位の座標で大きく描いて縮小する）
import math, sys, os
from PIL import Image, ImageDraw
OUT = sys.argv[1]
os.makedirs(OUT, exist_ok=True)
S = 24

def p(v): return v * S
def poly(d, pts, col): d.polygon([(p(x), p(y)) for x, y in pts], fill=col)

im = Image.new('RGBA', (32 * S, 32 * S), (0, 0, 0, 0))
d = ImageDraw.Draw(im)
# 太陽（右上）と光線
cx, cy, r = 24.5, 7.5, 3.6
d.ellipse([p(cx - r), p(cy - r), p(cx + r), p(cy + r)], fill='#f5a623')
for k in range(8):
    a = math.radians(k * 45)
    x0, y0 = cx + math.cos(a) * 5.0, cy + math.sin(a) * 5.0
    x1, y1 = cx + math.cos(a) * 6.8, cy + math.sin(a) * 6.8
    d.line([(p(x0), p(y0)), (p(x1), p(y1))], fill='#f5a623', width=round(p(1.2)))
# 床に落ちる影
poly(d, [(3, 25.5), (14, 30.5), (22, 27), (11, 22)], '#b9c4cf')
# 立方体（上面・左面・右面で明るさを変える）
poly(d, [(4, 14), (11, 10.5), (18, 14), (11, 17.5)], '#9fc8ef')
poly(d, [(4, 14), (11, 17.5), (11, 27), (4, 23.5)], '#1d6fc2')
poly(d, [(11, 17.5), (18, 14), (18, 23.5), (11, 27)], '#0b4f96')

im.resize((96, 96), Image.LANCZOS).save(f'{OUT}/realtime_render.png', dpi=(288, 288))
im.resize((16, 16), Image.LANCZOS).save(f'{OUT}/realtime_render_16.png')

# 重複要素整理（DuplicateInstance）のアイコンを描くスクリプト
# 使い方: python3 draw_duplicate_instance_icon.py <出力フォルダ>
# 96px は 288DPI（pHYs）で保存する（無いと WPF でリボンからはみ出す。Docs/DEVLOG.md 参照）
import sys, os
from PIL import Image, ImageDraw
OUT = sys.argv[1]; os.makedirs(OUT, exist_ok=True)
S = 24  # 32単位 → 768px で描いて縮小（アンチエイリアス）
def p(v): return round(v * S)
im = Image.new('RGBA', (32 * S, 32 * S), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
def box(x, y, fill, outline):
    d.rounded_rectangle([p(x), p(y), p(x + 15), p(y + 15)], radius=p(1), fill=fill, outline=outline, width=p(1.3))
    d.rounded_rectangle([p(x + 4.2), p(y + 4.2), p(x + 10.8), p(y + 10.8)], radius=p(0.5), fill='#ffffff')
# 奥（残す方）と手前（重なった複製）の同じ形の箱
box(3, 3, '#67b1e6', '#4f93e0')
box(10, 10, '#0066cc', '#005bb5')
# 手前の箱に削除の×
cx, cy, r = 25.2, 25.2, 5.6
d.ellipse([p(cx - r), p(cy - r), p(cx + r), p(cy + r)], fill='#ffffff')
d.ellipse([p(cx - r + 0.9), p(cy - r + 0.9), p(cx + r - 0.9), p(cy + r - 0.9)], fill='#d9534f')
k = 2.4
for a, b in (((cx - k, cy - k), (cx + k, cy + k)), ((cx - k, cy + k), (cx + k, cy - k))):
    d.line([(p(a[0]), p(a[1])), (p(b[0]), p(b[1]))], fill='#ffffff', width=p(1.5))
im.resize((96, 96), Image.LANCZOS).save(f'{OUT}/duplicate_instance.png', dpi=(288, 288))
im.resize((16, 16), Image.LANCZOS).save(f'{OUT}/duplicate_instance_16.png')

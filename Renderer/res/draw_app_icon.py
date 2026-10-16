# 描画プログラム（Tools28Render.exe）と、書き出したビューア exe のアイコンを描くスクリプト。
# リボンの「リアルタイムレンダリング」のアイコン（青い立方体と太陽）を、角の丸い紺の四角に載せ、
# 右下に「28」の札を付ける（28 Tools から書き出した物だと一目で分かるように）。
# 使い方: python3 draw_app_icon.py  → 同じフォルダに Tools28Render.ico（16〜256 画素）と確認用の png を作る
import math, os
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
S = 32  # 32 単位の座標を 32 倍（1024 画素）で描いて縮小する

def p(v): return v * S
def poly(d, pts, col): d.polygon([(p(x), p(y)) for x, y in pts], fill=col)

im = Image.new('RGBA', (32 * S, 32 * S), (0, 0, 0, 0))
d = ImageDraw.Draw(im)
# 背景: 角の丸い紺の四角（上が少し明るい）
for i in range(32 * S):
    t = i / (32 * S)
    c = (int(28 + 20 * (1 - t)), int(44 + 30 * (1 - t)), int(72 + 40 * (1 - t)), 255)
    d.line([(0, i), (32 * S, i)], fill=c)
mask = Image.new('L', im.size, 0)
ImageDraw.Draw(mask).rounded_rectangle([p(0.5), p(0.5), p(31.5), p(31.5)], radius=p(6), fill=255)
bg = Image.new('RGBA', im.size, (0, 0, 0, 0))
bg.paste(im, (0, 0), mask)
im = bg
d = ImageDraw.Draw(im)

# 太陽（右上）と光線
cx, cy, r = 23.0, 8.5, 3.3
d.ellipse([p(cx - r), p(cy - r), p(cx + r), p(cy + r)], fill='#f5a623')
for k in range(8):
    a = math.radians(k * 45)
    x0, y0 = cx + math.cos(a) * 4.6, cy + math.sin(a) * 4.6
    x1, y1 = cx + math.cos(a) * 6.2, cy + math.sin(a) * 6.2
    d.line([(p(x0), p(y0)), (p(x1), p(y1))], fill='#f5a623', width=round(p(1.1)))
# 床に落ちる影と立方体
poly(d, [(4.5, 24.5), (14, 29), (21, 26), (11.5, 21.5)], (255, 255, 255, 60))
poly(d, [(5, 14), (11.5, 10.8), (18, 14), (11.5, 17.2)], '#9fc8ef')
poly(d, [(5, 14), (11.5, 17.2), (11.5, 26), (5, 22.8)], '#2f86dc')
poly(d, [(11.5, 17.2), (18, 14), (18, 22.8), (11.5, 26)], '#1663b0')

# 右下の「28」の札（オレンジ）
d.rounded_rectangle([p(17.0), p(19.5), p(30.5), p(30.0)], radius=p(2.4), fill='#ff7a1a', outline='#ffffff', width=round(p(0.7)))
font = ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', round(p(8.6)))
text = '28'
box = d.textbbox((0, 0), text, font=font)
tw, th = box[2] - box[0], box[3] - box[1]
d.text((p(23.75) - tw / 2 - box[0], p(24.75) - th / 2 - box[1]), text, font=font, fill='#ffffff')

sizes = [16, 24, 32, 48, 64, 128, 256]
im.resize((256, 256), Image.LANCZOS).save(os.path.join(HERE, 'Tools28Render.ico'), sizes=[(s, s) for s in sizes])
im.resize((256, 256), Image.LANCZOS).save(os.path.join(HERE, 'Tools28Render_256.png'))
im.resize((32, 32), Image.LANCZOS).save(os.path.join(HERE, 'Tools28Render_32.png'))
print('ok')

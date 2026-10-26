"""
viewport_match（ビューポート位置合わせ）のアイコンを描く。

viewport_paste.png のシート（灰色の縁・用紙・表題欄）をそのまま使い、中のビューポートを
「元の位置（青の破線）→ そろえた位置（青の実線）」の矢印付きの絵に描き替える。
96px は 288 DPI を入れて保存する（入れないとリボンからはみ出す。Docs/DEVLOG.md 参照）。16px は 96px の縮小。

使い方（リポジトリのルートで）:
  python3 Resources/Icons/DesignConcept/draw_viewport_match_icon.py Resources/Icons/viewport_paste.png Resources/Icons
"""
import sys, math
from PIL import Image, ImageDraw
src, outdir = sys.argv[1], sys.argv[2]
SS = 8
BLUE = (51, 153, 255, 255)
FILL = (232, 232, 232, 255)

def clean_base():
    base = Image.open(src).convert('RGBA')
    d = ImageDraw.Draw(base)
    # 用紙の中（表題欄を除く）を白に戻す
    d.rectangle([13, 27, 61, 72], fill=(255, 255, 255, 255))
    d.rectangle([13, 27, 83, 56], fill=(255, 255, 255, 255))
    return base

def dashed_rect(d, x0, y0, x1, y1, w, dash, gap):
    for (ax, ay, bx, by) in [(x0,y0,x1,y0),(x1,y0,x1,y1),(x1,y1,x0,y1),(x0,y1,x0,y0)]:
        L = math.hypot(bx-ax, by-ay); ux, uy = (bx-ax)/L, (by-ay)/L; p = 0
        while p < L:
            q = min(p+dash, L)
            d.line([(ax+ux*p, ay+uy*p), (ax+ux*q, ay+uy*q)], fill=BLUE, width=w)
            p = q + gap

def arrow(d, x0, y0, x1, y1, w, head):
    ang = math.atan2(y1-y0, x1-x0)
    bx, by = x1 - math.cos(ang)*head*0.8, y1 - math.sin(ang)*head*0.8
    d.line([(x0,y0),(bx,by)], fill=(50,50,50,255), width=w)
    l = (x1 - math.cos(ang-0.5)*head, y1 - math.sin(ang-0.5)*head)
    r = (x1 - math.cos(ang+0.5)*head, y1 - math.sin(ang+0.5)*head)
    d.polygon([(x1,y1), l, r], fill=(50,50,50,255))

def make(name, solid, ghost, arr):
    base = clean_base()
    big = Image.new('RGBA', (96*SS, 96*SS), (0,0,0,0))
    d = ImageDraw.Draw(big)
    k = SS
    w = int(4.5*k)
    gx0,gy0,gx1,gy1 = [v*k for v in ghost]
    dashed_rect(d, gx0, gy0, gx1, gy1, w, 7.5*k, 4.5*k)
    sx0,sy0,sx1,sy1 = [v*k for v in solid]
    d.rectangle([sx0,sy0,sx1,sy1], fill=FILL)
    d.rectangle([sx0,sy0,sx1,sy1], outline=BLUE, width=w)
    if arr:
        arrow(d, *[v*k for v in arr], w=int(3.6*k), head=11*k)
    small = big.resize((96,96), Image.LANCZOS)
    base.alpha_composite(small)
    base.save(f'{outdir}/{name}.png', dpi=(288.0106, 288.0106))
    return base

# 採用した案。96px（288 DPI）と 16px（縮小）を書き出す
final = make('viewport_match', solid=(15,29,45,54), ghost=(29,40,58,65), arr=(50,58,36,45))
final.resize((16,16), Image.LANCZOS).save(f'{outdir}/viewport_match_16.png')

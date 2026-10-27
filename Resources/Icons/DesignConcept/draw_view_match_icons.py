"""
view_match（視点合わせ）と sectionbox_match（切断ボックス合わせ）のアイコンを描く。

旧アイコンのコピー側（左上）とペースト側（右下）を 60% に縮めて並べ、間に矢印を入れる
（viewport_match と同じ「コピー → 矢印 → ペースト」の考え方）。
96px で 288 DPI を入れて保存する（入れないとリボンからはみ出す。Docs/DEVLOG.md 参照）。

使い方（リポジトリのルートで）:
  python3 Resources/Icons/DesignConcept/draw_view_match_icons.py Resources/Icons Resources/Icons
"""
import sys, math
from PIL import Image, ImageDraw
icons, out = sys.argv[1], sys.argv[2]
SS = 8
def load(n): return Image.open(f'{icons}/{n}.png').convert('RGBA')

def arrow(img, x0, y0, x1, y1, w=3.6, head=11):
    big = Image.new('RGBA', (96*SS, 96*SS), (0,0,0,0)); d = ImageDraw.Draw(big)
    k=SS; x0,y0,x1,y1=[v*k for v in (x0,y0,x1,y1)]; head*=k
    ang = math.atan2(y1-y0, x1-x0)
    bx, by = x1-math.cos(ang)*head*0.8, y1-math.sin(ang)*head*0.8
    # 白い縁取り（背景の絵と重なっても読めるように）
    for col, ww, hh in [((255,255,255,255), w+3.0, head+3*k), ((50,50,50,255), w, head)]:
        d.line([(x0,y0),(bx,by)], fill=col, width=int(ww*k))
        l=(x1-math.cos(ang-0.5)*hh, y1-math.sin(ang-0.5)*hh); r=(x1-math.cos(ang+0.5)*hh, y1-math.sin(ang+0.5)*hh)
        d.polygon([(x1,y1),l,r], fill=col)
    img.alpha_composite(big.resize((96,96), Image.LANCZOS))

def combo(src, dst, scale, src_pos, dst_pos, arr):
    img = Image.new('RGBA', (96,96), (0,0,0,0))
    n = int(96*scale)
    img.alpha_composite(load(src).resize((n,n), Image.LANCZOS), src_pos)
    img.alpha_composite(load(dst).resize((n,n), Image.LANCZOS), dst_pos)
    arrow(img, *arr)
    return img

view = combo('view_copy', 'view_paste', 0.60, (-2,-2), (38,38), (40,22, 64,40))
sbox = combo('sectionbox_copy', 'sectionbox_paste', 0.60, (-2,-2), (38,38), (40,22, 64,40))
for name, im in [('view_match', view), ('sectionbox_match', sbox)]:
    im.save(f'{out}/{name}.png', dpi=(288.0106, 288.0106))


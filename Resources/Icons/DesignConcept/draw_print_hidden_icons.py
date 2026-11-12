# 印刷非表示 ON/OFF（PrintHiddenText）のアイコンを作るスクリプト
# 使い方: python3 draw_print_hidden_icons.py <出力フォルダ>
# 絵柄: print_hidden_source.png（運営者が用意したフリー素材のプリンタ線画）を
#       共通パレットの DARK（#323232）に塗り替え、右下に状態バッジ（ON=緑のチェック / OFF=赤の×）を重ねる
# 96px は 288DPI（pHYs）で保存する（無いと WPF でリボンからはみ出す。Docs/DEVLOG.md 参照）
import sys, os
from PIL import Image, ImageDraw, ImageOps, ImageFilter
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = sys.argv[1]; os.makedirs(OUT, exist_ok=True)
S = 24                      # 32単位 → 768px で描いて縮小（アンチエイリアス）
W = 32 * S
DARK = (0x32, 0x32, 0x32)
def p(v): return round(v * S)

# 素材の線画 → 濃さを透明度にした DARK 一色の画像（白い背景は透明になる）
# 素材の線は縮小すると細く薄くなるので、太らせて（MaxFilter）ほかのアイコンと線の太さをそろえる。
# 16px はさらに太らせる（細部はつぶれるが、形が見えることを優先）
src = Image.open(os.path.join(HERE, 'print_hidden_source.png')).convert('RGBA')
white = Image.new('RGBA', src.size, (255, 255, 255, 255))
gray = Image.alpha_composite(white, src).convert('L')
base_alpha = ImageOps.invert(gray).point(lambda v: 0 if v < 20 else min(255, int(v * 255 / 190)))

def make_art(grow):
    a = base_alpha.filter(ImageFilter.MaxFilter(grow)) if grow > 1 else base_alpha
    art = Image.new('RGBA', src.size, DARK + (255,)); art.putalpha(a)
    # 32 単位の枠に、余白 1 単位で縦横比を保って収める
    box = 30 * S
    scale = min(box / art.width, box / art.height)
    return art.resize((round(art.width * scale), round(art.height * scale)), Image.LANCZOS)

def draw(on, art):
    im = Image.new('RGBA', (W, W), (0, 0, 0, 0))
    im.alpha_composite(art, ((W - art.width) // 2, (W - art.height) // 2))
    d = ImageDraw.Draw(im)
    cx, cy, r = 25.0, 25.0, 6.6
    d.ellipse([p(cx - r), p(cy - r), p(cx + r), p(cy + r)], fill='#ffffff')
    d.ellipse([p(cx - r + 0.9), p(cy - r + 0.9), p(cx + r - 0.9), p(cy + r - 0.9)],
              fill='#2e9e4f' if on else '#d9534f')
    if on:
        d.line([(p(cx - 3.0), p(cy + 0.2)), (p(cx - 0.8), p(cy + 2.5)), (p(cx + 3.2), p(cy - 2.4))],
               fill='#ffffff', width=p(1.6), joint='curve')
    else:
        k = 2.5
        for a, b in (((cx - k, cy - k), (cx + k, cy + k)), ((cx - k, cy + k), (cx + k, cy - k))):
            d.line([(p(a[0]), p(a[1])), (p(b[0]), p(b[1]))], fill='#ffffff', width=p(1.6))
    return im

art96, art16 = make_art(23), make_art(45)
for on, name in ((True, 'print_hidden_on'), (False, 'print_hidden_off')):
    draw(on, art96).resize((96, 96), Image.LANCZOS).save(f'{OUT}/{name}.png', dpi=(288, 288))
    draw(on, art16).resize((16, 16), Image.LANCZOS).save(f'{OUT}/{name}_16.png')

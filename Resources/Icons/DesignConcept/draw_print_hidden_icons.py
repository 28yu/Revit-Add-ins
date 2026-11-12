# 印刷非表示 ON/OFF（PrintHiddenText）のアイコンを描くスクリプト
# 使い方: python3 draw_print_hidden_icons.py <出力フォルダ>
# 96px は 288DPI（pHYs）で保存する（無いと WPF でリボンからはみ出す。Docs/DEVLOG.md 参照）
# 絵柄: プリンタ（共通パレットの DARK/MID）＋右下に状態バッジ（ON=緑のチェック / OFF=赤の×）
import sys, os
from PIL import Image, ImageDraw
OUT = sys.argv[1]; os.makedirs(OUT, exist_ok=True)
S = 24  # 32単位 → 768px で描いて縮小（アンチエイリアス）
DARK, MID = '#323232', '#5C5C5C'
def p(v): return round(v * S)

def draw(on):
    im = Image.new('RGBA', (32 * S, 32 * S), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
    # 給紙側の用紙
    d.rectangle([p(8), p(2), p(24), p(12)], fill='#ffffff', outline=MID, width=p(1.3))
    # プリンタ本体
    d.rounded_rectangle([p(2), p(10), p(30), p(23)], radius=p(2.5), fill=MID, outline=DARK, width=p(1.3))
    # 排紙された用紙
    d.rectangle([p(8), p(19), p(24), p(30)], fill='#ffffff', outline=MID, width=p(1.3))
    # 状態バッジ（白縁つき）
    cx, cy, r = 24.6, 24.6, 6.4
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

for on, name in ((True, 'print_hidden_on'), (False, 'print_hidden_off')):
    im = draw(on)
    im.resize((96, 96), Image.LANCZOS).save(f'{OUT}/{name}.png', dpi=(288, 288))
    im.resize((16, 16), Image.LANCZOS).save(f'{OUT}/{name}_16.png')

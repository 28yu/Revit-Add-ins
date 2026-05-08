#!/usr/bin/env python3
"""
Generate redesigned icon PNGs for Tools28 Revit add-in.
Targets: sectionbox_copy, sectionbox_paste, cropbox_copy, cropbox_paste
Output: Resources/Icons/*_96.png (96x96, DPI=288)
"""

from PIL import Image, ImageDraw
import math
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(SCRIPT_DIR, '..', 'Resources', 'Icons')

# Supersampling scale: draw at 4x then downsample to 96x96
SS = 4

def s(v):
    """Scale from 32-coord-space to supersampled pixel space (3 * SS)"""
    return v * 3 * SS

def hex_rgba(h, a=255):
    h = h.lstrip('#')
    return (int(h[0:2],16), int(h[2:4],16), int(h[4:6],16), a)

def iw(v):
    """Integer width (minimum 1)"""
    return max(1, round(v))

def dashed_line(draw, x1, y1, x2, y2, color, width, dash_len, gap_len):
    dx, dy = x2 - x1, y2 - y1
    length = math.hypot(dx, dy)
    if length < 0.001:
        return
    ux, uy = dx / length, dy / length
    pos = 0.0
    draw_seg = True
    w = iw(width)
    while pos < length:
        seg = dash_len if draw_seg else gap_len
        end = min(pos + seg, length)
        if draw_seg:
            p1 = (x1 + ux * pos, y1 + uy * pos)
            p2 = (x1 + ux * end, y1 + uy * end)
            draw.line([p1, p2], fill=color, width=w)
        pos = end
        draw_seg = not draw_seg

def filled_poly(draw, pts, fill, stroke, stroke_w):
    draw.polygon(pts, fill=fill)
    sw = iw(stroke_w)
    n = len(pts)
    for i in range(n):
        draw.line([pts[i], pts[(i+1) % n]], fill=stroke, width=sw)

def bubble(draw, cx, cy, r, fill, stroke, stroke_w):
    draw.ellipse(
        [cx - r, cy - r, cx + r, cy + r],
        fill=fill,
        outline=stroke,
        width=iw(stroke_w)
    )

def new_canvas():
    size = 96 * SS
    return Image.new('RGBA', (size, size), (0, 0, 0, 0))

def save_icon(img, name):
    result = img.resize((96, 96), Image.LANCZOS)
    path = os.path.join(OUT_DIR, name + '_96.png')
    result.save(path, 'PNG', dpi=(288, 288))
    print(f'  Saved: {path}')

# ─────────────────────────────────────────
# sectionbox_copy
# ─────────────────────────────────────────
def make_sectionbox_copy():
    img = new_canvas()
    draw = ImageDraw.Draw(img)

    BLUE = hex_rgba('#0066CC')
    EDGE = hex_rgba('#3C3C3C')

    fx1, fy1, fx2, fy2 = 2, 8, 25, 30
    dox, doy = 5, -6
    bx1, by1 = fx1+dox, fy1+doy
    bx2, by2 = fx2+dox, fy2+doy

    lw = s(1.5)
    dk, gp = s(2.5), s(1.5)

    # Step 1: Back face (dashed)
    dashed_line(draw, s(bx1),s(by1), s(bx2),s(by1), BLUE, lw, dk, gp)
    dashed_line(draw, s(bx2),s(by1), s(bx2),s(by2), BLUE, lw, dk, gp)
    dashed_line(draw, s(bx2),s(by2), s(bx1),s(by2), BLUE, lw, dk, gp)
    dashed_line(draw, s(bx1),s(by2), s(bx1),s(by1), BLUE, lw, dk, gp)
    # Depth lines (dashed)
    dashed_line(draw, s(fx1),s(fy1), s(bx1),s(by1), BLUE, lw, dk, gp)
    dashed_line(draw, s(fx2),s(fy1), s(bx2),s(by1), BLUE, lw, dk, gp)
    dashed_line(draw, s(fx2),s(fy2), s(bx2),s(by2), BLUE, lw, dk, gp)
    dashed_line(draw, s(fx1),s(fy2), s(bx1),s(by2), BLUE, lw, dk, gp)

    # Step 2: Inner solid cube — front(9,14)-(18,23), depth(+4,-4)
    cf = (9, 14, 18, 23)
    cdx, cdy = 4, -4
    cb = (cf[0]+cdx, cf[1]+cdy, cf[2]+cdx, cf[3]+cdy)
    elw = s(0.7)

    # Top face
    top = [(s(cf[0]),s(cf[1])), (s(cf[2]),s(cf[1])),
           (s(cb[2]),s(cb[1])), (s(cb[0]),s(cb[1]))]
    filled_poly(draw, top, hex_rgba('#A8B8C8'), EDGE, elw)
    # Right face
    right = [(s(cf[2]),s(cf[1])), (s(cb[2]),s(cb[1])),
             (s(cb[2]),s(cb[3])), (s(cf[2]),s(cf[3]))]
    filled_poly(draw, right, hex_rgba('#7A8A9A'), EDGE, elw)
    # Front face
    front = [(s(cf[0]),s(cf[1])), (s(cf[2]),s(cf[1])),
             (s(cf[2]),s(cf[3])), (s(cf[0]),s(cf[3]))]
    filled_poly(draw, front, hex_rgba('#B8C8D8'), EDGE, elw)

    # Step 3: Front face on top (dashed)
    dashed_line(draw, s(fx1),s(fy1), s(fx2),s(fy1), BLUE, lw, dk, gp)
    dashed_line(draw, s(fx2),s(fy1), s(fx2),s(fy2), BLUE, lw, dk, gp)
    dashed_line(draw, s(fx2),s(fy2), s(fx1),s(fy2), BLUE, lw, dk, gp)
    dashed_line(draw, s(fx1),s(fy2), s(fx1),s(fy1), BLUE, lw, dk, gp)

    save_icon(img, 'sectionbox_copy')

# ─────────────────────────────────────────
# sectionbox_paste
# ─────────────────────────────────────────
def make_sectionbox_paste():
    img = new_canvas()
    draw = ImageDraw.Draw(img)

    BLUE = hex_rgba('#2288DD')
    EDGE = hex_rgba('#3C3C3C')

    fx1, fy1, fx2, fy2 = 2, 8, 25, 30
    dox, doy = 5, -6
    bx1, by1 = fx1+dox, fy1+doy
    bx2, by2 = fx2+dox, fy2+doy

    lw = iw(s(1.5))

    # Step 1: Back face (solid)
    draw.line([(s(bx1),s(by1)), (s(bx2),s(by1))], fill=BLUE, width=lw)
    draw.line([(s(bx2),s(by1)), (s(bx2),s(by2))], fill=BLUE, width=lw)
    draw.line([(s(bx2),s(by2)), (s(bx1),s(by2))], fill=BLUE, width=lw)
    draw.line([(s(bx1),s(by2)), (s(bx1),s(by1))], fill=BLUE, width=lw)
    # Depth lines (solid)
    draw.line([(s(fx1),s(fy1)), (s(bx1),s(by1))], fill=BLUE, width=lw)
    draw.line([(s(fx2),s(fy1)), (s(bx2),s(by1))], fill=BLUE, width=lw)
    draw.line([(s(fx2),s(fy2)), (s(bx2),s(by2))], fill=BLUE, width=lw)
    draw.line([(s(fx1),s(fy2)), (s(bx1),s(by2))], fill=BLUE, width=lw)

    # Step 2: Inner solid cube — blue tones
    cf = (9, 14, 18, 23)
    cdx, cdy = 4, -4
    cb = (cf[0]+cdx, cf[1]+cdy, cf[2]+cdx, cf[3]+cdy)
    elw = s(0.7)

    top = [(s(cf[0]),s(cf[1])), (s(cf[2]),s(cf[1])),
           (s(cb[2]),s(cb[1])), (s(cb[0]),s(cb[1]))]
    filled_poly(draw, top, hex_rgba('#5090D0'), EDGE, elw)
    right = [(s(cf[2]),s(cf[1])), (s(cb[2]),s(cb[1])),
             (s(cb[2]),s(cb[3])), (s(cf[2]),s(cf[3]))]
    filled_poly(draw, right, hex_rgba('#2870B0'), EDGE, elw)
    front = [(s(cf[0]),s(cf[1])), (s(cf[2]),s(cf[1])),
             (s(cf[2]),s(cf[3])), (s(cf[0]),s(cf[3]))]
    filled_poly(draw, front, hex_rgba('#70B0E0'), EDGE, elw)

    # Step 3: Front face on top (solid)
    draw.line([(s(fx1),s(fy1)), (s(fx2),s(fy1))], fill=BLUE, width=lw)
    draw.line([(s(fx2),s(fy1)), (s(fx2),s(fy2))], fill=BLUE, width=lw)
    draw.line([(s(fx2),s(fy2)), (s(fx1),s(fy2))], fill=BLUE, width=lw)
    draw.line([(s(fx1),s(fy2)), (s(fx1),s(fy1))], fill=BLUE, width=lw)

    save_icon(img, 'sectionbox_paste')

# ─────────────────────────────────────────
# cropbox_copy
# ─────────────────────────────────────────
def make_cropbox_copy():
    img = new_canvas()
    draw = ImageDraw.Draw(img)

    DARK = hex_rgba('#323232')
    WHITE = (255, 255, 255, 255)
    BLUE = hex_rgba('#0066CC')
    R = 2

    # Step 1: Grid lines
    lw = iw(s(1))
    draw.line([(s(10),s(4+R)), (s(10),s(28-R))], fill=DARK, width=lw)
    draw.line([(s(22),s(4+R)), (s(22),s(28-R))], fill=DARK, width=lw)
    draw.line([(s(4+R),s(10)), (s(28-R),s(10))], fill=DARK, width=lw)
    draw.line([(s(4+R),s(22)), (s(28-R),s(22))], fill=DARK, width=lw)

    # Step 2: Bubbles (R=2)
    sw = iw(s(1.2))
    for bx, by in [(10,4),(22,4),(4,10),(28,10),(4,22),(28,22),(10,28),(22,28)]:
        bubble(draw, s(bx), s(by), s(R), WHITE, DARK, sw)

    # Step 3: Blue dashed rectangle (8,8)-(24,24)
    blw = s(1.2)
    dk, gp = s(2.5), s(1.5)
    dashed_line(draw, s(8),s(8),  s(24),s(8),  BLUE, blw, dk, gp)
    dashed_line(draw, s(24),s(8), s(24),s(24), BLUE, blw, dk, gp)
    dashed_line(draw, s(24),s(24),s(8), s(24), BLUE, blw, dk, gp)
    dashed_line(draw, s(8),s(24), s(8), s(8),  BLUE, blw, dk, gp)

    save_icon(img, 'cropbox_copy')

# ─────────────────────────────────────────
# cropbox_paste
# ─────────────────────────────────────────
def make_cropbox_paste():
    img = new_canvas()
    draw = ImageDraw.Draw(img)

    DARK = hex_rgba('#323232')
    WHITE = (255, 255, 255, 255)
    BLUE = hex_rgba('#2288DD')
    R = 2

    # Step 1: Grid lines
    lw = iw(s(1))
    draw.line([(s(10),s(4+R)), (s(10),s(28-R))], fill=DARK, width=lw)
    draw.line([(s(22),s(4+R)), (s(22),s(28-R))], fill=DARK, width=lw)
    draw.line([(s(4+R),s(10)), (s(28-R),s(10))], fill=DARK, width=lw)
    draw.line([(s(4+R),s(22)), (s(28-R),s(22))], fill=DARK, width=lw)

    # Step 2: Bubbles (R=2)
    sw = iw(s(1.2))
    for bx, by in [(10,4),(22,4),(4,10),(28,10),(4,22),(28,22),(10,28),(22,28)]:
        bubble(draw, s(bx), s(by), s(R), WHITE, DARK, sw)

    # Step 3: Blue solid rectangle (8,8)-(24,24)
    blw = iw(s(1.2))
    # Pillow rectangle uses inclusive coords; adjust for line width
    draw.rectangle([s(8), s(8), s(24), s(24)], outline=BLUE, width=blw)

    save_icon(img, 'cropbox_paste')

# ─────────────────────────────────────────
if __name__ == '__main__':
    print('Generating icons...')
    make_sectionbox_copy()
    make_sectionbox_paste()
    make_cropbox_copy()
    make_cropbox_paste()
    print('Done.')

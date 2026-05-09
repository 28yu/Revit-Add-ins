#!/usr/bin/env python3
"""
Generate redesigned icon PNGs for Tools28 Revit add-in.
Targets: sectionbox_copy, sectionbox_paste, cropbox_copy, cropbox_paste, filled_region
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

def dashdot_line(draw, x1, y1, x2, y2, color, width):
    """Draw a dash-dot line (長破線・ギャップ・点・ギャップ…)."""
    dx, dy = x2 - x1, y2 - y1
    length = math.hypot(dx, dy)
    if length < 0.001:
        return
    ux, uy = dx / length, dy / length
    lens = [s(3), s(1.2), s(0.5), s(1.2)]  # dash, gap, dot, gap
    w = iw(width)
    pos, phase = 0.0, 0
    while pos < length:
        seg = lens[phase % 4]
        end = min(pos + seg, length)
        if phase % 2 == 0:  # draw segment (dash or dot)
            p1 = (x1 + ux * pos, y1 + uy * pos)
            p2 = (x1 + ux * end, y1 + uy * end)
            draw.line([p1, p2], fill=color, width=w)
        pos = end
        phase += 1

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

def save_icon_as(img, filename):
    """Save 384×384 draw canvas as 96×96 @ DPI=288 with exact filename."""
    result = img.resize((96, 96), Image.LANCZOS)
    path = os.path.join(OUT_DIR, filename)
    result.save(path, 'PNG', dpi=(288.0106, 288.0106))
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
# filled_region
# ─────────────────────────────────────────
def make_filled_region():
    img = new_canvas()
    draw = ImageDraw.Draw(img)

    GRAY = hex_rgba('#646464')

    def hatch_rect(rx, ry, rw, rh, line_w, step, cross):
        """Draw hatched rectangle with clip using Pillow (mask approach).
        All parameters in 32-coord space."""
        hatch = Image.new('RGBA', img.size, (0, 0, 0, 0))
        hd = ImageDraw.Draw(hatch)
        lw = iw(s(line_w))
        # \ diagonal lines: d steps in 32-coord space
        d = -rh
        while d <= rw:
            hd.line([(s(rx + d), s(ry)), (s(rx + d + rh), s(ry + rh))],
                    fill=GRAY, width=lw)
            d += step
        if cross:
            # / diagonal lines
            d = 0
            while d <= rw + rh:
                hd.line([(s(rx + d), s(ry)), (s(rx + d - rh), s(ry + rh))],
                        fill=GRAY, width=lw)
                d += step
        # Clip to rectangle using mask
        mask = Image.new('L', img.size, 0)
        md = ImageDraw.Draw(mask)
        md.rectangle([s(rx), s(ry), s(rx + rw), s(ry + rh)], fill=255)
        img.paste(hatch, mask=mask)
        # Border on top
        draw.rectangle([s(rx), s(ry), s(rx + rw), s(ry + rh)],
                       outline=GRAY, width=iw(s(0.8)))

    # Left: cross-hatch (\ + /), step=4 in 32-coord
    hatch_rect(1.5, 1.5, 12, 22, 0.8, 4, True)
    # Right: single \ diagonal, step=4 in 32-coord
    hatch_rect(18.5, 1.5, 12, 22, 0.7, 4, False)

    # Arrows at bottom
    ay = s(27)
    aw = iw(s(0.7))
    # Left arrow <
    draw.line([(s(12), s(24.5)), (s(8), ay)], fill=GRAY, width=aw)
    draw.line([(s(8), ay), (s(12), s(29.5))], fill=GRAY, width=aw)
    # Right arrow >
    draw.line([(s(20), s(24.5)), (s(24), ay)], fill=GRAY, width=aw)
    draw.line([(s(24), ay), (s(20), s(29.5))], fill=GRAY, width=aw)
    # Connecting line
    draw.line([(s(8), ay), (s(24), ay)], fill=GRAY, width=aw)

    save_icon(img, 'filled_region')

# ─────────────────────────────────────────
# beam_under_level
# ─────────────────────────────────────────
def make_beam_under_level():
    img = new_canvas()
    draw = ImageDraw.Draw(img)

    DARK = hex_rgba('#505050')
    MID  = hex_rgba('#787878')
    PINK   = (255, 128, 148, 255)
    YELLOW = (218, 185,  47, 255)
    BLUE_C = ( 30, 144, 255, 255)

    # I-beam
    draw.rectangle([s(3), s(1),  s(17), s(4)],  fill=MID)  # top flange
    draw.rectangle([s(3), s(14), s(17), s(17)], fill=MID)  # bottom flange (bottom=y17)
    draw.rectangle([s(8), s(4),  s(12), s(14)], fill=MID)  # web

    # Up arrow: tip at y=17 (beam bottom), head base at y=21, shaft y=21→25
    lw = iw(s(0.8))
    draw.line([(s(10), s(21)), (s(10), s(25))], fill=DARK, width=lw)
    draw.polygon([
        (s(10), s(17)),    # tip
        (s(7.5), s(21)),   # left
        (s(12.5), s(21)),  # right
    ], fill=DARK)

    # Dash-dot FL line at y=27 (一点破線)
    dashdot_line(draw, s(1), s(27), s(20), s(27), DARK, s(0.8))

    # ▼ Triangle: base at top (y=23), vertex at FL line (y=27)
    draw.polygon([
        (s(2.5), s(23)),  # top-left
        (s(5.5), s(23)),  # top-right
        (s(4),   s(27)),  # bottom vertex touching FL line
    ], fill=DARK)

    # 3 color blocks on right
    bx, bw, bh = s(22), s(8), s(8)
    for color, y in [(PINK, 1), (YELLOW, 11), (BLUE_C, 21)]:
        draw.rectangle([bx, s(y), bx + bw, s(y) + bh], fill=color)

    save_icon(img, 'beam_under_level')

# ─────────────────────────────────────────
# beam_top_level
# ─────────────────────────────────────────
def make_beam_top_level():
    img = new_canvas()
    draw = ImageDraw.Draw(img)

    DARK = hex_rgba('#505050')
    MID  = hex_rgba('#787878')
    PINK   = (255, 128, 148, 255)
    YELLOW = (218, 185,  47, 255)
    BLUE_C = ( 30, 144, 255, 255)

    # Dash-dot FL line at y=5 (一点破線)
    dashdot_line(draw, s(1), s(5), s(20), s(5), DARK, s(0.8))

    # ▼ Triangle: base at top (y=1), vertex (bottom) touching FL line at y=5
    draw.polygon([
        (s(2.5), s(1)),  # top-left
        (s(5.5), s(1)),  # top-right
        (s(4),   s(5)),  # bottom vertex touching FL line
    ], fill=DARK)

    # Down arrow: shaft y=7→y=11, tip at beam top (y=15)
    lw = iw(s(0.8))
    draw.line([(s(10), s(7)), (s(10), s(11))], fill=DARK, width=lw)
    draw.polygon([
        (s(10),   s(15)),  # tip touches beam top
        (s(7.5),  s(11)),  # left base
        (s(12.5), s(11)),  # right base
    ], fill=DARK)

    # I-beam at bottom: same shape as beam_under_level (flange3+web10+flange3)
    # Top edge y=15, bottom edge y=31
    draw.rectangle([s(3), s(15), s(17), s(18)], fill=MID)  # top flange
    draw.rectangle([s(3), s(28), s(17), s(31)], fill=MID)  # bottom flange
    draw.rectangle([s(8), s(18), s(12), s(28)], fill=MID)  # web

    # 3 color blocks on right
    bx, bw, bh = s(22), s(8), s(8)
    for color, y in [(PINK, 1), (YELLOW, 11), (BLUE_C, 21)]:
        draw.rectangle([bx, s(y), bx + bw, s(y) + bh], fill=color)

    save_icon(img, 'beam_top_level')

# ─────────────────────────────────────────
# excel_export / excel_import
# ─────────────────────────────────────────
def _draw_excel_file_icon(img):
    """Draw Excel file icon: paper(top-right fold) + green square(black border, extends left) + flat-top X + dashes."""
    draw = ImageDraw.Draw(img)
    DARK  = (26, 26, 26, 255)
    WHITE = (255, 255, 255, 255)
    GREEN = hex_rgba('#217346')
    sw = iw(s(1.3))  # thick black border

    # Paper pentagon: (4,2)-(16,2)-(21,7)-(21,26)-(4,26) — width=17, height=24, ratio 1:1.41
    paper_poly = [
        (s(4),  s(2)),
        (s(16), s(2)),
        (s(21), s(7)),
        (s(21), s(26)),
        (s(4),  s(26)),
    ]
    draw.polygon(paper_poly, fill=WHITE, outline=DARK, width=sw)

    # Fold flap triangle: (16,2)-(21,2)-(21,7)
    fold_poly = [
        (s(16), s(2)),
        (s(21), s(2)),
        (s(21), s(7)),
    ]
    draw.polygon(fold_poly, fill=WHITE, outline=DARK, width=sw)

    # Green square with thick black border: x=1..13, y=8..20 (12×12)
    draw.rectangle([s(1), s(8), s(13), s(20)], fill=GREEN, outline=DARK, width=sw)

    # White X with flat horizontal cuts — 2.5-unit padding inside green (1..13, 8..20)
    # X bounds: x=3.5..10.5, y=10.5..17.5, stroke width=2.5
    # Left stroke (top-left to bottom-right)
    draw.polygon([
        (s(3.5),  s(10.5)),
        (s(6.0),  s(10.5)),
        (s(10.5), s(17.5)),
        (s(8.0),  s(17.5)),
    ], fill=WHITE)
    # Right stroke (top-right to bottom-left)
    draw.polygon([
        (s(8.0),  s(10.5)),
        (s(10.5), s(10.5)),
        (s(6.0),  s(17.5)),
        (s(3.5),  s(17.5)),
    ], fill=WHITE)

    # Green dashes: 3 rows × 2 columns, centered in paper (y=2..26 → rows at 10,14,18)
    # Right column ends at x=20 to stay clear of paper border at x=21
    dash_w = iw(s(1.6))
    for y_32 in [10, 14, 18]:
        draw.line([(s(14),   s(y_32)), (s(17),   s(y_32))], fill=GREEN, width=dash_w)
        draw.line([(s(17.8), s(y_32)), (s(20),   s(y_32))], fill=GREEN, width=dash_w)


def make_excel_export():
    img = new_canvas()
    _draw_excel_file_icon(img)
    draw = ImageDraw.Draw(img)

    # Blue up arrow: shaft y=8→22, tip at y=6
    BLUE = (0, 102, 204, 255)
    lw = iw(s(1.2))
    draw.line([(s(27), s(8)), (s(27), s(22))], fill=BLUE, width=lw)
    draw.polygon([
        (s(27),   s(6)),
        (s(23.5), s(11)),
        (s(30.5), s(11)),
    ], fill=BLUE)

    save_icon(img, 'excel_export')


def make_excel_import():
    img = new_canvas()
    _draw_excel_file_icon(img)
    draw = ImageDraw.Draw(img)

    # Blue down arrow: shaft y=8→22, tip at y=26
    BLUE = (0, 102, 204, 255)
    lw = iw(s(1.2))
    draw.line([(s(27), s(8)), (s(27), s(22))], fill=BLUE, width=lw)
    draw.polygon([
        (s(27),   s(26)),
        (s(23.5), s(21)),
        (s(30.5), s(21)),
    ], fill=BLUE)

    save_icon(img, 'excel_import')


# ─────────────────────────────────────────
# Settings panel icons (flags, ver, manual)
# 16px logical  → 48×48 PNG @ DPI=288  (coord space 0..16, s() value range same)
# 32px logical  → 96×96 PNG @ DPI=288  (coord space 0..32, same as large icons)
# ─────────────────────────────────────────

SMALL_PHYS    = 48
SMALL_DRAW_PX = SMALL_PHYS * SS   # 192

def new_canvas_small():
    return Image.new('RGBA', (SMALL_DRAW_PX, SMALL_DRAW_PX), (0, 0, 0, 0))

def save_icon_small(draw_img, name):
    out = draw_img.resize((SMALL_PHYS, SMALL_PHYS), Image.LANCZOS)
    path = os.path.join(OUT_DIR, f'{name}.png')
    out.save(path, dpi=(288.0106, 288.0106))
    print(f'  Saved: {path}')

def star_pts(cx, cy, r_out, r_in, n=5, start_deg=-90):
    """Return vertices of an n-pointed star polygon."""
    pts = []
    for i in range(n * 2):
        a = math.radians(start_deg + i * 180.0 / n)
        r = r_out if i % 2 == 0 else r_in
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return pts

# Flag color constants
JP_RED   = (188,   0,  45, 255)
US_RED   = (178,  34,  52, 255)
US_BLUE  = ( 60,  59, 110, 255)
CN_RED   = (222,  41,  16, 255)
CN_YEL   = (255, 222,   0, 255)
ICON_BLU = ( 37,  99, 235, 255)
FG_WHITE = (255, 255, 255, 255)
BORDER_K = ( 30,  30,  30, 255)


def make_flag_jp():
    # ── 16px logical (48px physical, 16-unit coord space) ──────────────
    img = new_canvas_small()
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, SMALL_DRAW_PX - 1, SMALL_DRAW_PX - 1], fill=FG_WHITE)
    r = s(4.8)
    d.ellipse([s(8) - r, s(8) - r, s(8) + r, s(8) + r], fill=JP_RED)
    d.rectangle([0, 0, SMALL_DRAW_PX - 1, SMALL_DRAW_PX - 1],
                outline=BORDER_K, width=iw(s(0.5)))
    save_icon_small(img, 'flag_jp_16')

    # ── 32px logical (96px physical, 32-unit coord space) ───────────────
    img2 = new_canvas()
    d2 = ImageDraw.Draw(img2)
    d2.rectangle([0, 0, 383, 383], fill=FG_WHITE)
    r2 = s(9.5)
    d2.ellipse([s(16) - r2, s(16) - r2, s(16) + r2, s(16) + r2], fill=JP_RED)
    d2.rectangle([0, 0, 383, 383], outline=BORDER_K, width=iw(s(0.5)))
    save_icon_as(img2, 'flag_jp_32.png')


def make_flag_us():
    # ── 16px logical ─────────────────────────────────────────────────────
    img = new_canvas_small()
    d = ImageDraw.Draw(img)
    stripe_h = 16.0 / 7
    for i in range(7):
        y0 = round(s(i * stripe_h))
        y1 = round(s((i + 1) * stripe_h))
        d.rectangle([0, y0, SMALL_DRAW_PX - 1, y1],
                    fill=US_RED if i % 2 == 0 else FG_WHITE)
    # Blue canton: 40% width × 4 stripe heights
    canton_x = round(s(6.5))
    canton_y = round(s(4 * stripe_h))
    d.rectangle([0, 0, canton_x, canton_y], fill=US_BLUE)
    # Stars: 3 cols × 2 rows (simplified)
    for sy in [1.5, 4.5]:
        for sx in [1.4, 3.2, 5.0]:
            r = s(0.65)
            d.ellipse([s(sx) - r, s(sy) - r, s(sx) + r, s(sy) + r], fill=FG_WHITE)
    d.rectangle([0, 0, SMALL_DRAW_PX - 1, SMALL_DRAW_PX - 1],
                outline=BORDER_K, width=iw(s(0.5)))
    save_icon_small(img, 'flag_us_16')

    # ── 32px logical ─────────────────────────────────────────────────────
    img2 = new_canvas()
    d2 = ImageDraw.Draw(img2)
    stripe_h2 = 32.0 / 7
    for i in range(7):
        y0 = round(s(i * stripe_h2))
        y1 = round(s((i + 1) * stripe_h2))
        d2.rectangle([0, y0, 383, y1],
                     fill=US_RED if i % 2 == 0 else FG_WHITE)
    canton_x2 = round(s(13.0))
    canton_y2 = round(s(4 * stripe_h2))
    d2.rectangle([0, 0, canton_x2, canton_y2], fill=US_BLUE)
    for sy2 in [2.5, 6.0, 9.5, 13.0]:
        for sx2 in [2.0, 5.5, 9.0, 12.0]:
            r2 = s(1.3)
            d2.ellipse([s(sx2) - r2, s(sy2) - r2, s(sx2) + r2, s(sy2) + r2],
                       fill=FG_WHITE)
    d2.rectangle([0, 0, 383, 383], outline=BORDER_K, width=iw(s(0.5)))
    save_icon_as(img2, 'flag_us_32.png')


def make_flag_cn():
    # ── 16px logical ─────────────────────────────────────────────────────
    img = new_canvas_small()
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, SMALL_DRAW_PX - 1, SMALL_DRAW_PX - 1], fill=CN_RED)
    # Large star: center (4,5) r_out=3.0 r_in=1.2
    d.polygon(star_pts(s(4), s(5), s(3.0), s(1.2)), fill=CN_YEL)
    # 4 small stars in arc (upper-right of large star)
    for sx, sy in [(9.0, 2.5), (11.0, 4.5), (11.0, 7.0), (9.0, 9.0)]:
        d.polygon(star_pts(s(sx), s(sy), s(1.3), s(0.55)), fill=CN_YEL)
    d.rectangle([0, 0, SMALL_DRAW_PX - 1, SMALL_DRAW_PX - 1],
                outline=BORDER_K, width=iw(s(0.5)))
    save_icon_small(img, 'flag_cn_16')

    # ── 32px logical ─────────────────────────────────────────────────────
    img2 = new_canvas()
    d2 = ImageDraw.Draw(img2)
    d2.rectangle([0, 0, 383, 383], fill=CN_RED)
    d2.polygon(star_pts(s(8), s(10), s(6.0), s(2.4)), fill=CN_YEL)
    for sx2, sy2 in [(18.0, 5.0), (22.0, 9.0), (22.0, 14.0), (18.0, 18.0)]:
        d2.polygon(star_pts(s(sx2), s(sy2), s(2.6), s(1.1)), fill=CN_YEL)
    d2.rectangle([0, 0, 383, 383], outline=BORDER_K, width=iw(s(0.5)))
    save_icon_as(img2, 'flag_cn_32.png')


def make_ver():
    """Version info icon: blue circle with white 'i' symbol."""
    # ── 16px logical ─────────────────────────────────────────────────────
    img = new_canvas_small()
    d = ImageDraw.Draw(img)
    bg_r = s(7.5)
    d.ellipse([s(8) - bg_r, s(8) - bg_r, s(8) + bg_r, s(8) + bg_r], fill=ICON_BLU)
    # Dot of 'i'
    dot_r = s(1.2)
    d.ellipse([s(8) - dot_r, s(4) - dot_r, s(8) + dot_r, s(4) + dot_r], fill=FG_WHITE)
    # Bar of 'i'
    d.line([(s(8), s(6.5)), (s(8), s(13.0))], fill=FG_WHITE, width=iw(s(2.5)))
    save_icon_small(img, 'ver_16')

    # ── 32px logical ─────────────────────────────────────────────────────
    img2 = new_canvas()
    d2 = ImageDraw.Draw(img2)
    bg_r2 = s(15.0)
    d2.ellipse([s(16) - bg_r2, s(16) - bg_r2, s(16) + bg_r2, s(16) + bg_r2],
               fill=ICON_BLU)
    dot_r2 = s(2.5)
    d2.ellipse([s(16) - dot_r2, s(8) - dot_r2, s(16) + dot_r2, s(8) + dot_r2],
               fill=FG_WHITE)
    d2.line([(s(16), s(13.0)), (s(16), s(26.0))], fill=FG_WHITE, width=iw(s(5.0)))
    save_icon_as(img2, 'ver_32.png')


def make_manual():
    """Manual icon: blue circle with white '?' symbol."""
    # ── 16px logical ─────────────────────────────────────────────────────
    img = new_canvas_small()
    d = ImageDraw.Draw(img)
    bg_r = s(7.5)
    d.ellipse([s(8) - bg_r, s(8) - bg_r, s(8) + bg_r, s(8) + bg_r], fill=ICON_BLU)
    # '?' hook arc: center (8,5) r=2.5, from 200° clockwise to 90° (PIL: 0°=right, CW)
    arc_w = iw(s(2.2))
    arc_r = s(2.5)
    d.arc([s(8) - arc_r, s(5) - arc_r, s(8) + arc_r, s(5) + arc_r],
          start=200, end=90, fill=FG_WHITE, width=arc_w)
    # Vertical tail below arc endpoint
    d.line([(s(8), s(7.6)), (s(8), s(10.5))], fill=FG_WHITE, width=arc_w)
    # Dot
    dot_r = s(1.0)
    d.ellipse([s(8) - dot_r, s(12.5) - dot_r, s(8) + dot_r, s(12.5) + dot_r],
              fill=FG_WHITE)
    save_icon_small(img, 'manual_16')

    # ── 32px logical ─────────────────────────────────────────────────────
    img2 = new_canvas()
    d2 = ImageDraw.Draw(img2)
    bg_r2 = s(15.0)
    d2.ellipse([s(16) - bg_r2, s(16) - bg_r2, s(16) + bg_r2, s(16) + bg_r2],
               fill=ICON_BLU)
    arc_w2 = iw(s(4.0))
    arc_r2 = s(5.0)
    d2.arc([s(16) - arc_r2, s(10) - arc_r2, s(16) + arc_r2, s(10) + arc_r2],
           start=200, end=90, fill=FG_WHITE, width=arc_w2)
    d2.line([(s(16), s(15.2)), (s(16), s(21.0))], fill=FG_WHITE, width=arc_w2)
    dot_r2 = s(2.0)
    d2.ellipse([s(16) - dot_r2, s(25) - dot_r2, s(16) + dot_r2, s(25) + dot_r2],
               fill=FG_WHITE)
    save_icon_as(img2, 'manual_32.png')


# ─────────────────────────────────────────
if __name__ == '__main__':
    print('Generating icons...')
    make_sectionbox_copy()
    make_sectionbox_paste()
    make_cropbox_copy()
    make_cropbox_paste()
    make_filled_region()
    make_beam_under_level()
    make_beam_top_level()
    make_excel_export()
    make_excel_import()
    make_flag_jp()
    make_flag_us()
    make_flag_cn()
    make_ver()
    make_manual()
    print('Done.')

#!/usr/bin/env python3
"""
Generate redesigned icon PNGs for Tools28 Revit add-in.
Targets: sectionbox_copy, sectionbox_paste, cropbox_copy, cropbox_paste, filled_region
Output: Resources/Icons/*_96.png (96x96, DPI=288)
"""

from PIL import Image, ImageDraw, ImageFont
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


# Flag rect geometry (3:2 landscape ratio, centered in canvas)
# 16-unit canvas: flag y=2.5..13.5  (width=16, height=11, ratio≈1.45)
# 32-unit canvas: flag y=5..27      (width=32, height=22, ratio≈1.45)
FLAG16_Y0, FLAG16_Y1 = 2.5, 13.5
FLAG32_Y0, FLAG32_Y1 = 5.0, 27.0


def make_flag_jp():
    # ── 16px logical ─────────────────────────────────────────────────────
    img = new_canvas_small()
    d = ImageDraw.Draw(img)
    # White flag rect
    d.rectangle([0, round(s(FLAG16_Y0)), SMALL_DRAW_PX - 1, round(s(FLAG16_Y1))],
                fill=FG_WHITE)
    # Red circle centered in flag rect
    cy = s((FLAG16_Y0 + FLAG16_Y1) / 2)   # = s(8)
    r = s(3.3)                              # diameter ≈ 60% of flag height (11*0.6/2=3.3)
    d.ellipse([s(8) - r, cy - r, s(8) + r, cy + r], fill=JP_RED)
    d.rectangle([0, round(s(FLAG16_Y0)), SMALL_DRAW_PX - 1, round(s(FLAG16_Y1))],
                outline=BORDER_K, width=iw(s(0.5)))
    save_icon_small(img, 'flag_jp_16')

    # ── 32px logical ─────────────────────────────────────────────────────
    img2 = new_canvas()
    d2 = ImageDraw.Draw(img2)
    d2.rectangle([0, round(s(FLAG32_Y0)), 383, round(s(FLAG32_Y1))], fill=FG_WHITE)
    cy2 = s((FLAG32_Y0 + FLAG32_Y1) / 2)  # = s(16)
    r2 = s(6.6)
    d2.ellipse([s(16) - r2, cy2 - r2, s(16) + r2, cy2 + r2], fill=JP_RED)
    d2.rectangle([0, round(s(FLAG32_Y0)), 383, round(s(FLAG32_Y1))],
                 outline=BORDER_K, width=iw(s(0.5)))
    save_icon_as(img2, 'flag_jp_32.png')


def make_flag_us():
    def _draw_us(d, draw_w, y0, y1, n_stripes, canton_x_unit,
                 star_ro, star_ri, star_rows_unit, star_cols_unit):
        """Draw US flag: n_stripes stripes, canton, stars.
        star_rows_unit = row positions as multiples of stripe_h from y0.
        star_cols_unit = absolute x positions (units).
        """
        total_h = y1 - y0
        stripe_h = total_h / n_stripes
        for i in range(n_stripes):
            ya = round(s(y0 + i * stripe_h))
            yb = round(s(y0 + (i + 1) * stripe_h))
            d.rectangle([0, ya, draw_w - 1, yb],
                        fill=US_RED if i % 2 == 0 else FG_WHITE)
        canton_x = round(s(canton_x_unit))
        # Canton covers top (n_stripes+1)//2 stripes (all red stripes)
        n_canton = (n_stripes + 1) // 2
        canton_y = round(s(y0 + n_canton * stripe_h))
        d.rectangle([0, round(s(y0)), canton_x, canton_y], fill=US_BLUE)
        for sy in star_rows_unit:
            for sx in star_cols_unit:
                pts = star_pts(s(sx), s(y0 + sy * stripe_h),
                               s(star_ro), s(star_ri))
                d.polygon(pts, fill=FG_WHITE)

    # ── 16px logical: 7 stripes, 3 cols × 2 rows, canton 4 stripes × 6.4u ─
    # Uniform spacing: cols at x=6.4/(3*2)*[1,3,5]=[1.067,3.2,5.333]
    # rows in 4-stripe canton at stripe mult [1.0, 3.0]
    img = new_canvas_small()
    d = ImageDraw.Draw(img)
    _draw_us(d, SMALL_DRAW_PX,
             FLAG16_Y0, FLAG16_Y1,
             n_stripes=7,
             canton_x_unit=6.4,
             star_ro=0.65, star_ri=0.65 * 0.382,
             star_rows_unit=[1.0, 3.0],
             star_cols_unit=[1.067, 3.2, 5.333])
    d.rectangle([0, round(s(FLAG16_Y0)), SMALL_DRAW_PX - 1, round(s(FLAG16_Y1))],
                outline=BORDER_K, width=iw(s(0.5)))
    save_icon_small(img, 'flag_us_16')

    # ── 32px logical: 13 stripes (7R+6W), 4 cols × 3 rows, canton 7 stripes × 12.8u ─
    # Uniform spacing: cols at x=12.8/(4*2)*[1,3,5,7]=[1.6,4.8,8.0,11.2]
    # rows in 7-stripe canton at stripe mult [7/6, 7/2, 35/6] ≈ [1.167,3.5,5.833]
    img2 = new_canvas()
    d2 = ImageDraw.Draw(img2)
    _draw_us(d2, 384,
             FLAG32_Y0, FLAG32_Y1,
             n_stripes=13,
             canton_x_unit=12.8,
             star_ro=1.2, star_ri=1.2 * 0.382,
             star_rows_unit=[7/6, 7/2, 35/6],
             star_cols_unit=[1.6, 4.8, 8.0, 11.2])
    d2.rectangle([0, round(s(FLAG32_Y0)), 383, round(s(FLAG32_Y1))],
                 outline=BORDER_K, width=iw(s(0.5)))
    save_icon_as(img2, 'flag_us_32.png')


def make_flag_cn():
    """China flag using Wikimedia SVG proportions with slight left-shift.
    Base (900×600): large star (225,150), small (450,54)(525,105)(525,195)(450,246).
    Applied x_shift = -0.07 (7% of flag width to the left) to match actual flag appearance.
    """
    X_SHIFT = -0.07       # all stars: 7% left
    X_SHIFT_SMALL = -0.06  # small stars: additional 6% left

    def _draw_cn(d, draw_w, y0, flag_w, flag_h, big_r, small_r):
        y1 = y0 + flag_h
        d.rectangle([0, round(s(y0)), draw_w - 1, round(s(y1))], fill=CN_RED)

        def star_at(fx, fy, r):
            cx = s(flag_w * (fx + X_SHIFT))
            cy = s(y0 + flag_h * fy)
            d.polygon(star_pts(cx, cy, s(r), s(r * 0.382)), fill=CN_YEL)

        star_at(225/900, 150/600, big_r)
        for fx, fy in [(450/900, 54/600),
                       (525/900, 105/600),
                       (525/900, 195/600),
                       (450/900, 246/600)]:
            cx = s(flag_w * (fx + X_SHIFT + X_SHIFT_SMALL))
            cy = s(y0 + flag_h * fy)
            d.polygon(star_pts(cx, cy, s(small_r), s(small_r * 0.382)), fill=CN_YEL)

        d.rectangle([0, round(s(y0)), draw_w - 1, round(s(y1))],
                    outline=BORDER_K, width=iw(s(0.5)))

    # 16px: flag 16×11 units
    img = new_canvas_small()
    _draw_cn(ImageDraw.Draw(img), SMALL_DRAW_PX,
             FLAG16_Y0, 16, FLAG16_Y1 - FLAG16_Y0, big_r=1.8, small_r=0.65)
    save_icon_small(img, 'flag_cn_16')

    # 32px: flag 32×22 units
    img2 = new_canvas()
    _draw_cn(ImageDraw.Draw(img2), 384,
             FLAG32_Y0, 32, FLAG32_Y1 - FLAG32_Y0, big_r=3.5, small_r=1.2)
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
    """Manual icon: blue circle with white '?' rendered using TrueType font.
    Using font rendering guarantees a clean, recognisable '?' shape.
    """
    FONT_PATH = '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf'

    def _draw_icon(canvas, cx_u, cy_u, bg_r_u, font_sz_u):
        d = ImageDraw.Draw(canvas)
        cx, cy = s(cx_u), s(cy_u)
        bg_r = s(bg_r_u)
        d.ellipse([cx - bg_r, cy - bg_r, cx + bg_r, cy + bg_r], fill=ICON_BLU)
        try:
            fnt = ImageFont.truetype(FONT_PATH, size=iw(s(font_sz_u)))
            bb = d.textbbox((0, 0), '?', font=fnt)
            tw, th = bb[2] - bb[0], bb[3] - bb[1]
            tx = cx - tw / 2 - bb[0]
            ty = cy - th / 2 - bb[1]
            d.text((tx, ty), '?', font=fnt, fill=FG_WHITE)
        except Exception as e:
            print(f'  Font unavailable ({e}), using arc fallback')
            arc_r = s(bg_r_u * 0.42)
            arc_w = iw(arc_r * 0.28)
            d.arc([cx - arc_r, cy - arc_r * 1.1, cx + arc_r, cy + arc_r * 0.5],
                  start=230, end=90, fill=FG_WHITE, width=arc_w)
            d.line([(cx, cy + arc_r * 0.5), (cx, cy + arc_r * 0.85)],
                   fill=FG_WHITE, width=arc_w)
            dr = arc_r * 0.18
            d.ellipse([cx - dr, cy + arc_r * 1.1 - dr,
                       cx + dr, cy + arc_r * 1.1 + dr], fill=FG_WHITE)

    # ── 16px logical ─────────────────────────────────────────────────────
    img = new_canvas_small()
    _draw_icon(img, cx_u=8, cy_u=8, bg_r_u=7.5, font_sz_u=11.0)
    save_icon_small(img, 'manual_16')

    # ── 32px logical ─────────────────────────────────────────────────────
    img2 = new_canvas()
    _draw_icon(img2, cx_u=16, cy_u=16, bg_r_u=15.0, font_sz_u=22.0)
    save_icon_as(img2, 'manual_32.png')


# ─────────────────────────────────────────
# formwork (型枠数量算出)
# ─────────────────────────────────────────
def make_formwork():
    """型枠数量算出: コンクリート(灰) + 型枠板(橙縦縞) + 横端太(鋼灰横バー) + セパレータ(点線)"""
    CONC    = (185, 185, 180, 255)  # コンクリート
    PL_BG   = (212, 118,  55, 255)  # 型枠板 明面
    PL_BT   = (148,  66,  18, 255)  # 縦桁（バテン）暗
    WLR     = (168, 172, 180, 255)  # 横端太
    WLR_SH  = ( 95, 100, 112, 255)  # 横端太輪郭
    TIE     = (125, 125, 118, 255)  # セパレータ（点線）
    BK      = ( 38,  38,  38, 255)  # 輪郭

    img = new_canvas()
    d   = ImageDraw.Draw(img)

    y0, y1 = 2.0, 30.0   # 描画範囲
    uh = y1 - y0          # 28 units
    cx = 11.5             # コンクリート右端
    fw = 32.0 - cx        # 型枠幅
    n  = 4                # 縦板枚数

    # ① コンクリート
    d.rectangle([s(0), s(y0), s(cx), s(y1)], fill=CONC)
    d.rectangle([s(0), s(y0), s(cx), s(y1)], outline=BK, width=iw(s(0.45)))

    # ② 型枠板（縦板4枚 + 縦桁）
    for i in range(n):
        px0 = cx + fw * i / n
        px1 = cx + fw * (i + 1) / n
        d.rectangle([s(px0), s(y0), s(px1), s(y1)], fill=PL_BG)
        bt = fw / n * 0.30   # 縦桁幅
        d.rectangle([s(px1 - bt), s(y0), s(px1), s(y1)], fill=PL_BT)

    # 型枠全体の外枠
    d.rectangle([s(cx), s(y0), s(31.5), s(y1)], outline=BK, width=iw(s(0.45)))

    # ③ セパレータ点線（コンクリート内、横端太と同高）
    for fy in [0.25, 0.50, 0.75]:
        ty = y0 + uh * fy
        dashed_line(d, s(1.0), s(ty), s(cx - 0.8), s(ty),
                    TIE, s(0.65), s(1.5), s(0.9))

    # ④ 横端太（型枠面に重ねる）
    wh = 1.3   # 端太の半高
    for fy in [0.25, 0.50, 0.75]:
        ty = y0 + uh * fy
        d.rectangle([s(cx - 0.3), s(ty - wh), s(31.8), s(ty + wh)],
                    fill=WLR, outline=WLR_SH, width=iw(s(0.35)))

    save_icon(img, 'formwork')


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
    make_formwork()
    print('Done.')

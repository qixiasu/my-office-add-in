"""Generate colorful flat-style ribbon icons for Office Add-in buttons."""
from PIL import Image, ImageDraw
import math

SIZES = [16, 32, 80]
SCALE = 4  # supersampling factor for anti-aliasing

# Color palette
BLUE = (59, 130, 246)       # #3B82F6
TEAL = (16, 185, 129)       # #10B981
ROYAL = (79, 70, 229)       # #4F46E5
EMERALD = (16, 185, 129)    # #10B981
AMBER = (245, 158, 11)      # #F59E0B
PURPLE = (139, 92, 246)     # #8B5CF6
INDIGO = (99, 102, 241)     # #6366F1
WHITE = (255, 255, 255)
DARK_BG = (30, 41, 59)      # #1E293B

def filled_rounded_rect(draw, xy, radius, fill):
    """Draw a filled rounded rectangle using pies and rectangles."""
    x1, y1, x2, y2 = xy
    r = radius
    draw.pieslice([x1, y1, x1 + 2 * r, y1 + 2 * r], 180, 270, fill=fill)
    draw.pieslice([x2 - 2 * r, y1, x2, y1 + 2 * r], 270, 360, fill=fill)
    draw.pieslice([x1, y2 - 2 * r, x1 + 2 * r, y2], 90, 180, fill=fill)
    draw.pieslice([x2 - 2 * r, y2 - 2 * r, x2, y2], 0, 90, fill=fill)
    draw.rectangle([x1 + r, y1, x2 - r, y2], fill=fill)
    draw.rectangle([x1, y1 + r, x2, y2 - r], fill=fill)


def draw_concat(draw, w):
    """Two overlapping rounded squares (blue + teal) with a white arrow/plus."""
    margin = w * 0.08
    sq = w * 0.42
    r = w * 0.1

    # Left square (blue)
    lx = margin
    ly = (w - sq) / 2 + w * 0.02
    filled_rounded_rect(draw, [lx, ly, lx + sq, ly + sq], int(r), BLUE)

    # Right square (teal) - slightly overlapping
    rx = w - margin - sq
    ry = (w - sq) / 2 - w * 0.02
    filled_rounded_rect(draw, [rx, ry, rx + sq, ry + sq], int(r), TEAL)

    # White plus/merge symbol in the center gap
    cx = w / 2
    cy = w / 2
    arm = sq * 0.25
    lw_arm = max(2, int(w * 0.06))
    draw.line([cx - arm, cy, cx + arm, cy], fill=WHITE, width=lw_arm)
    draw.line([cx, cy - arm, cx, cy + arm], fill=WHITE, width=lw_arm)


def draw_import(draw, w):
    """Royal blue document with emerald down arrow."""
    margin = w * 0.1

    # Document body (royal blue)
    doc_l = margin
    doc_t = w * 0.08
    doc_r = w - margin
    doc_b = w * 0.65
    r = w * 0.08
    filled_rounded_rect(draw, [doc_l, doc_t, doc_r, doc_b], int(r), ROYAL)

    # Folded corner effect (lighter triangle at top-right)
    fold = w * 0.18
    draw.polygon([doc_r - fold, doc_t, doc_r, doc_t, doc_r, doc_t + fold],
                 fill=(129, 120, 248))  # lighter variant

    # White text lines on document
    lw_line = max(1, int(w * 0.045))
    line_l = doc_l + w * 0.12
    line_r = doc_r - w * 0.18
    y1 = doc_t + w * 0.2
    y2 = doc_t + w * 0.34
    y3 = doc_t + w * 0.48
    for y in [y1, y2, y3]:
        draw.line([line_l, y, line_r, y], fill=WHITE, width=lw_line)

    # Down arrow (emerald) below document
    arrow_cx = w / 2
    arrow_top = doc_b + w * 0.02
    arrow_bottom = w - margin
    arrow_lw = max(2, int(w * 0.07))

    # Arrow body
    draw.line([arrow_cx, arrow_top, arrow_cx, arrow_bottom * 0.85],
              fill=EMERALD, width=arrow_lw)
    # Arrow head
    head_sz = w * 0.15
    hx, hy = arrow_cx, arrow_bottom * 0.88
    draw.line([hx - head_sz, hy - head_sz * 0.6, hx, hy], fill=EMERALD, width=arrow_lw)
    draw.line([hx + head_sz, hy - head_sz * 0.6, hx, hy], fill=EMERALD, width=arrow_lw)


def draw_lookup(draw, w):
    """Blue table cell with amber magnifying glass overlay."""
    margin = w * 0.08

    # Table cell (blue) - bottom area
    cell_l = margin
    cell_t = w * 0.4
    cell_r = w * 0.62
    cell_b = w - margin
    r = w * 0.08
    filled_rounded_rect(draw, [cell_l, cell_t, cell_r, cell_b], int(r), BLUE)

    # White grid lines on cell
    lw_grid = max(1, int(w * 0.04))
    # Horizontal divider
    mid_y = (cell_t + cell_b) / 2
    draw.line([cell_l + r * 0.5, mid_y, cell_r - r * 0.5, mid_y], fill=WHITE, width=lw_grid)
    # Vertical divider
    mid_x = (cell_l + cell_r) / 2
    draw.line([mid_x, cell_t + r * 0.5, mid_x, cell_b - r * 0.5], fill=WHITE, width=lw_grid)

    # Magnifying glass (amber)
    glass_cx = w * 0.54
    glass_cy = w * 0.33
    glass_r = w * 0.2

    # Lens circle
    lw_glass = max(2, int(w * 0.06))
    draw.ellipse([glass_cx - glass_r, glass_cy - glass_r,
                  glass_cx + glass_r, glass_cy + glass_r],
                 outline=AMBER, width=lw_glass)
    # Inner highlight
    inner_r = glass_r * 0.55
    draw.ellipse([glass_cx - inner_r, glass_cy - inner_r,
                  glass_cx + inner_r, glass_cy + inner_r],
                 fill=(253, 224, 71, 100))  # light amber semi-transparent

    # Handle
    angle = math.radians(42)
    hx1 = glass_cx + glass_r * 0.85 * math.cos(angle)
    hy1 = glass_cy + glass_r * 0.85 * math.sin(angle)
    h_len = w * 0.22
    h_lw = max(2, int(w * 0.07))
    hx2 = hx1 + h_len * math.cos(angle)
    hy2 = hy1 + h_len * math.sin(angle)
    draw.line([hx1, hy1, hx2, hy2], fill=AMBER, width=h_lw)


def draw_tools(draw, w):
    """Purple/indigo gear/settings icon."""
    margin = w * 0.08
    cx = w / 2
    cy = w / 2
    outer_r = w * 0.35
    inner_r = w * 0.18
    lw = max(2, int(w * 0.05))

    # Draw gear shape using circle + teeth
    n_teeth = 8
    tooth_h = w * 0.08
    tooth_w = w * 0.07

    # Main circle
    draw.ellipse([cx - outer_r, cy - outer_r, cx + outer_r, cy + outer_r],
                 fill=PURPLE)
    # Inner hole
    draw.ellipse([cx - inner_r, cy - inner_r, cx + inner_r, cy + inner_r],
                 fill=WHITE)
    # Inner dot
    dot_r = w * 0.06
    draw.ellipse([cx - dot_r, cy - dot_r, cx + dot_r, cy + dot_r],
                 fill=INDIGO)

    # Teeth around the circle
    for i in range(n_teeth):
        angle_deg = i * (360 / n_teeth) - 90  # start from top
        angle_rad = math.radians(angle_deg)
        tx = cx + outer_r * math.cos(angle_rad)
        ty = cy + outer_r * math.sin(angle_rad)
        # Draw tooth as small rectangle rotated
        tooth_l = tx - tooth_w / 2
        tooth_t = ty - tooth_h
        tooth_r_p = tx + tooth_w / 2
        tooth_b = ty + tooth_h

        # Simple approach: draw small ellipse at each tooth position
        draw.ellipse([tx - tooth_w, ty - tooth_h, tx + tooth_w, ty + tooth_h],
                     fill=PURPLE)


def draw_expand(draw, w):
    """Left column expanding into multiple columns on the right."""
    margin = w * 0.08

    # Left column (dark blue) - the key column
    lx = margin
    ly = margin
    lw_col = w * 0.22
    lh_col = w - 2 * margin
    r = w * 0.06
    filled_rounded_rect(draw, [lx, ly, lx + lw_col, ly + lh_col], int(r), DARK_BG)

    # Arrow from left column to right side
    arrow_cx = lx + lw_col + w * 0.1
    arrow_cy = w / 2
    arrow_len = w * 0.12
    lw_arrow = max(2, int(w * 0.05))
    draw.line([arrow_cx, arrow_cy, arrow_cx + arrow_len, arrow_cy], fill=BLUE, width=lw_arrow)
    # Arrow head
    draw.polygon([
        arrow_cx + arrow_len, arrow_cy,
        arrow_cx + arrow_len - w * 0.06, arrow_cy - w * 0.05,
        arrow_cx + arrow_len - w * 0.06, arrow_cy + w * 0.05,
    ], fill=BLUE)

    # Right side - 3 small columns spreading outward (emerald, teal, amber)
    colors = [EMERALD, TEAL, BLUE]
    start_x = w * 0.55
    col_w = w * 0.12
    spacing = w * 0.12

    for i, color in enumerate(colors):
        cx = start_x + i * spacing
        cy = w / 2
        h = (w * 0.35) + (i - 1) * w * 0.08  # varying heights for visual interest
        filled_rounded_rect(draw, [
            cx - col_w / 2, cy - h / 2,
            cx + col_w / 2, cy + h / 2
        ], int(w * 0.04), color)


def generate_icon(name, draw_func):
    """Generate icon at all sizes."""
    for size in SIZES:
        canvas_w = size * SCALE
        img = Image.new("RGBA", (canvas_w, canvas_w), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw_func(draw, canvas_w)
        img = img.resize((size, size), Image.LANCZOS)
        path = f"assets/{name}-{size}.png"
        img.save(path)
        print(f"  Saved {path} ({size}x{size})")


if __name__ == "__main__":
    icons = [
        ("concat", draw_concat),
        ("import", draw_import),
        ("lookup", draw_lookup),
        ("tools", draw_tools),
        ("expand", draw_expand),
    ]

    for name, func in icons:
        print(f"Generating {name}...")
        generate_icon(name, func)

    print("\nDone!")

"""Generate simple line-style ribbon icons for Office Add-in buttons."""
from PIL import Image, ImageDraw
import math

SIZES = [16, 32, 80]
COLOR = (51, 51, 51)  # #333333
SCALE = 4  # supersampling factor for anti-aliasing
LINE_WIDTH = 1.8  # base line width at native size

def draw_rounded_rect(draw, xy, radius, fill=None, outline=None, width=1):
    """Draw a rounded rectangle."""
    x1, y1, x2, y2 = xy
    r = radius
    # Corner circles
    draw.arc([x1, y1, x1 + 2 * r, y1 + 2 * r], 180, 270, fill=outline, width=width)
    draw.arc([x2 - 2 * r, y1, x2, y1 + 2 * r], 270, 360, fill=outline, width=width)
    draw.arc([x1, y2 - 2 * r, x1 + 2 * r, y2], 90, 180, fill=outline, width=width)
    draw.arc([x2 - 2 * r, y2 - 2 * r, x2, y2], 0, 90, fill=outline, width=width)
    # Connecting lines
    draw.line([x1 + r, y1, x2 - r, y1], fill=outline, width=width)
    draw.line([x1 + r, y2, x2 - r, y2], fill=outline, width=width)
    draw.line([x1, y1 + r, x1, y2 - r], fill=outline, width=width)
    draw.line([x2, y1 + r, x2, y2 - r], fill=outline, width=width)
    if fill:
        draw.rectangle([x1 + width, y1 + width, x2 - width, y2 - width], fill=fill)


def draw_grid_cell(draw, x1, y1, x2, y2, lw):
    """Draw a table/cell icon - rectangle with horizontal and vertical lines."""
    # Outer rectangle
    draw.rectangle([x1, y1, x2, y2], outline=COLOR, width=lw)
    # Horizontal divider
    mid_y = (y1 + y2) // 2
    draw.line([x1, mid_y, x2, mid_y], fill=COLOR, width=lw)
    # Vertical divider
    mid_x = (x1 + x2) // 2
    draw.line([mid_x, y1, mid_x, y2], fill=COLOR, width=lw)


def draw_icon_concat(draw, w, lw):
    """Two cells with a plus sign between them."""
    margin = w * 0.12
    cell_w = (w - 2 * margin) * 0.32
    gap = w * 0.1
    cell_h = w * 0.7
    top = (w - cell_h) / 2

    # Left cell
    x1 = margin
    draw_grid_cell(draw, x1, top, x1 + cell_w, top + cell_h, max(1, int(lw)))
    # Right cell
    x2 = w - margin - cell_w
    draw_grid_cell(draw, x2, top, x2 + cell_w, top + cell_h, max(1, int(lw)))
    # Plus sign in the middle
    cx = w / 2
    cy = w / 2
    arm = cell_w * 0.45
    draw.line([cx - arm, cy, cx + arm, cy], fill=COLOR, width=max(1, int(lw * 1.4)))
    draw.line([cx, cy - arm, cx, cy + arm], fill=COLOR, width=max(1, int(lw * 1.4)))


def draw_icon_import(draw, w, lw):
    """Document/file with a down arrow."""
    margin = w * 0.12
    doc_left = w * 0.2
    doc_right = w * 0.78
    doc_top = w * 0.1
    doc_bottom = w * 0.75

    # Document body (rounded top-right, folded corner look)
    radius = w * 0.08
    draw_rounded_rect(draw, [doc_left, doc_top, doc_right, doc_bottom], int(radius),
                      outline=COLOR, width=max(1, int(lw)))

    # Down arrow below the document
    arrow_top = doc_bottom + w * 0.04
    arrow_bottom = w * 0.92
    arrow_mid_x = (doc_left + doc_right) / 2
    arrow_head = w * 0.12

    # Arrow shaft
    draw.line([arrow_mid_x, arrow_top, arrow_mid_x, arrow_bottom - arrow_head],
              fill=COLOR, width=max(1, int(lw * 1.2)))
    # Arrow head
    ah = arrow_head
    draw.line([arrow_mid_x - ah, arrow_bottom - ah, arrow_mid_x, arrow_bottom],
              fill=COLOR, width=max(1, int(lw * 1.2)))
    draw.line([arrow_mid_x + ah, arrow_bottom - ah, arrow_mid_x, arrow_bottom],
              fill=COLOR, width=max(1, int(lw * 1.2)))

    # Horizontal lines inside document (text lines)
    line_y1 = doc_top + w * 0.18
    line_y2 = doc_top + w * 0.32
    for ly in [line_y1, line_y2]:
        draw.line([doc_left + w * 0.08, ly, doc_right - w * 0.08, ly],
                  fill=COLOR, width=max(1, int(lw * 0.7)))


def draw_icon_lookup(draw, w, lw):
    """Magnifying glass over a table cell."""
    margin = w * 0.1

    # Table/cell in the bottom-left
    cell_left = margin
    cell_top = w * 0.4
    cell_right = w * 0.58
    cell_bottom = w - margin
    draw_grid_cell(draw, cell_left, cell_top, cell_right, cell_bottom, max(1, int(lw)))

    # Magnifying glass in the top-right
    glass_cx = w * 0.55
    glass_cy = w * 0.35
    glass_r = w * 0.22

    # Circle (lens)
    draw.ellipse([glass_cx - glass_r, glass_cy - glass_r,
                  glass_cx + glass_r, glass_cy + glass_r],
                 outline=COLOR, width=max(1, int(lw)))

    # Handle
    angle = math.radians(45)
    handle_start_x = glass_cx + glass_r * math.cos(angle)
    handle_start_y = glass_cy + glass_r * math.sin(angle)
    handle_len = w * 0.22
    handle_end_x = handle_start_x + handle_len * math.cos(angle)
    handle_end_y = handle_start_y + handle_len * math.sin(angle)
    draw.line([handle_start_x, handle_start_y, handle_end_x, handle_end_y],
              fill=COLOR, width=max(1, int(lw * 1.3)))


def draw_icon_tools(draw, w, lw):
    """Wrench icon."""
    margin = w * 0.13

    # Wrench handle - diagonal line
    start_x = w * 0.2
    start_y = w * 0.75
    end_x = w * 0.65
    end_y = w * 0.25
    handle_w = max(1, int(lw * 1.3))

    # Draw handle with some thickness (two parallel-ish lines)
    offset = w * 0.04
    draw.line([start_x - offset, start_y + offset, end_x - offset, end_y + offset],
              fill=COLOR, width=handle_w)
    draw.line([start_x + offset, start_y - offset, end_x + offset, end_y - offset],
              fill=COLOR, width=handle_w)

    # Wrench head (U-shape at the end)
    head_cx = end_x
    head_cy = end_y
    head_r = w * 0.16

    # Draw the U-shaped head
    # Left jaw
    jaw_w = max(1, int(lw * 1.0))
    jaw_len = w * 0.2
    jl_x1 = head_cx - head_r
    jl_x2 = head_cx - head_r
    jl_y1 = head_cy - w * 0.06
    jl_y2 = head_cy - jaw_len
    draw.line([jl_x1, jl_y1, jl_x2, jl_y2], fill=COLOR, width=jaw_w)

    # Right jaw
    jr_x1 = head_cx + head_r
    jr_x2 = head_cx + head_r
    draw.line([jr_x1, jl_y1, jr_x2, jl_y2], fill=COLOR, width=jaw_w)

    # Top arc connecting the jaws
    draw.arc([head_cx - head_r, head_cy - jaw_len - head_r,
              head_cx + head_r, head_cy - jaw_len + head_r],
             0, 180, fill=COLOR, width=jaw_w)


def generate_icon(name, draw_func):
    """Generate icon at all sizes."""
    for size in SIZES:
        canvas_w = size * SCALE
        lw = LINE_WIDTH * SCALE * (size / 16.0)  # scale line width with icon size

        img = Image.new("RGBA", (canvas_w, canvas_w), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw_func(draw, canvas_w, lw)

        # Downscale with anti-aliasing
        img = img.resize((size, size), Image.LANCZOS)
        path = f"assets/{name}-{size}.png"
        img.save(path)
        print(f"  Saved {path} ({size}x{size})")


if __name__ == "__main__":
    icons = [
        ("concat", draw_icon_concat),
        ("import", draw_icon_import),
        ("lookup", draw_icon_lookup),
        ("tools", draw_icon_tools),
    ]

    for name, func in icons:
        print(f"Generating {name} icons...")
        generate_icon(name, func)

    print("\nDone! All icons generated.")

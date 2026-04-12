"""
Generate proxlsx-skill cover image (2560x1280).
Deep green silk-like gradient, bold monospace title, muted subtitle, rounded corners.
Greenish palette to match the xlsx file-type color convention.
"""

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

W, H = 2560, 1280
CORNER_RADIUS = 80

# ── 1. Base canvas (deep forest) ─────────────────────────────────────────────
base = Image.new("RGBA", (W, H), (8, 16, 12, 255))  # #08100c — almost black-green

# ── 2. Color blobs ───────────────────────────────────────────────────────────
def make_blob(size, color_rgba, cx, cy, rx, ry):
    layer = Image.new("RGBA", size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    draw.ellipse([cx - rx, cy - ry, cx + rx, cy + ry], fill=color_rgba)
    return layer

# Green palette inspired by Excel's native file-type green + Macabacus link green
blobs = [
    # (r, g, b, alpha,  cx,    cy,   rx,   ry,  blur)
    ( 16, 124,  54, 200,   700,  600, 700, 500, 120),   # deep xlsx green, centered-left
    ( 34, 197, 100, 150,  1900,  200, 600, 400, 100),   # bright emerald, top-right
    (  6,  60,  30, 180,  1280, 1150, 900, 350,  90),   # dark pine, bottom
    ( 52, 168,  83, 130,  1400,  640, 500, 350,  80),   # mid Google-sheets green, center
    (112, 173,  71,  90,   300, 1050, 600, 400,  95),   # Macabacus link green, bottom-left
]

canvas = base.copy()
for r, g, b, a, cx, cy, rx, ry, blur in blobs:
    blob = make_blob((W, H), (r, g, b, a), cx, cy, rx, ry)
    blob = blob.filter(ImageFilter.GaussianBlur(radius=blur))
    canvas = Image.alpha_composite(canvas, blob)

canvas = canvas.filter(ImageFilter.GaussianBlur(radius=8))

# ── 3. Film grain ────────────────────────────────────────────────────────────
rng = np.random.default_rng(42)
noise = rng.integers(0, 255, (H, W), dtype=np.uint8)
grain_alpha = (noise * 0.22).astype(np.uint8)
grain_layer = np.stack([noise, noise, noise, grain_alpha], axis=-1).astype(np.uint8)
grain_img = Image.fromarray(grain_layer, "RGBA")
canvas = Image.alpha_composite(canvas, grain_img)

# ── 4. Fonts ─────────────────────────────────────────────────────────────────
TITLE_SIZE = 220
SUBTITLE_SIZE = 70

try:
    font_title = ImageFont.truetype("/System/Library/Fonts/Menlo.ttc", TITLE_SIZE, index=1)
    font_subtitle = ImageFont.truetype("/System/Library/Fonts/Menlo.ttc", SUBTITLE_SIZE, index=0)
except Exception:
    font_title = ImageFont.truetype("/System/Library/Fonts/Supplemental/Courier New Bold.ttf", TITLE_SIZE)
    font_subtitle = ImageFont.truetype("/System/Library/Fonts/Supplemental/Courier New Bold.ttf", SUBTITLE_SIZE)

TITLE_TEXT = "proxlsx-skill"
SUBTITLE_TEXT = "Pro-grade Excel. Zero AI-slop. Four themes, one builder."

# ── 5. Measure text ──────────────────────────────────────────────────────────
tmp = Image.new("RGBA", (W, H), (0, 0, 0, 0))
d = ImageDraw.Draw(tmp)

t_bbox = d.textbbox((0, 0), TITLE_TEXT, font=font_title)
t_w = t_bbox[2] - t_bbox[0]
t_h = t_bbox[3] - t_bbox[1]

s_bbox = d.textbbox((0, 0), SUBTITLE_TEXT, font=font_subtitle)
s_w = s_bbox[2] - s_bbox[0]
s_h = s_bbox[3] - s_bbox[1]

GAP = 48
total_h = t_h + GAP + s_h
block_top = (H - total_h) // 2 - 30

title_x = (W - t_w) // 2 - t_bbox[0]
title_y = block_top - t_bbox[1]

subtitle_x = (W - s_w) // 2 - s_bbox[0]
subtitle_y = title_y + t_h + GAP

# ── 6. Title with green glow ─────────────────────────────────────────────────
def draw_text_layer(text, x, y, font, color):
    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(layer).text((x, y), text, font=font, fill=color)
    return layer

glow_specs = [
    (( 80, 255, 140,  60), 18),   # outer glow, wide, dim green
    ((120, 255, 180,  90),  9),   # mid glow
    ((180, 255, 210, 120),  4),   # tight glow
]
for color, blur_r in glow_specs:
    glow = draw_text_layer(TITLE_TEXT, title_x, title_y, font_title, color)
    glow = glow.filter(ImageFilter.GaussianBlur(radius=blur_r))
    canvas = Image.alpha_composite(canvas, glow)

title_layer = draw_text_layer(TITLE_TEXT, title_x, title_y, font_title, (255, 255, 255, 245))
canvas = Image.alpha_composite(canvas, title_layer)

# ── 7. Subtitle ──────────────────────────────────────────────────────────────
subtitle_layer = draw_text_layer(
    SUBTITLE_TEXT, subtitle_x, subtitle_y, font_subtitle, (180, 220, 190, 210)
)
canvas = Image.alpha_composite(canvas, subtitle_layer)

# ── 8. Rounded corner mask ───────────────────────────────────────────────────
mask = Image.new("L", (W, H), 0)
ImageDraw.Draw(mask).rounded_rectangle([(0, 0), (W - 1, H - 1)], radius=CORNER_RADIUS, fill=255)
canvas.putalpha(mask)

canvas = canvas.filter(ImageFilter.GaussianBlur(radius=1))

# ── 9. Save ──────────────────────────────────────────────────────────────────
import os
out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cover.png")
canvas.save(out_path, "PNG", dpi=(400, 400))
print(f"Saved: {out_path}")
print(f"Size: {canvas.size}")

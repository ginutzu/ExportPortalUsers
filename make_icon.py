#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate app.ico - a simple 'users export' glyph, no external assets."""
from PIL import Image, ImageDraw

BASE = 256
BLUE = (0, 120, 200, 255)
DARK = (0, 90, 158, 255)
WHITE = (255, 255, 255, 255)
GREEN = (46, 175, 80, 255)


def rounded(draw, box, r, fill):
    draw.rounded_rectangle(box, radius=r, fill=fill)


img = Image.new("RGBA", (BASE, BASE), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# Document / card background
rounded(d, (36, 24, 220, 232), 24, DARK)
rounded(d, (30, 18, 214, 226), 24, BLUE)

# A person head + shoulders (the "user")
cx, cy = 78, 92
d.ellipse((cx - 22, cy - 22, cx + 22, cy + 22), fill=WHITE)          # head
d.pieslice((cx - 34, cy + 10, cx + 34, cy + 74), 180, 360, fill=WHITE)  # shoulders

# List rows (the "export list")
for i, y in enumerate((78, 118, 158)):
    d.rounded_rectangle((124, y, 196, y + 16), radius=8, fill=WHITE)

# Little green "export/check" badge
bx, by = 168, 176
d.ellipse((bx - 34, by - 34, bx + 34, by + 34), fill=GREEN)
d.line((bx - 16, by, bx - 4, by + 14), fill=WHITE, width=8)
d.line((bx - 4, by + 14, bx + 18, by - 14), fill=WHITE, width=8)

sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
img.save("app.ico", sizes=sizes)
print("app.ico written")

from PIL import Image, ImageDraw, ImageChops, ImageFilter
from render import render
from icons_v1 import ICONS
import math

M = 1024
def vgrad(w, h, top, bot):
    g = Image.new('RGBA', (w, h))
    px = g.load()
    for y in range(h):
        t = y / (h - 1)
        c = tuple(int(top[i] + (bot[i] - top[i]) * t) for i in range(3)) + (255,)
        for x in range(w): px[x, y] = c
    return g

def master(simple=False):
    img = Image.new('RGBA', (M, M), (0, 0, 0, 0))
    # 深藍圓角底
    tile = vgrad(M, M, (30, 40, 92), (13, 17, 44))
    mask = Image.new('L', (M, M), 0)
    ImageDraw.Draw(mask).rounded_rectangle([40, 40, M - 40, M - 40], radius=210, fill=255)
    img.paste(tile, (0, 0), mask)
    if not simple:
        # 內描邊
        ring = Image.new('L', (M, M), 0)
        d = ImageDraw.Draw(ring)
        d.rounded_rectangle([40, 40, M - 40, M - 40], radius=210, outline=255, width=6)
        edge = Image.new('RGBA', (M, M), (70, 100, 200, 120)); img.paste(edge, (0, 0), ring)
        # 背後光暈
        glow = Image.new('RGBA', (M, M), (0, 0, 0, 0))
        ImageDraw.Draw(glow).ellipse([200, 180, 800, 780], fill=(70, 150, 255, 120))
        glow = glow.filter(ImageFilter.GaussianBlur(90))
        img = Image.alpha_composite(img, ImageChops.multiply(glow, Image.merge('RGBA', (mask, mask, mask, mask))) if False else glow)
        # 光暈限制在圓角底內
        clip = Image.new('RGBA', (M, M), (0, 0, 0, 0)); clip.paste(img, (0, 0), mask); img = clip
    # 資料夾（亮藍漸層）
    fm = render(ICONS['folder'], 700, (255, 255, 255))   # 只取 alpha
    fa = fm.split()[3]
    fg = vgrad(700, 700, (126, 198, 255), (52, 144, 240))
    folder = Image.new('RGBA', (700, 700), (0, 0, 0, 0)); folder.paste(fg, (0, 0), fa)
    ox, oy = 118, 150
    img.alpha_composite(folder, (ox, oy))
    # 放大鏡：深藍描邊隔開 + 白環 + 淺藍鏡片
    cx, cy, R = 640, 640, 150 if not simple else 165
    lay = Image.new('RGBA', (M, M), (0, 0, 0, 0)); d = ImageDraw.Draw(lay)
    ang = math.radians(45); hl = 190
    hx0, hy0 = cx + (R) * math.cos(ang), cy + (R) * math.sin(ang)
    hx1, hy1 = cx + (R + hl) * math.cos(ang), cy + (R + hl) * math.sin(ang)
    def thick(draw, w, col, extra=0):
        draw.ellipse([cx - R - extra, cy - R - extra, cx + R + extra, cy + R + extra], fill=col)
        draw.line([(hx0, hy0), (hx1, hy1)], fill=col, width=int(w + 2 * extra))
        for (x, y) in ((hx1, hy1),):
            r = (w + 2 * extra) / 2; draw.ellipse([x - r, y - r, x + r, y + r], fill=col)
    # 外側深藍隔離帶
    thick(d, 74, (14, 19, 48, 255), extra=26)
    # 白色握把與環
    thick(d, 74, (255, 255, 255, 255))
    d.ellipse([cx - R + 40, cy - R + 40, cx + R - 40, cy + R - 40], fill=(62, 150, 242, 255))
    if not simple:
        d.ellipse([cx - R + 40, cy - R + 40, cx + R - 40, cy + R - 40], fill=None)
    img.alpha_composite(lay)
    return img

def build(sizes):
    out = {}
    big = master(False); small = master(True)
    for s in sizes:
        src = small if s <= 32 else big
        out[s] = src.resize((s, s), Image.LANCZOS)
    return out

if __name__ == '__main__':
    out = build([16, 24, 32, 48, 64, 128, 256])
    for s, im in out.items(): im.save(f'app_{s}.png')
    sheet = Image.new('RGB', (256 + 128 + 64 + 48 + 32 + 24 + 16 + 7 * 14, 270), (232, 232, 235))
    x = 7
    for s in [256, 128, 64, 48, 32, 24, 16]:
        sheet.paste(out[s], (x, 7), out[s]); x += s + 14
    sheet2 = Image.new('RGB', sheet.size, (28, 28, 30)); x = 7
    for s in [256, 128, 64, 48, 32, 24, 16]:
        sheet2.paste(out[s], (x, 7), out[s]); x += s + 14
    c = Image.new('RGB', (sheet.width, 540)); c.paste(sheet, (0, 0)); c.paste(sheet2, (0, 270)); c.save('app_preview.png')

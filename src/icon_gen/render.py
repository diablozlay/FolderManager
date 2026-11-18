from PIL import Image, ImageDraw, ImageChops
import math
from icons import ICONS

SS = 8
def arcpts(cx, cy, r, a0, sw, n=96):
    return [(cx + r*math.cos(math.radians(a0+sw*i/n)), cy + r*math.sin(math.radians(a0+sw*i/n))) for i in range(n+1)]

def render(ops, px, color, ss=SS):
    S = px*ss; u = S/32.0
    fill = Image.new('L', (S, S), 0)
    for op in ops:
        t = op.split(); k = t[0]; v = [float(x) for x in t[1:]]
        m = Image.new('L', (S, S), 0); d = ImageDraw.Draw(m)
        if k == 'c':
            cx, cy, r = v; d.ellipse([(cx-r)*u, (cy-r)*u, (cx+r)*u, (cy+r)*u], fill=255)
        elif k == 'r':
            x, y, w, h, rad = v; d.rounded_rectangle([x*u, y*u, (x+w)*u, (y+h)*u], radius=rad*u, fill=255)
        elif k == 'p':
            d.polygon([(v[i]*u, v[i+1]*u) for i in range(0, len(v), 2)], fill=255)
        elif k in 'la':
            continue
        fill = ImageChops.difference(fill, m)   # even-odd (XOR)
    # 描邊
    st = Image.new('L', (S, S), 0); d = ImageDraw.Draw(st)
    for op in ops:
        t = op.split(); k = t[0]; v = [float(x) for x in t[1:]]
        if k == 'l':
            x1, y1, x2, y2, w = v; W = w*u
            d.line([(x1*u, y1*u), (x2*u, y2*u)], fill=255, width=int(round(W)))
            for (x, y) in ((x1, y1), (x2, y2)):
                d.ellipse([x*u-W/2, y*u-W/2, x*u+W/2, y*u+W/2], fill=255)
        elif k == 'a':
            cx, cy, r, a0, sw, w = v; W = w*u
            pts = [(x*u, y*u) for x, y in arcpts(cx, cy, r, a0, sw)]
            d.line(pts, fill=255, width=int(round(W)), joint='curve')
            for (x, y) in (pts[0], pts[-1]):
                d.ellipse([x-W/2, y-W/2, x+W/2, y+W/2], fill=255)
    mask = ImageChops.lighter(fill, st).resize((px, px), Image.LANCZOS)
    out = Image.new('RGBA', (px, px), color + (0,))
    out.putalpha(mask)
    return out

if __name__ == '__main__':
    BLUE = (92, 176, 255)
    names = list(ICONS)
    tile = 120; pad = 16
    sheet = Image.new('RGB', (len(names)*tile, tile*2), (28, 28, 30))
    for i, n in enumerate(names):
        for row, (bg, col) in enumerate([((38, 38, 42), BLUE), ((196, 62, 62), (255, 255, 255))]):
            im = render(ICONS[n], 88, col)
            base = Image.new('RGBA', (tile, tile), bg + (255,))
            base.alpha_composite(im, (16, 16))
            sheet.paste(base.convert('RGB'), (i*tile, row*tile))
    sheet.save('preview.png')
    print(names)

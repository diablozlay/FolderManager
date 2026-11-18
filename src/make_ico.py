# 由 圖示\app_*.png 產生 圖示\app.ico（16~128 用 BMP/DIB，256 用 PNG）。路徑相對於本檔，可在任何機器/CI 執行。
import io, os, struct
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(HERE, '..', '圖示')
SIZES = [16, 24, 32, 48, 64, 128, 256]

def dib(img):
    w, h = img.size; px = img.load()
    hdr = struct.pack('<IiiHHIIiiII', 40, w, h * 2, 1, 32, 0, w * h * 4, 0, 0, 0, 0)
    xor = b''.join(b''.join(bytes((px[x, y][2], px[x, y][1], px[x, y][0], px[x, y][3])) for x in range(w)) for y in range(h - 1, -1, -1))
    mask = bytearray()
    for y in range(h - 1, -1, -1):
        bits = bytearray((w + 7) // 8)
        for x in range(w):
            if px[x, y][3] == 0: bits[x >> 3] |= 0x80 >> (x & 7)
        while len(bits) % 4: bits.append(0)
        mask += bits
    return hdr + xor + bytes(mask)

frames = []
for s in SIZES:
    img = Image.open(os.path.join(D, f'app_{s}.png')).convert('RGBA')
    if img.size != (s, s): img = img.resize((s, s), Image.LANCZOS)
    px = img.load()
    for y in range(s):
        for x in range(s):
            if px[x, y][3] < 3: px[x, y] = (0, 0, 0, 0)
    if s >= 256:
        b = io.BytesIO(); img.save(b, 'PNG', optimize=True); data = b.getvalue()
    else:
        data = dib(img)
    frames.append((s, data))

head = struct.pack('<HHH', 0, 1, len(frames)); off = 6 + 16 * len(frames); ent = b''; blobs = b''
for s, d in frames:
    w = 0 if s >= 256 else s
    ent += struct.pack('<BBBBHHII', w, w, 0, 0, 1, 32, len(d), off); off += len(d); blobs += d
out = os.path.join(D, 'app.ico')
open(out, 'wb').write(head + ent + blobs)
print('ICO build success, size:', os.path.getsize(out), 'bytes')

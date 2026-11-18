# 原位修補 exe 內 RT_VERSION 的 VS_FIXEDFILEINFO：File/Product Version 1.3.0.0 -> 1.3.3.0
import struct, sys

OLD_F = struct.pack('<II', (1 << 16) | 3, 0)          # dwFileVersionMS, dwFileVersionLS
NEW_F = struct.pack('<II', (1 << 16) | 3, (3 << 16) | 0)

SIG = 0xFEEF04BD

def patch(path):
    b = bytearray(open(path, 'rb').read())
    hits = 0
    pos = 0
    while True:
        i = b.find(struct.pack('<I', SIG), pos)
        if i < 0:
            break
        sig, struc, fms, fls, pms, pls = struct.unpack_from('<6I', b, i)
        cur = (fms, fls, pms, pls)
        if (fms, fls) == (0x00010003, 0x00000000):
            struct.pack_into('<I', b, i + 8, (1 << 16) | 3)      # FileVersionMS
            struct.pack_into('<I', b, i + 12, (3 << 16) | 0)     # FileVersionLS
            hits += 1
        if (pms, pls) == (0x00010003, 0x00000000):
            struct.pack_into('<I', b, i + 16, (1 << 16) | 3)     # ProductVersionMS
            struct.pack_into('<I', b, i + 20, (3 << 16) | 0)     # ProductVersionLS
            hits += 1
        print('  VS_FIXEDFILEINFO @%d  before=(%08x,%08x,%08x,%08x)' % ((i,) + cur))
        pos = i + 4
    # StringFileInfo 內的 UTF-16LE 字串（FileVersion / ProductVersion / Assembly Version）
    old16 = '1.3.0.0'.encode('utf-16le')
    new16 = '1.3.3.0'.encode('utf-16le')
    n16 = b.count(old16)
    b = b.replace(old16, new16)

    open(path, 'wb').write(bytes(b))
    print('patched fields:', hits, ' utf16 strings:', n16, '->', path)

if __name__ == '__main__':
    patch(sys.argv[1])

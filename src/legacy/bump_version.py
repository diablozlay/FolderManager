# 原位修改 exe 內的版本號（無 csc 環境時使用）。必須在簽章「之前」執行（修改會使簽章失效）。
# 用法：python bump_version.py <exe> 1.3.3.0 1.3.4.0
import struct, sys

def bump(path, old, new):
    o = tuple(int(x) for x in old.split('.')); n = tuple(int(x) for x in new.split('.'))
    assert len(o) == len(n) == 4 and all(0 <= v < 65536 for v in o + n)
    b = bytearray(open(path, 'rb').read())
    # 1) 版本資源內的字串（UTF-16）：FileVersion / ProductVersion / Assembly Version
    so, sn = old.encode('utf-16le'), new.encode('utf-16le')
    assert len(so) == len(sn), 'version strings must keep the same length'
    cnt = 0; i = b.find(so)
    while i >= 0: b[i:i+len(so)] = sn; cnt += 1; i = b.find(so, i + 1)
    assert cnt >= 2, ('utf16 hits', cnt)
    # 2) AssemblyFileVersion 自訂屬性（UTF-8, #Blob）
    so8, sn8 = old.encode(), new.encode()
    c8 = 0; i = b.find(so8)
    while i >= 0: b[i:i+len(so8)] = sn8; c8 += 1; i = b.find(so8, i + 1)
    assert c8 == 1, ('utf8 hits', c8)
    # 3) VS_FIXEDFILEINFO：FileVersion / ProductVersion（MS,LS 各兩組）
    sig = struct.pack('<I', 0xFEEF04BD); i = b.find(sig); assert i >= 0
    ms, ls = (o[0] << 16) | o[1], (o[2] << 16) | o[3]
    nms, nls = (n[0] << 16) | n[1], (n[2] << 16) | n[3]
    fms, fls, pms, pls = struct.unpack_from('<4I', b, i + 8)
    assert (fms, fls, pms, pls) == (ms, ls, ms, ls), (fms, fls, pms, pls)
    struct.pack_into('<4I', b, i + 8, nms, nls, nms, nls)
    # 4) .NET Assembly 表列：HashAlgId(0x8004) + Major/Minor/Build/Revision
    row = struct.pack('<IHHHH', 0x8004, *o); r = b.find(row); assert r >= 0 and b.find(row, r + 1) < 0
    b[r:r+len(row)] = struct.pack('<IHHHH', 0x8004, *n)
    open(path, 'wb').write(bytes(b))
    print(f'bumped {old} -> {new}: utf16 x{cnt}, blob x{c8}, FIXEDFILEINFO, Assembly row')

if __name__ == '__main__':
    bump(*sys.argv[1:4])

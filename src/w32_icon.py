# 用 Win32 UpdateResource API 把圖示寫進 exe（由系統處理 PE 資源結構，最可靠）
import ctypes, ctypes.wintypes as wt, struct, sys

RT_ICON = 3
RT_GROUP_ICON = 14

kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
kernel32.BeginUpdateResourceW.argtypes = [wt.LPCWSTR, wt.BOOL]
kernel32.BeginUpdateResourceW.restype = wt.HANDLE
kernel32.UpdateResourceW.argtypes = [wt.HANDLE, wt.LPVOID, wt.LPVOID, wt.WORD, wt.LPVOID, wt.DWORD]
kernel32.UpdateResourceW.restype = wt.BOOL
kernel32.EndUpdateResourceW.argtypes = [wt.HANDLE, wt.BOOL]
kernel32.EndUpdateResourceW.restype = wt.BOOL


def parse_ico(path):
    b = open(path, 'rb').read()
    res, typ, cnt = struct.unpack_from('<HHH', b, 0)
    assert res == 0 and typ == 1
    out = []
    for i in range(cnt):
        o = 6 + 16 * i
        w, h, cc, rsv, planes, bpp, size, off = struct.unpack_from('<BBBBHHII', b, o)
        out.append(dict(w=w, h=h, cc=cc, planes=planes, bpp=bpp, data=b[off:off + size]))
    return out


def inject(exe_path, ico_path):
    ico = parse_ico(ico_path)
    print('ico frames:', [(i['w'] or 256, i['h'] or 256, i['bpp']) for i in ico])

    h = kernel32.BeginUpdateResourceW(exe_path, False)
    if not h:
        raise OSError(ctypes.get_last_error(), 'BeginUpdateResource failed')

    try:
        lang = 0  # 中性語言，Windows 會 fallback
        for idx, im in enumerate(ico, start=1):
            data = im['data']
            buf = ctypes.create_string_buffer(data, len(data))
            ok = kernel32.UpdateResourceW(h, ctypes.cast(RT_ICON, wt.LPVOID),
                                          ctypes.cast(idx, wt.LPVOID), lang,
                                          buf, len(data))
            if not ok:
                raise OSError(ctypes.get_last_error(), f'UpdateResource RT_ICON {idx} failed')

        # GRPICONDIR
        grp = bytearray(struct.pack('<HHH', 0, 1, len(ico)))
        for idx, im in enumerate(ico, start=1):
            w = 0 if im['w'] >= 256 else im['w']
            hh = 0 if im['h'] >= 256 else im['h']
            grp.extend(struct.pack('<BBBBHHIH', w, hh, im['cc'], 0, im['planes'], im['bpp'], len(im['data']), idx))
        gbuf = ctypes.create_string_buffer(bytes(grp), len(grp))
        ok = kernel32.UpdateResourceW(h, ctypes.cast(RT_GROUP_ICON, wt.LPVOID),
                                      ctypes.cast(1, wt.LPVOID), lang, gbuf, len(grp))
        if not ok:
            raise OSError(ctypes.get_last_error(), 'UpdateResource RT_GROUP_ICON failed')
    finally:
        if not kernel32.EndUpdateResourceW(h, False):
            raise OSError(ctypes.get_last_error(), 'EndUpdateResource failed')
    print('icon injected ->', exe_path)


if __name__ == '__main__':
    inject(sys.argv[1], sys.argv[2])

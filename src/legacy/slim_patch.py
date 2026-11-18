# 將 FileManager.ps1 精簡（去全行註解、去行尾空白、去空行）後，原位替換 exe 內嵌資源
# 條件：精簡後 <= 46138 bytes（原資源長度），metadata 不動
import struct, sys

SRC_EXE = r'C:/Users/asus/WorkBuddy/2026-11-17-21-47-01/_fm/檔案管理工具-v1.3.0/檔案管理工具.exe'
NEW_PS1 = r'C:/Users/asus/WorkBuddy/2026-11-17-21-47-01/_fm/FileManager.ps1'
OUT_EXE = r'C:/Users/asus/WorkBuddy/2026-11-17-21-47-01/_fm/檔案管理工具_v133.exe'
RES_OFF = 0x640          # 資源 blob 起點（4-byte 長度前綴）
OLD_LEN = 46138          # 原資源資料長度

# --- 1. 精簡腳本（無 here-string，去整行註解/行尾空白/空行/行首縮排皆安全）---
raw = open(NEW_PS1, 'rb').read().decode('utf-8-sig')
assert "@'" not in raw and "'@" not in raw and '@"' not in raw and '"@' not in raw, 'here-string found, slim unsafe'
kept = []
for line in raw.split('\n'):
    s = line.rstrip()
    t = s.lstrip()
    if not t:
        continue
    if t.startswith('#') and not t.startswith('#requires') and not t.startswith('#>'):
        continue
    kept.append(t)      # 去行首縮排：PowerShell 不靠縮排分層，可省約 4KB 資源空間
slim = ('﻿' + '\n'.join(kept)).encode('utf-8')
print('slim bytes:', len(slim))
assert len(slim) <= OLD_LEN, 'slim script too big'
assert slim.startswith(b'\xef\xbb\xbf#requires'), 'BOM/#requires broken'

# --- 2. 語法快速自檢（{ } 配平 + 關鍵符碼存在）---
txt = slim.decode('utf-8-sig')
assert txt.count('{') == txt.count('}'), 'brace mismatch'
for needle in ('Invoke-AbortAll', '$script:AbortAll', 'btnAbort', 'KeyPreview', 'lblCredit', '阿家_diabloz', 'New-Glyph', 'pnlProg'):
    assert needle in txt, 'missing ' + needle

# --- 3. 原位寫入 exe ---
b = bytearray(open(SRC_EXE, 'rb').read())
struct.pack_into('<I', b, RES_OFF, len(slim))          # 更新長度前綴
b[RES_OFF + 4 : RES_OFF + 4 + len(slim)] = slim        # 寫入新資源
tail = RES_OFF + 4 + len(slim)
old_end = RES_OFF + 4 + OLD_LEN
for i in range(tail, old_end):                          # 殘段清零
    b[i] = 0

# --- 4. 版本號 1.3.0.0 -> 1.3.3.0（等長替換）---
cnt_str = b.count(b'1.3.0.0')
print('version string occurrences:', cnt_str)
b = b.replace(b'1.3.0.0', b'1.3.3.0')
# AssemblyVersion 二進位（4 x uint16 = 1,3,0,0）
pat = struct.pack('<4H', 1, 3, 0, 0)
cnt_bin = b.count(pat)
print('assembly version binary occurrences:', cnt_bin)
if cnt_bin == 1:
    b = b.replace(pat, struct.pack('<4H', 1, 3, 3, 0))

open(OUT_EXE, 'wb').write(bytes(b))
print('wrote', OUT_EXE, len(b))

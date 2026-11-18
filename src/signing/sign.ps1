# 對 exe 做 Authenticode 簽章並驗證。需要 Windows SDK 的 signtool.exe（GitHub windows-latest 已內建）。
# 用法：
#   .\sign.ps1 -Thumbprint <憑證指紋>                      # 憑證在目前使用者的憑證存放區
#   .\sign.ps1 -Pfx C:\path\cert.pfx -Password '****'     # 使用 .pfx 檔
param(
    [string]$Exe = (Join-Path $PSScriptRoot '..\..\檔案管理工具.exe'),
    [string]$Pfx,
    [string]$Password,
    [string]$Thumbprint,
    [string]$TimestampUrl = 'http://timestamp.digicert.com'
)
$ErrorActionPreference = 'Stop'
if (-not $Pfx -and -not $Thumbprint) { throw '請指定 -Pfx 或 -Thumbprint' }
$st = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\signtool.exe" -ErrorAction SilentlyContinue |
      Sort-Object FullName -Descending | Select-Object -First 1
if (-not $st) { throw '找不到 signtool.exe，請安裝 Windows SDK（Signing Tools for Desktop Apps）' }
$a = @('sign', '/fd', 'SHA256', '/td', 'SHA256', '/tr', $TimestampUrl, '/d', '檔案管理工具 Folder Manager')
if ($Pfx) { $a += @('/f', $Pfx); if ($Password) { $a += @('/p', $Password) } } else { $a += @('/sha1', $Thumbprint) }
& $st.FullName @a $Exe
if ($LASTEXITCODE -ne 0) { throw "signtool sign 失敗（$LASTEXITCODE）" }
Get-AuthenticodeSignature $Exe | Format-List Status, StatusMessage, SignerCertificate
# 自簽憑證在未信任前 Status 會是 UnknownError / NotTrusted，屬正常；用 trust-cert.ps1 信任後即為 Valid。

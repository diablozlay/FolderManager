# 建立自簽程式碼簽章憑證（內部使用 / 測試）。私鑰只會留在你的電腦，不要提交到 Git。
# 對外公開發佈請改用商業憑證（OV/EV）或 Azure Trusted Signing，詳見 README「程式碼簽章」。
param(
    [string]$Subject = 'CN=Folder Manager (阿家_diabloz)',
    [string]$OutDir  = $PSScriptRoot,
    [int]$Years      = 5
)
$ErrorActionPreference = 'Stop'
$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $Subject `
    -KeyAlgorithm RSA -KeyLength 3072 -HashAlgorithm SHA256 `
    -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears($Years)
$cer = Join-Path $OutDir 'FolderManager-codesign.cer'
Export-Certificate -Cert $cert -FilePath $cer | Out-Null
$pwd = Read-Host '設定 .pfx 匯出密碼（留空則不匯出 .pfx）' -AsSecureString
if ($pwd.Length -gt 0) {
    $pfx = Join-Path $OutDir 'FolderManager-codesign.pfx'
    Export-PfxCertificate -Cert $cert -FilePath $pfx -Password $pwd | Out-Null
    Write-Host "已匯出：$pfx（請妥善保管；已被 .gitignore 排除）"
}
Write-Host "已匯出公開憑證：$cer"
Write-Host "憑證指紋（Thumbprint）：$($cert.Thumbprint)"
Write-Host "簽章：.\sign.ps1 -Thumbprint $($cert.Thumbprint)"

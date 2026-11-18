# 讓這台電腦信任自簽憑證（內部部署用）：匯入「受信任的根授權單位」與「受信任的發行者」。
# 匯入根憑證會跳出 Windows 安全性確認視窗，請按「是」。企業環境建議改用群組原則（GPO）派送。
param([string]$Cer = (Join-Path $PSScriptRoot 'FolderManager-codesign.cer'))
$ErrorActionPreference = 'Stop'
Import-Certificate -FilePath $Cer -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
Import-Certificate -FilePath $Cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null
Write-Host '已信任：' $Cer
Write-Host '提醒：這只讓「這台電腦」信任此憑證，不會消除其他電腦上的未知發行者警告。'

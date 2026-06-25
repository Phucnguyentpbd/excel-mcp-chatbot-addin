$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$CertDir = Join-Path $env:USERPROFILE ".office-addin-dev-certs"
$BackupRoot = Join-Path $ProjectRoot "backups"
$Stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$BackupDir = Join-Path $BackupRoot "office-addin-dev-certs-$Stamp"

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
powershell -ExecutionPolicy Bypass -File (Join-Path $ProjectRoot "stop-excel-chatbot.ps1") | Out-Host

New-Item -ItemType Directory -Force -Path $BackupRoot | Out-Null
if (Test-Path $CertDir) {
    Copy-Item -LiteralPath $CertDir -Destination $BackupDir -Recurse -Force
    Remove-Item -LiteralPath $CertDir -Recurse -Force
}

node -e "const { generateCertificates } = require('office-addin-dev-certs'); generateCertificates(undefined, undefined, undefined, 365).catch((e) => { console.error(e); process.exit(1); });"

$CaPath = Join-Path $CertDir "ca.crt"
if (-not (Test-Path $CaPath)) {
    throw "CA certificate was not generated: $CaPath"
}

Import-Certificate -FilePath $CaPath -CertStoreLocation Cert:\CurrentUser\Root | Out-Null

Write-Host "Renewed Office add-in localhost certificate."
Write-Host "Backup: $BackupDir"
Write-Host "CurrentUser Root certificate:"
Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -like "*Developer CA for Microsoft Office Add-ins*" } | Select-Object Subject, Thumbprint, NotBefore, NotAfter

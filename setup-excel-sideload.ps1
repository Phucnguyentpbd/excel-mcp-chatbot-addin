$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$CatalogPath = Join-Path $ProjectRoot "catalog"
$ManifestPath = Join-Path $ProjectRoot "manifest.xml"
$CatalogManifestPath = Join-Path $CatalogPath "manifest.xml"
$CatalogId = "{a5e4b14b-d500-4a29-9175-60dbdad8988f}"
$RegistryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogId"

New-Item -ItemType Directory -Force -Path $CatalogPath | Out-Null
Copy-Item -LiteralPath $ManifestPath -Destination $CatalogManifestPath -Force

$ShareName = "ExcelMcpChatbotAddin"
$CatalogUrl = "\\$env:COMPUTERNAME\$ShareName"
$shareCreated = $false

try {
    $existingShare = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
    if (-not $existingShare) {
        New-SmbShare -Name $ShareName -Path $CatalogPath -ChangeAccess $env:USERNAME | Out-Null
        $shareCreated = $true
    }
} catch {
    $adminSharePath = "\\localhost\C$" + $CatalogPath.Substring(2)
    if (Test-Path (Join-Path $adminSharePath "manifest.xml")) {
        $CatalogUrl = $adminSharePath
    } else {
        throw
    }
}

New-Item -Path $RegistryPath -Force | Out-Null
New-ItemProperty -Path $RegistryPath -Name Id -Value $CatalogId -PropertyType String -Force | Out-Null
New-ItemProperty -Path $RegistryPath -Name Url -Value $CatalogUrl -PropertyType String -Force | Out-Null
New-ItemProperty -Path $RegistryPath -Name Flags -Value 1 -PropertyType DWord -Force | Out-Null

Write-Host "Excel sideload catalog configured."
Write-Host "Manifest: $CatalogManifestPath"
Write-Host "Catalog Url: $CatalogUrl"
Write-Host "Registry: $RegistryPath"
if (-not $shareCreated) {
    Write-Host "Used existing share or admin-share fallback."
}
Write-Host "Restart Excel, then open Home > Add-ins > Advanced > SHARED FOLDER."

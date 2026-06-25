$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$BackupRoot = Join-Path $ProjectRoot "backups"
$Stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$WefPath = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
$BackupPath = Join-Path $BackupRoot "Wef-cleanup-$Stamp"
$OfficeRoot = "HKCU:\Software\Microsoft\Office\16.0"
$ProviderRoot = Join-Path $OfficeRoot "WEF\Providers"
$TrustedCatalogRoot = Join-Path $OfficeRoot "WEF\TrustedCatalogs"
$NewCatalogId = "{a5e4b14b-d500-4a29-9175-60dbdad8988f}"
$OldCatalogIds = @(
    "{72435E44-3DDA-4A1D-A905-175F910922D6}",
    "{7c5a6d5a-5945-4bda-bda5-9b3b1ef2af58}"
)
$OldPathPatterns = @(
    "AGENT_AI\agent excel\subagent_excel\addin",
    "AGENT_AI\COPILOT EXCEL\excel_copilot\addin"
)
$OldManifestIds = @(
    "cd309a47-c990-48f8-b6ec-78ea0ba88a07",
    "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
)

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

New-Item -ItemType Directory -Force -Path $BackupRoot | Out-Null
if (Test-Path $WefPath) {
    Copy-Item -LiteralPath $WefPath -Destination $BackupPath -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $WefPath -Recurse -Force -ErrorAction SilentlyContinue
}

foreach ($oldCatalogId in $OldCatalogIds) {
    $oldCatalogPath = Join-Path $TrustedCatalogRoot $oldCatalogId
    if (Test-Path $oldCatalogPath) {
        Remove-Item -Path $oldCatalogPath -Recurse -Force
    }
}

if (Test-Path $ProviderRoot) {
    Get-ChildItem $ProviderRoot | ForEach-Object {
        $provider = Get-ItemProperty $_.PsPath
        foreach ($pattern in $OldPathPatterns) {
            if (($provider.UniqueId -as [string]) -like "*$pattern*") {
                Remove-Item -Path $_.PsPath -Recurse -Force
                break
            }
        }
    }
}

$customUiCache = Join-Path $OfficeRoot "Common\CustomUIValidationCache"
if (Test-Path $customUiCache) {
    $props = Get-ItemProperty $customUiCache
    foreach ($prop in $props.PSObject.Properties) {
        foreach ($oldId in $OldManifestIds) {
            if ($prop.Name -like "$oldId*") {
                Remove-ItemProperty -Path $customUiCache -Name $prop.Name -Force
            }
        }
    }
}

$mruRoot = Join-Path $OfficeRoot "Excel\Web Extension User MRU"
if (Test-Path $mruRoot) {
    Get-ChildItem $mruRoot -Recurse | Where-Object { $_.PSChildName -eq "File MRU" } | ForEach-Object {
        $props = Get-ItemProperty $_.PsPath
        foreach ($prop in $props.PSObject.Properties) {
            $value = $prop.Value -as [string]
            if (-not $value) {
                continue
            }
            $isOld = $false
            foreach ($oldId in $OldManifestIds) {
                if ($value -like "*$oldId*") {
                    $isOld = $true
                }
            }
            foreach ($pattern in $OldPathPatterns) {
                if ($value -like "*$pattern*") {
                    $isOld = $true
                }
            }
            if ($isOld) {
                Remove-ItemProperty -Path $_.PsPath -Name $prop.Name -Force
            }
        }
    }
}

Write-Host "Old Excel add-in traces cleaned."
Write-Host "WEF backup: $BackupPath"
Write-Host "Current trusted catalogs:"
Get-ChildItem $TrustedCatalogRoot -ErrorAction SilentlyContinue | ForEach-Object {
    Get-ItemProperty $_.PsPath | Select-Object PSChildName, Id, Url, Flags
}
Write-Host "Expected active catalog id: $NewCatalogId"

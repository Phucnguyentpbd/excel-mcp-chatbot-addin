$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogPath = Join-Path $ProjectRoot ".server-out.log"
$ErrorLogPath = Join-Path $ProjectRoot ".server-err.log"
$PidPath = Join-Path $ProjectRoot ".server.pid"
$BundledMcpPython = Join-Path $ProjectRoot "vendor\excel-mcp-server\.venv\Scripts\python.exe"
$BundledMcpSetup = Join-Path $ProjectRoot "setup-bundled-excel-mcp.ps1"

if (-not $env:OPENAI_API_KEY) {
    Write-Warning "OPENAI_API_KEY is not set. The pane will open, but chat requests will not run until you set it."
}

$env:EXCEL_CHATBOT_PROTOCOL = if ($env:EXCEL_CHATBOT_PROTOCOL) { $env:EXCEL_CHATBOT_PROTOCOL } else { "https" }
$BaseUrl = "$($env:EXCEL_CHATBOT_PROTOCOL)://localhost:3100"

$existing = Get-NetTCPConnection -LocalPort 3100 -State Listen -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "$BaseUrl is already listening."
    $existing | Select-Object LocalAddress, LocalPort, State, OwningProcess
    return
}

if (-not (Test-Path -LiteralPath $BundledMcpPython)) {
    Write-Host "Bundled Excel MCP runtime not found. Bootstrapping local Python environment..."
    & $BundledMcpSetup
}

$process = Start-Process node `
    -ArgumentList "server/app.mjs" `
    -WorkingDirectory $ProjectRoot `
    -WindowStyle Hidden `
    -RedirectStandardOutput $LogPath `
    -RedirectStandardError $ErrorLogPath `
    -PassThru

Set-Content -LiteralPath $PidPath -Value $process.Id -Encoding ascii
Start-Sleep -Seconds 3

$health = curl.exe -k -s "$BaseUrl/health"
Write-Host "Excel MCP Chatbot is running at $BaseUrl"
Write-Host "Manifest: $ProjectRoot\manifest.xml"
Write-Host "PID: $($process.Id)"
Write-Host "Health: $health"

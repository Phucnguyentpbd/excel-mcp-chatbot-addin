$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$PidPath = Join-Path $ProjectRoot ".server.pid"

if (Test-Path $PidPath) {
    $serverPid = [int](Get-Content -LiteralPath $PidPath -Raw)
    $process = Get-Process -Id $serverPid -ErrorAction SilentlyContinue
    if ($process) {
        Stop-Process -Id $serverPid -Force
        Write-Host "Stopped Excel MCP Chatbot server PID $serverPid."
    }
    Remove-Item -LiteralPath $PidPath -Force
} else {
    Write-Host "No .server.pid file found."
}

$listeners = Get-NetTCPConnection -LocalPort 3100 -State Listen -ErrorAction SilentlyContinue
if ($listeners) {
    Write-Warning "Port 3100 still has a listener. Check OwningProcess before stopping it manually."
    $listeners | Select-Object LocalAddress, LocalPort, State, OwningProcess
}

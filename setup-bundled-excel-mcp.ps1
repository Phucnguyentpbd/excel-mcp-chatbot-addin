$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$BundledMcpRoot = Join-Path $ProjectRoot "vendor\excel-mcp-server"
$VenvRoot = Join-Path $BundledMcpRoot ".venv"
$VenvPython = Join-Path $VenvRoot "Scripts\python.exe"

if (-not (Test-Path -LiteralPath $BundledMcpRoot)) {
    throw "Bundled excel-mcp-server source not found: $BundledMcpRoot"
}

if (-not (Test-Path -LiteralPath $VenvPython)) {
    Write-Host "Creating bundled Excel MCP virtual environment..."
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3.12 -m venv $VenvRoot
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        & python -m venv $VenvRoot
    } else {
        throw "Python launcher not found. Install Python 3.10+ first."
    }
}

Write-Host "Installing bundled excel-mcp-server dependencies..."
& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -e $BundledMcpRoot

Write-Host "Bundled Excel MCP is ready."
Write-Host "Python: $VenvPython"
& $VenvPython -m excel_mcp --help | Select-Object -First 3

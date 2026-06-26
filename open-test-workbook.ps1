$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$WorkbookPath = Join-Path $ProjectRoot "sample-workbook.xlsx"
$BundledPython = Join-Path $ProjectRoot "vendor\excel-mcp-server\.venv\Scripts\python.exe"

if (-not (Test-Path $WorkbookPath)) {
    if (-not (Test-Path $BundledPython)) {
        & (Join-Path $ProjectRoot "setup-bundled-excel-mcp.ps1")
    }

    $python = $BundledPython
    $env:EXCEL_CHATBOT_SAMPLE_WORKBOOK = $WorkbookPath
    $scriptPath = Join-Path $ProjectRoot ".create-sample-workbook.py"
    Set-Content -LiteralPath $scriptPath -Encoding ascii -Value @'
import os
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.title = "Sales"
rows = [
    ["Month", "Region", "Revenue", "Cost"],
    ["Jan", "North", 128000, 73000],
    ["Jan", "South", 96000, 61000],
    ["Feb", "North", 142000, 78000],
    ["Feb", "South", 105000, 65000],
    ["Mar", "North", 151000, 81000],
    ["Mar", "South", 119000, 69000],
]
for row in rows:
    ws.append(row)
wb.save(os.environ["EXCEL_CHATBOT_SAMPLE_WORKBOOK"])
'@
    & $python $scriptPath
    Remove-Item -LiteralPath $scriptPath -Force -ErrorAction SilentlyContinue
    Remove-Item Env:\EXCEL_CHATBOT_SAMPLE_WORKBOOK -ErrorAction SilentlyContinue
}

Start-Process "C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE" -ArgumentList "`"$WorkbookPath`""
Write-Host "Opened test workbook: $WorkbookPath"

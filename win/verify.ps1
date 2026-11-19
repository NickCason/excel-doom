# Opens the built workbook in Excel for Windows, runs one macro and writes the string it
# returns to a text file. With the defaults this is the engine self-check.
# Run from the repo root:  powershell -ExecutionPolicy Bypass -File win\verify.ps1
#   -Macro Host.Audit   paint-cost measurements instead of the self-check
#   -Maximize           maximise the Excel window first (what a player would normally have)
#   -Linger 3           keep the workbook open that many seconds after the macro returns
param([string]$Workbook = "out\doom.xlsm", [string]$Macro = "Host.Verify",
      [string]$Out = "verify_result.txt", [switch]$Maximize, [int]$Linger = 0)
$path = (Resolve-Path $Workbook).Path
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
try {
    if ($Maximize) { $xl.WindowState = -4137 }   # xlMaximized
    $wb = $xl.Workbooks.Open($path)
    $result = $xl.Run($Macro)
    Set-Content -Path $Out -Value $result -Encoding UTF8
    Write-Host "Result written to $Out"
    if ($Linger -gt 0) { Start-Sleep -Seconds $Linger }
    $wb.Close($false)
} finally {
    $xl.Quit()
}

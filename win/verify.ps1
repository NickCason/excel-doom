# UNTESTED (written on a Mac). Opens the built workbook in Excel for Windows, runs the
# engine self-check macro and writes its result to verify_result.txt.
# Run from the repo root:  powershell -ExecutionPolicy Bypass -File win\verify.ps1
param([string]$Workbook = "out\doom.xlsm", [string]$Macro = "Host.Verify")
$path = (Resolve-Path $Workbook).Path
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
try {
    $wb = $xl.Workbooks.Open($path)
    $result = $xl.Run($Macro)
    Set-Content -Path "verify_result.txt" -Value $result -Encoding UTF8
    Write-Host "Result written to verify_result.txt"
    $wb.Close($false)
} finally {
    $xl.Quit()
}

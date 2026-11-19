# Opens the built workbook in Excel for Windows, runs one macro and writes the string it
# returns to a text file. With the defaults this is the engine self-check.
# Run from the repo root:  powershell -ExecutionPolicy Bypass -File win\verify.ps1
#   -Macro Host.Audit   paint-cost measurements instead of the self-check
#   -Maximize           maximise the Excel window first (what a player would normally have)
#   -Linger 3           keep the workbook open that many seconds after the macro returns
#   -State file.txt     also write what the workbook looks like afterwards (tabs, checklist row)
#   -Dump out\frame_    also write the three picture dumps to out\frame_0.txt .. _2.txt
param([string]$Workbook = "out\doom.xlsm", [string]$Macro = "Host.Verify",
      [string]$Out = "verify_result.txt", [switch]$Maximize, [int]$Linger = 0, [string]$State = "", [string]$Dump = "")
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
    if ($Dump) {   # Host.Dump 0, 1, 2: engine frame, painter's record, real cell colours (see hostvba.py)
        foreach ($k in 0, 1, 2) { Set-Content -Path "$Dump$k.txt" -Value $xl.Run("Host.Dump", $k) -Encoding ASCII }
    }
    if ($State) {
        $lines = @("active=" + $wb.ActiveSheet.Name, "saved=" + $wb.Saved)
        foreach ($ws in $wb.Worksheets) { $lines += "tab=" + $ws.Name + "," + $ws.Visible }
        $c = $wb.Worksheets.Item("Checklist")
        $lines += "note=" + $c.Cells.Item(10, 4).Text
        $lines += "tick=" + [int][char]$c.Shapes.Item("Tick").TextFrame.Characters().Text
        Set-Content -Path $State -Value $lines -Encoding UTF8
    }
    $wb.Close($false)
} finally {
    $xl.Quit()
}

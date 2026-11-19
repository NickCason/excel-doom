# The path a first-time player takes, with a real mouse: open the workbook, click the splash,
# play a little, quit, click the picture to play again, then start a third game from the
# checkbox on the Checklist tab. Also checks that arrow keys over the picture do not start a
# game. Screenshots go to out\click_*.png.
# Run from the repo root:  powershell -ExecutionPolicy Bypass -File win\clicktest.ps1
param([string]$Workbook = "out\doom.xlsm")
Add-Type -AssemblyName System.Drawing
Add-Type -Namespace W -Name I -MemberDefinition @'
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
[DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
[DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
[DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
'@
[void][W.I]::SetProcessDPIAware()
$script:hwnd = [IntPtr]::Zero
function Test-Front { [W.I]::GetForegroundWindow() -eq $script:hwnd }
function Set-Front { if (-not (Test-Front)) { [void][W.I]::SetForegroundWindow($script:hwnd); Start-Sleep -Milliseconds 500 }; Test-Front }
function Send-Key([byte]$vk, [int]$ms = 100) {
    if (-not (Set-Front)) { "  key $vk skipped: Excel is not in front"; return }
    $ext = if ($vk -ge 0x25 -and $vk -le 0x28) { 1 } else { 0 }
    [W.I]::keybd_event($vk, 0, $ext, [UIntPtr]::Zero); Start-Sleep -Milliseconds $ms
    [W.I]::keybd_event($vk, 0, ($ext -bor 2), [UIntPtr]::Zero); Start-Sleep -Milliseconds 150
}
function Send-Click([int]$x, [int]$y) {
    if (-not (Set-Front)) { "  click skipped: Excel is not in front"; return }
    [void][W.I]::SetCursorPos($x, $y); Start-Sleep -Milliseconds 200
    [W.I]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 60
    [W.I]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero)
}
function Save-Shot($name) {
    $r = New-Object W.I+RECT; [void][W.I]::GetWindowRect($script:hwnd, [ref]$r)
    $l = [Math]::Max($r.L, 0); $t = [Math]::Max($r.T, 0); $w = $r.R - $l; $h = $r.B - $t
    $bmp = New-Object System.Drawing.Bitmap $w, $h; $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($l, $t, 0, 0, $bmp.Size); $bmp.Save((Join-Path $PWD "out\$name.png")); $g.Dispose(); $bmp.Dispose()
    "  shot out\$name.png (Excel in front: $(Test-Front))"
}
# screen position of a point on the active sheet, given in points from the sheet's top left
function Get-Screen($win, [double]$left, [double]$top) {
    $k = $win.Zoom / 100.0 * 96.0 / 72.0
    @([int]($win.PointsToScreenPixelsX(0) + $left * $k), [int]($win.PointsToScreenPixelsY(0) + $top * $k))
}
function Get-Idle($xl) { try { "$($xl.StatusBar)" -notmatch 'DOOM' } catch { $false } }   # a running game makes Excel refuse the call
function Get-Note($wb) { $wb.Worksheets.Item("Checklist").Cells.Item(10, 4).Text }
function Get-PictureCentre($xl, $wb) { $c = $wb.Worksheets.Item(1).Cells.Item(60, 160); Get-Screen $xl.ActiveWindow $c.Left $c.Top }

$path = (Resolve-Path $Workbook).Path
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$ok = $true
try {
    $xl.WindowState = -4137
    $wb = $xl.Workbooks.Open($path)
    $script:hwnd = [IntPtr]$xl.Hwnd
    $ws = $wb.Worksheets.Item(1); $tick = $wb.Worksheets.Item("Checklist").Shapes.Item("Tick")
    "opened on tab '$($wb.ActiveSheet.Name)'; tabs: $(($wb.Worksheets | ForEach-Object { $_.Name }) -join ', ')"
    "   zoom fitted by Workbook_Open: $($xl.ActiveWindow.Zoom) % | macro hint still under the frame: $([bool]$ws.Cells.Item(203, 2).Text)"
    if ($wb.ActiveSheet.Name -ne $ws.Name -or $ws.Cells.Item(203, 2).Text) { $ok = $false }
    [void](Set-Front); Start-Sleep -Milliseconds 800
    Save-Shot "click_0_splash"

    "1. click the splash"
    $p = Get-PictureCentre $xl $wb
    Send-Click $p[0] $p[1]
    Start-Sleep -Seconds 8
    Save-Shot "click_1_game"
    Send-Key 0x11 80; Start-Sleep -Milliseconds 700; Send-Key 0x11 80; Start-Sleep -Milliseconds 700      # fire twice
    Send-Key 0x51 120; Start-Sleep -Seconds 3                                                             # Q
    $idle = Get-Idle $xl
    "   game ended: $idle | tab now '$($wb.ActiveSheet.Name)' | note: $(Get-Note $wb) | box: $([int][char]$tick.TextFrame.Characters().Text)"
    if (-not $idle -or (Get-Note $wb) -notmatch '^Run 1:') { $ok = $false }
    Save-Shot "click_2_after_quit"

    "2. arrow keys over the picture must not start a game"
    Send-Key 0x27; Send-Key 0x28; Send-Key 0x28; Send-Key 0x28; Start-Sleep -Seconds 2
    $idle = Get-Idle $xl
    "   still idle: $idle | active cell $($xl.ActiveCell.Address($false, $false))"
    if (-not $idle) { $ok = $false; Send-Key 0x51 120; Start-Sleep -Seconds 3 }

    "3. click the picture to play again"
    $p = Get-PictureCentre $xl $wb
    Send-Click $p[0] $p[1]
    Start-Sleep -Seconds 7
    Save-Shot "click_3_second_game"
    $running = -not (Get-Idle $xl)
    Send-Key 0x51 120; Start-Sleep -Seconds 3
    "   second game was running: $running | note: $(Get-Note $wb)"
    if (-not $running -or (Get-Note $wb) -notmatch '^Run 2:') { $ok = $false }

    "4. the checklist: its checkbox starts a game too"
    $wb.Worksheets.Item("Checklist").Activate(); Start-Sleep -Milliseconds 800
    Save-Shot "click_4_checklist"
    $p = Get-Screen $xl.ActiveWindow ($tick.Left + $tick.Width / 2) ($tick.Top + $tick.Height / 2)
    Send-Click $p[0] $p[1]
    Start-Sleep -Seconds 7
    $running = -not (Get-Idle $xl)
    Send-Key 0x51 120; Start-Sleep -Seconds 3
    "   third game was running: $running | tab now '$($wb.ActiveSheet.Name)' | note: $(Get-Note $wb)"
    if (-not $running -or (Get-Note $wb) -notmatch '^Run 3:') { $ok = $false }
    "   workbook thinks it is saved (no prompt on close): $($wb.Saved)"
    foreach ($vk in 0x11, 0x51, 0x25, 0x26, 0x27, 0x28) { [W.I]::keybd_event([byte]$vk, 0, 2, [UIntPtr]::Zero) }
    $wb.Close($false)
} finally {
    $xl.Quit()
}
if ($ok) { "CLICK TEST PASSED" } else { "CLICK TEST FAILED" }

# Every Windows check in one go: build, engine self-check in Excel, offline emulator, scripted play
# with key combinations, the first-time-player path with real mouse clicks, and the picture test.
# It takes the keyboard and mouse for about four minutes: leave the PC alone while it runs.
# Run from the repo root:  powershell -ExecutionPolicy Bypass -File win\check.ps1
$env:PYTHONUTF8 = '1'
$py = ".\.venv\Scripts\python.exe"
$results = [ordered]@{}
function Note($name, $ok, $detail) { $results[$name] = @($ok, $detail); "{0,-6} {1}: {2}" -f $(if ($ok) { 'ok' } else { 'FAILED' }), $name, $detail }

$o = & $py mkdoom.py 2>&1
Note 'build (lint for Mac and for Windows)' ($LASTEXITCODE -eq 0) (($o | Select-Object -Last 1) -replace '^built ', '')
if ($LASTEXITCODE -ne 0) { $o; exit 1 }

powershell -NoProfile -ExecutionPolicy Bypass -File win\verify.ps1 | Out-Null
$o = & $py tests\check_verify.py verify_result.txt 2>&1
Note 'engine self-check in Excel' ($LASTEXITCODE -eq 0) ($o | Select-Object -Last 1)

& $py ref.py 40 | Out-Null
$o = & $py emu_run.py 3 2>&1 | Select-Object -Last 1
$m = [regex]::Match("$o", 'frames produced (\d+), reference has \d+, identical prefix (\d+)')
Note 'offline VBA emulator against reference' ($m.Success -and $m.Groups[1].Value -eq $m.Groups[2].Value) "$o"

$o = & $py win\playtest.py combos 2>&1
$bad = $o | Select-String -Pattern 'SKIPPED|did not end|no result|gone before|WARNING|other Excel windows=\[.+\]'
$after = ($o | Select-String -Pattern '^AFTER:').Line
Note 'scripted play, 12 key combinations' ($LASTEXITCODE -eq 0 -and -not $bad -and $after -match 'note=Run 1:') (($o | Select-String -Pattern '^avg fps').Line -replace '\|.*', '')
$sound = ($o | Select-String -Pattern '^SOUND:').Line
Note 'sound effects accepted by Windows' ($sound -match 'refused 0') ($sound -replace '^SOUND: ', '' -replace ' \| .*', '')
Start-Sleep -Seconds 4

$o = powershell -NoProfile -ExecutionPolicy Bypass -File win\clicktest.ps1 2>&1
Note 'first-time-player path, real clicks' ([bool]($o | Select-String -Pattern 'CLICK TEST PASSED')) (($o | Select-String -Pattern 'opened on').Line)
Start-Sleep -Seconds 4

& $py win\playtest.py Host.PlayRaw still | Out-Null
$o = & $py tests\picture_check.py 2>&1
$line = ($o | Select-String -Pattern 'engine frame vs real cells').Line
Note 'picture exact after standing still' ($line -match 'differ\s+0\.0 %') (($line -replace '\s+', ' ') -replace ' \| mean.*', '')
$line = ($o | Select-String -Pattern 'painter record vs real cells').Line
Note 'cells hold what the painter intended' ($line -match 'differ\s+0\.0 %') (($line -replace '\s+', ' ') -replace ' \| mean.*', '')

$failed = @($results.Keys | Where-Object { -not $results[$_][0] })
if ($failed.Count) { "`n$($failed.Count) of $($results.Count) checks FAILED"; exit 1 } else { "`nall $($results.Count) checks passed" }

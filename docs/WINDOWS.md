# Bringing this up on Windows

Nothing here has been run on Windows yet. This page is the handoff for whoever tries first, human or coding agent.

## What you need

- Excel for Windows, **64-bit** (File > Account > About Excel). The generated VBA uses `LongLong` and will not compile on 32-bit Excel.
- Python 3 with `pillow` (`pip install -r requirements.txt`; the optional packages are only for the reference comparison).

## Step 1: does the engine run?

```
python mkdoom.py
powershell -ExecutionPolicy Bypass -File win\verify.ps1
python tests\check_verify.py verify_result.txt
```

`Host.Verify` boots the game on a fixed virtual clock, runs 40 ticks with no input and returns one checksum per frame. All 82 must match `tests/verify_sums.txt`. This exercises the whole translated engine and the game-data loader, and nothing platform-specific except that the host module has to compile.

If Excel reports a compile error, it is almost certainly in a `#Else` branch of `hostvba.py` or `televba.py`: those branches were written on a Mac and have never been compiled.

## Step 2: what does painting cost?

Run `Host.Audit` (same way as Verify, with `-Macro Host.Audit`). It returns timings in milliseconds for Excel's full-window redraw under different conditions. On Mac the full redraw costs about 17 ms plus about 0.5 microseconds per formatted cell on screen, and that fixed cost is what limits the frame rate in motion. Windows may be better or worse; nobody knows yet.

## Step 3: play, and look for these Mac-specific assumptions

| Area | Mac behaviour | What to check on Windows |
|---|---|---|
| Keyboard | Key state read with CoreGraphics; game keys disabled in Excel with `Application.OnKey` | `GetAsyncKeyState` branch; virtual-key codes in `InitKeys` |
| Mouse look | Pointer re-centred every frame on the empty menu bar, because a click on the grid makes Excel run its own drag loop and freezes the macro until the button is released | There is no Mac menu bar. Mouse look starts **off** on Windows. Find somewhere a click is harmless, or leave it keyboard-only |
| Front-app check | Process Manager calls | `GetForegroundWindow` against `Application.Hwnd` |
| Timing | `mach_absolute_time` for measurement; `Timer` for the game clock | `QueryPerformanceCounter` branch |
| Small frames | Painted without touching `Application.ScreenUpdating`; they do reach the screen at the next `DoEvents` | Whether the screen really updates. `playtest.py` on Mac compares screenshots with the engine's own state; do the equivalent |
| Layout | Column widths are corrected at start by reading `.Width` back | Whether the picture is square and fits the window |

Run the macro `PlayDoomWithStats` for a real session and keep the report it writes to `%USERPROFILE%\doom_telemetry_<stamp>.txt`. It breaks every frame down into engine, scan, fill and redraw time.

## Things already learned the hard way

See the "learned the hard way" sections of [LAB-NOTES.md](LAB-NOTES.md). The short list:

- Writing `Application.ScreenUpdating`, even to the value it already has, forces a full-window redraw.
- `Application.OnKey` handler macros do not run while another macro is running.
- A VBA compile error shows a modal dialog and leaves the VBA engine stuck until Excel restarts.
- Module-level declarations must come before every procedure; identifiers are case-insensitive.
- A VBA procedure call costs about 0.2 microseconds, more than most of the arithmetic in this code, so helpers are inlined.

## Sending changes back

Keep one codebase: platform differences belong inside `#If Mac Then ... #Else ... #End If` in `hostvba.py` and `televba.py`. The engine, translator and runtime are platform-neutral and are verified on the Mac side; if Step 1 fails with a checksum mismatch rather than a compile error, that is worth an issue with the index of the first differing frame.

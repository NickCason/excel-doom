# Bringing this up on Windows

First run on Windows: 2026-11-18, Excel 16.0 build 20527 (64-bit), Windows 11, Core Ultra 7 155H, 1920 × 1080. The steps below are the ones that were followed; what they found is in [First results](#first-results) at the end.

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

`python win\playtest.py` is the automated version: it starts the game, injects key presses, saves screenshots of the Excel window to `out\wpt*.png` and prints the engine's own state at the end. `python win\playtest.py combos` also holds Shift and Ctrl together with game keys. Each run writes one telemetry report to `%USERPROFILE%`.

## Things already learned the hard way

See the "learned the hard way" sections of [LAB-NOTES.md](LAB-NOTES.md). The short list:

- Writing `Application.ScreenUpdating`, even to the value it already has, forces a full-window redraw.
- `Application.OnKey` handler macros do not run while another macro is running.
- A VBA compile error shows a modal dialog and leaves the VBA engine stuck until Excel restarts.
- Module-level declarations must come before every procedure; identifiers are case-insensitive.
- A VBA procedure call costs about 0.2 microseconds, more than most of the arithmetic in this code, so helpers are inlined.

## Sending changes back

Keep one codebase: platform differences belong inside `#If Mac Then ... #Else ... #End If` in `hostvba.py` and `televba.py`. The engine, translator and runtime are platform-neutral and are verified on the Mac side; if Step 1 fails with a checksum mismatch rather than a compile error, that is worth an issue with the index of the first differing frame.

## First results

**Step 1 passes.** The workbook builds with Python 3.11 on Windows, every module compiles, and all 82 frame checksums match. Data load 0.25 s, boot 0.5 s, 40 ticks in 0.7 s.

**Step 2, full redraw.** `Host.Audit` chose 50 % zoom in the unmaximised window Excel opens under automation: base redraw 26 ms, the other conditions between 8 and 39 ms. Single runs with `Timer`, and they are noisy (the same measurement gave 16 ms and 8 ms at the start and end of one run), so treat them as order of magnitude.

**Step 3, play.** With the window maximised the layout picks 75 % zoom, the picture is square and fits. The game boots into E1M1, the screen follows the engine (ammo, health and position in the screenshots match the engine's state), and Q quits cleanly.

| Area | Result on Windows |
|---|---|
| Keyboard | `GetAsyncKeyState` branch works: taps and holds of Ctrl, Shift, arrows and W all reach the engine (five 80 ms fire taps, five shots) |
| Front-app check | `GetForegroundWindow` against `Application.Hwnd` works |
| Timing | `QueryPerformanceCounter` branch works, 100 ns resolution |
| Small frames | They do reach the screen, but see the cost below |
| Layout | Square, fits a maximised 1920 × 1080 window at 75 % |
| Mouse look | Not tested. Still off by default |

Two input problems were found and fixed in `hostvba.py`:

- **Ctrl plus a game key is an Excel shortcut.** Holding fire while walking forward is Ctrl+W, which closes the workbook: a "Save your changes?" dialog opens over the game and the game stops taking keys. `HookKeys` now also disables Ctrl and Ctrl+Shift with every game key on Windows. Checked with Ctrl+W, Ctrl+Shift+W, Ctrl+A, Ctrl+D, Ctrl+S, Ctrl+1, Ctrl+Space and Ctrl+Left.
- **Alt cannot be a game key.** Alt puts the ribbon into KeyTips mode and the next game keys then drive the ribbon (Alt, A, S walks into Data > Sort). `Application.OnKey` does not prevent it. Alt is no longer read on Windows; strafe with A and D.

**Speed: the switch between the two paint paths was in the wrong place for Windows, and is now moved.** With the Mac threshold (full redraw from 1800 rectangles) a 31 s run with firing and running looked like this:

| Frame class | Frames | Rectangles per frame | Fill time per frame | Fill time per rectangle |
|---|---|---|---|---|
| `direct` (small frame, no full redraw) | 105 | 142 | 119 ms | about 840 µs |
| `full_big` (full redraw) | 70 | 3750 | 23 ms | about 6 µs |

A fill made while `ScreenUpdating` is on costs 0.4 to 0.8 ms per rectangle on Windows, about a hundred times more than one made while it is off. A full redraw costs about 50 ms whatever the frame (Mac: about 17 ms). So the small-frame path only pays below about a hundred rectangles, and `Present` now uses `BATCHWIN = 100` on Windows instead of `BATCHMIN`.

The three modes on the same scripted run (five fire taps, a turn, a walk), then the new threshold:

| Mode | Average fps | Small-frame path | Full-redraw path |
|---|---|---|---|
| `Host.PlayA`, never a full redraw | 9.6 | 643 rectangles, 295 ms each | |
| `Host.PlayB`, always a full redraw | 11.8 | | 1000 rectangles, 59 ms each |
| `Host.Play`, threshold 1800 | 22.5 | 155 rectangles, 56 ms each | 3949 rectangles, 70 ms each |
| `Host.Play`, threshold 100 | 30.3 | 10 rectangles, 5 ms each | 2541 rectangles, 55 ms each |

By what the player is doing, from the telemetry reports (frames per second):

| Run | Threshold | Idle | Turning | Moving | Whole run |
|---|---|---|---|---|---|
| Basic | 1800 | 25 | 11 | 10 | 22.5 |
| Basic | 100 | 36 | 11 | 10 | 30.3 |
| With firing and running | 1800 | 22 | 10 | 5 | 15.5 |
| With firing and running | 100 | 27 | 9 | 10 | 21.3 |

These are single runs of 14 to 31 s on one machine. A still picture runs at about 45 fps. Motion is now held at about 10 fps by the full redraw: a moving frame is about 15 ms of engine, 5 ms of fills and 50 ms of redraw. Always taking the full redraw is worse than either threshold, because the weapon bob and status-bar frames then pay 50 ms each. Anything faster in motion has to make the full redraw itself cheaper (the audit suggests it scales with what is on screen: zoom and the number of visible rows).

Not yet done: a session played by a person, mouse look, and any second machine or Excel version.

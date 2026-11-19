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

## Picture quality: streaks and flat patches

The complaint from the first person to play it on Windows: horizontal smearing and grey or black patches. `Host.PlayRaw` plus `Host.Dump` (written out by `win\verify.ps1 -Dump`, compared by `tests\picture_check.py`) take three pictures of the same instant: the engine's frame, what the painter believes each cell shows, and the fill colour each cell really has.

- **The cells are right.** Painter's record against real cells: no difference in any run. Excel paints exactly what it is told.
- **The painter is the cause.** Engine frame against cells, stopping in mid-motion: 55 % of pixels differ, by up to 40 per colour channel, and the mean same-colour run along a row is 4.9 pixels against the engine's 2.1. That is the near-colour merging in `Present`. It paints colours within a tolerance as one, the tolerance rises to 24 while frames are busy, and a cell already within tolerance of the right colour is not repainted, so pieces of old frames stay.
- **It did not clear at rest.** The tolerance only came back down while frames were small. In a room with an animated floor they never are, so the walls stayed streaked for as long as the player stood there.

What changed, on Windows only (`TolLo`, `TolHi`, `TolSettle` in `RunDoom`; `ViewCheck`):

- The ceiling is 16 instead of 24. At 24 dark areas lose their detail: 8,500 pixels were off by 17 to 24; at 16, under 100.
- When the player has neither moved nor turned for ten frames, the tolerance drops to 0. That frame rescans everything and repaints it exactly (5,600 rectangles, 70 ms in the test), and it stays exact while only parts of the picture change. Measured after standing still: 0.0 % of pixels differ.
- A fix on the way: the old rule could overshoot its ceiling (18 seen with a ceiling of 16).

Frame rate on the scripted run is unchanged within noise (about 22 fps average, about 10 moving). While moving, the picture is still merged at tolerance 8 to 16; that is the price of the frame rate. `DOOM_TOL=lo,hi,settle` at build time overrides the three numbers (`8,24,8` is the Mac behaviour). The Mac build is unchanged and has the same merging.

Seen while measuring, not fixed: when frames are slow (200 ms and more) the game falls behind the keys. A 2.5 s key hold kept the view moving for about 4.5 s.

## Sound

Sound effects play on Windows. Music does not, and the Mac build stays silent.

- **Engine side** (`doom/dg_xl.c`): a sound module for the engine's own `sound_module_t` interface, switched on with `-DFEATURE_SOUND`. When the game starts an effect, the module finds the lump in the WAD, checks its header and hands the 8-bit samples to the host through one new import, `play_sound(pcm, length, rate, volume, separation)`. `doom/SDL_mixer.h` is an empty file: `i_sound.c` includes that header when sound is on and uses nothing from it.
- **Host side** (`imp_play_sound` in `hostvba.py`): copies the samples behind a 44-byte wave header, scaled by the game's volume, and calls `PlaySound` with `SND_MEMORY` and `SND_ASYNC`. Two buffers take turns because the one that is playing must not be written. `RunDoom` stops playback on the way out, before the buffers can go away. On Mac, and in `Host.Verify`, the host declines and nothing is played.
- **One voice.** `PlaySound` plays one sound at a time, so a new effect cuts the last one off. The engine module lets a new effect in unless the one playing has the more important priority (the lower number in `sounds.c`), and counts the playing time in game tics. Stereo separation is passed but not used.

Checked on the first machine:

- The frame checksums are unchanged: 82 of 82 under wasmtime and in Excel, so the sound code does not disturb the game.
- Under wasmtime, five pistol shots give five calls of 5629 samples at 11025 Hz.
- In Excel, a 29 s scripted run asked for 19 effects and Windows started all 19, at 0.3 ms each in the host. The frame rate did not change.
- The peak meter of the default output device was silent for the whole of a second run except the 3.5 s of the five fire taps.

Not checked: how it sounds. Nobody has listened to it yet.

## Around the game

The workbook opens on the game sheet: the title screen in the cells with a click-to-play sign over it. The sheet is stored at 50 % zoom so the whole frame shows even on a small screen; `Workbook_Open` calls `Splash`, which fits it to the window and removes the two lines under the frame that say how to allow macros (they are only seen by someone whose macros are off). Behind it are a cheat sheet and a checklist (`frontsheets.py`). The checklist is the scoreboard: after each game `LogRun` ticks the box on its last task and writes kills, items, secrets, time and map into the row. That box is a text box that runs `PlayDoom`.

Three things that were learned building this:

- **One more picture on the game sheet halves the frame rate.** The first click-to-play sign was a fifth picture over the middle of the game. With it the same scripted run averaged 14 fps instead of 28, and hiding it or parking it below the frame during play made no difference. So the sign is not a picture: it is part of the title screen in the cells at build time, and when a game ends `GameOff` writes it into the engine's last frame and sends that through `Present`.
- **A click on the picture is a `SelectionChange`.** With no picture to carry a macro, the sheet's `SelectionChange` event calls `CellClick`. Arrow keys fire the same event, but only ever move one cell, so only a selection that jumps into the picture from further away starts a game. `win\clicktest.ps1` checks both.
- **The sign must be painted before `Playing` goes false.** `Present` ends in `DoEvents`, and a click waiting in the queue would otherwise start a second game inside the first one's clean-up.

Keys the host handles itself rather than passing to the engine: V (sound), 8, 9 and 0 (it queues the key presses of `iddqd`, `idkfa` and `idspispopd`). Their Ctrl combinations are disabled with the rest; Ctrl+P, Ctrl+V and Ctrl+0 were checked.

`win\clicktest.ps1` opens the workbook, clicks the splash with a real mouse click, fires, quits, presses arrow keys over the picture, clicks the picture to play again, starts a third game from the checklist's box and reads the checklist row back.

## Sticky Keys

Shift is the run key, and Windows listens to it: five quick taps offer Sticky Keys, and eight seconds on the right-hand Shift offer Filter Keys. Either puts a prompt over the game and takes the keyboard.

`AccessKeys` switches both shortcuts off when a game starts and puts them back when it ends, with `SystemParametersInfo` (`SPI_SETSTICKYKEYS`, `SPI_SETFILTERKEYS`). The rules it follows:

- If the feature itself is turned on, the player needs it, and that one is not touched.
- Only the two shortcut bits change, and only for the Windows session; nothing is written to the profile.
- Putting back reads the settings again and sets only the bits that were taken out, so a change made in Settings during the game survives.
- It runs in the same clean-up as the key hooks, so an error or Esc path restores it too. If Excel is killed mid-game, the shortcuts stay off until sign-out.

`win\stickytest.py` reads the two settings before, during and after a short game. On this PC: shortcut active before, off during, the same value as before afterwards (510 and 126). Injected Shift taps do not raise the prompt, on the old build or the new one, so the prompt itself is not covered by a test: only a person tapping Shift can show that.

Not yet done: a session played by a person, mouse look, and any second machine or Excel version. The automated tests open the workbook through automation, where macros are already allowed; nobody has yet opened it by double-clicking and pressed Enable Content.

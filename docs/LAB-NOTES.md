> Working notes kept during development, in order. File names refer to the layout at the time; early
> measurement scripts now live in `experiments/`. Numbers are from one Apple Silicon MacBook.

# Excel Doom feasibility spike — results (2026-11-18)

Throwaway spike. Machine: M5 MacBook, Excel for Mac 16.107 (`/Applications/Microsoft Excel.app`, used as the test copy).

## Measured
| Question | Result |
|---|---|
| Can Excel load a huge generated VBA project? | 249,348 lines / 40 modules opened in ~4 s and ran correctly |
| Procedure size limit | ~2,500-line procedure compiled; ~10,000-line one failed ("Procedure too large") |
| Doom as WebAssembly (doomgeneric, zig cc -O2, 8-bit output) | 801 functions, 112,881 instructions, 26 MB memory; largest function 4,677 instructions |
| Operations Doom needs | boot 25.6 M; title screen 0.7 M/frame; gameplay ~2.2 M/frame median (p95 ~2.8 M) |
| Translated VBA speed (3 Doom-shaped kernels) | 72–132 M ops/s, checksums identical to wasmtime |
| Emulator-in-VBA speed (same kernels) | 13–19 M ops/s, checksums identical |
| Byte-array memory vs Long-array memory | Long-array faster on all three kernels (Byte: 42–105 M ops/s) |
| VBA primitives | statement ~5 ns; wrapped 32-bit add ~14 ns; procedure call ~221 ns |
| Display: full frame as 192,000 cell values + colour scales | ~2.35 s per frame (one run) |
| Display: per-cell fill colour | ~19 µs per cell (64,000 cells ≈ 1.2 s) |

## Projected (not measured on Doom itself)
- Translated: ~17–31 ms of engine time per gameplay frame, boot under 0.5 s.
- Emulated: ~115–170 ms per frame, boot ~2 s.
- The cell display, not the engine, is the bottleneck.

## Not proven
- Windows (untested). Generated code uses LongLong, so 32-bit Office would need another arithmetic path.
- Full Doom translation: needs floats (11 functions), more 64-bit ops, memory.copy/fill/grow, 15 host imports (file I/O for the WAD).
- Trusting compiler alignment hints for 32-bit loads; wrapping 64-bit arithmetic; input; sound; embedding the 4 MB WAD.

## Files
- `mkxlsm.py` builds a .xlsm from VBA source with no Excel involved; `xl.py` drives the test Excel copy.
- `wasm.py` parser, `w2vba.py` translator, `w2interp.py` emulator generator, `vbalint.py` pre-flight check.
- `build_doom.sh`, `count_ops.py`, `bench/bench.c`, `mkbench.py`, `runbench.py`, `mkdisp.py`.
- `out/disp_screen.png`: a real Doom frame shown in cells.

## Display follow-up (same day) — `mkdisp2.py`, two real frames 8 game tics apart (75% of pixels differ)
Per frame, write + repaint, test Excel frontmost at 60% zoom:

| Scheme | Cells | Colour | Time |
|---|---|---|---|
| Values, 3 cells/pixel, colour scales (original) | 192,000 | full, dim | 2.5–2.6 s |
| Same at half resolution | 48,000 | full, dim | 2.2 s (does not scale: cost is the 960-area conditional format) |
| Values, 1 cell/pixel, one colour scale | 64,000 | greyscale | 0.34 s |
| Fill colour on every cell | 64,000 | full | 0.57 s |
| Fill colour on changed cells only | 64,000 | full | 0.45 s |
| Fill colour per same-colour run | 64,000 | full | 0.33 s |
| Fill colour per changed same-colour run | 64,000 | full | 0.26 s |

Typed (Double) arrays made no difference to value writes. Screenshots: `out/disp2_fill.png`, `out/disp2_grey.png`.

## Display tricks (same day) — `dispsim.py` (30 real frames, simulated) and `mkdisp3.py` (Excel, 3 frames, run twice)
Measured in Excel, write + repaint per frame, full colour, one cell per pixel:

| Scheme | Fill calls | Time | Quality vs true frame |
|---|---|---|---|
| Changed runs (previous best) | 18,847 | ~280 ms | exact |
| Changed runs joined into rectangles | 13,687 | ~215 ms | exact |
| Near-colour merging, tolerance 8 | 7,275 | ~150 ms | 38.6 dB, near-lossless |
| Near-colour merging, tolerance 16 | 4,165 | ~126 ms | 33.3 dB, walls smear slightly, HUD crisp |
| Doom low-detail mode | 9,272 | ~230 ms | no gain |
| Small view (size 8) | 9,993 | ~229 ms | no gain |
| Not batching with ScreenUpdating | 7,275 | ~340 ms | worse |

Fixed full-window redraw after each batch: 48–88 ms depending on how much sheet is visible; gridlines/headings off made no difference.
Simulated only: plain half res (22.5 dB, HUD unreadable), interlacing (22 dB, combing), fixed call budget worst-first (speckle),
2-colour quadrant glyphs at half res (6,400 calls at 34 dB: no better than tolerance 16, plus font risk).
Side-by-side picture: `out/display_options.png`.

## First playable build (same day) — `mkdoom.py` → `out/doom.xlsm` (copy at `../doom.xlsm`)
- Whole engine translated: 802 functions, 112,905 wasm instructions → ~44,700 lines of VBA (+ data), 32 modules, 7.3 MB workbook.
- Shareware WAD stored as ~1M numeric cells (columns 401–416), loaded in 0.10 s.
- Verified in real Excel: 82 frames (boot + 40 ticks of E1M1, deterministic clock) byte-identical to wasmtime (`mkdoom.py verify`).
- Engine alone: ~24 ms per tick in Excel. Live play with the cell display: ~8–11 fps on this Mac.
- Pipeline: `build_xl.sh` (zig → wasm) → `w2v.py` (wasm → VBA) + `rtvba.py` (runtime, WASI shim) + `hostvba.py` (display, keys, loop).
- Offline checks: `vbaemu.py` runs the generated VBA in Python (`emu_run.py` compares frames with `ref.py`); `mkselftest.py` confirms
  real Excel matches the emulator on 1,164 runtime-helper probes.
- Macros: `PlayDoom` (or `Host.Play`), `Host.Verify`. Keys: arrows/WASD, F or Ctrl fire, Space/E use, Shift run, Esc menu, Q quit.
- Not done: Windows test, sound, saving, design doc. Keyboard uses an OS key-state call and only listens while Excel is frontmost.

## Second playable build (2026-11-18 evening): mouse, tap handling, tutorial frame
- Mouse look + left click fire (cursor re-centred each frame while Excel is frontmost; M toggles). Engine gets `ev_mouse` from `dg_tick`.
- Keys polled many times per frame; a release is only forwarded once the engine has consumed the matching press, so short taps register.
- Tutorial frame: four PNG pictures (Doom logo, border flat, HUD font) anchored over a few large cells; game is 320x200 cells of 4 px at B2.
- Adaptive tolerance (8 to 20) when a frame needs more than 6000 fills.
- Measured on this Mac: ~11 fps average, ~6 fps while turning; display trails the engine by about one frame.
- `dg_stat` export + per-frame log give ground truth; `playtest.py`, `latency.py`, `freezetest.py` drive it with injected input.

### Things learned the hard way
- Calling `CGEventSourceCounterForEventType` from VBA stops Excel repainting the grid for the rest of the macro. Do not use it.
- Setting `Application.ScreenUpdating` (even to its current value) forces a full-window redraw costing ~1.1 us per visible cell;
  a frame drawn as 132k cells cost ~250 ms per redraw, as pictures over big cells ~80 ms.
- Skipping the ScreenUpdating toggle for small frames ran at ~28 fps but left stale pictures in the live game (fine in isolated probes). Unsolved.
- `sheetFormatPr defaultRowHeight` is ignored: put `ht` on every row. XML column widths do not convert to pixels predictably: fix in VBA
  by reading `.Width` back. `Timer` on Mac ticks every 15.6 ms. `Application.OnKey` macros do not run while another macro is running.
- Always compare a screenshot against engine truth before believing a display change works.

## Performance audit (2026-11-18 night)
Per-frame budget before (ms): engine 20, picture scan 7, cell fills 22 (16 us each), Excel full redraw 68. Total ~117 (8.5 fps).
Real 9-minute session: 8.7 fps, 90 of 114 ms per frame in paint with only ~380 fills per frame.

What the full redraw costs (Audit macro): ~17 ms for the window + ~0.16 ms per visible row (32 ms for 200 rows; column count,
zoom, the frame pictures and the formula bar make no difference) + ~1.25 us per colour run on screen.

Changes and effect:
- Frames under 800 rectangles are painted without toggling ScreenUpdating (no full redraw). Over half of all frames are under 200.
- Rectangles are bucketed by palette colour and sent ~15 per call through multi-area Range addresses: 16 -> 3.3 us per rectangle.
- Rows whose pixels did not change are skipped in the scan: 7 -> 1 ms.
- One full "settle" redraw after 4 still frames; full redraw forced while a non-modifier key is held and every 30 small frames.
- WAD cells moved below the game rows (about -4 ms per full redraw).
- Scripted test: 8.4 -> 22.5 fps average. Mouse turning: 6.4 -> 9.0 fps. Full-redraw frames ~72-81 ms for 3000-4300 rectangles.

Findings to remember:
- Small un-batched updates DO reach the screen, except that holding F (fire) left the picture stale in both paint paths while
  Ctrl-fire and W/arrow movement were fine. Cause not found; F was removed as a fire key (fire = Ctrl or left click).
- Un-batched painting is slower than a full redraw above roughly 1000 rectangles (4.4 vs 8.9 fps when turning).
- Ctrl+Space style combos in test scripts can open the macOS emoji picker, which then swallows injected keys.
- Next levers, not built: a fast mode (100 rows + Doom low detail, est. 13-15 fps while moving); engine time (23 ms) is now
  the bulk of quiet frames.

## First telemetry session (2026-11-18 19:43, 516 s, 13,911 frames, 27.0 fps average)
- Time split: idle 61% at 35.5 fps (Doom's 35 Hz cap); moving 19% at 10.2 fps; moving+turning 13% at 10.7; turning 4% at 9.3.
- Action (full-redraw) frame, 98 ms: redraw 57.6, engine 23.0, fills 10.4, scan 6.6. Regression over 1,957 frames:
  paint = 63.6 ms + 3.3 us per rectangle, so the full redraw is effectively a fixed cost; rectangle count hardly matters.
- Direct (no full redraw) frames: paint = 3.3 ms + 28 us per rectangle; break-even with the full redraw at ~2,400 rectangles
  (threshold raised from 800 to 1,800).
- Engine time does not depend on game tics per frame (slope ~0): it is all rendering, ~23 ms.
- 44 frames of 300-480 ms matched 52 mouse clicks: a click on the grid puts Excel in a drag-tracking loop for as long as the
  button is held, freezing the macro. Fix: park the cursor on the empty menu bar (0.54 x screen width, y = 4). After the fix
  three 0.6 s clicks gave no frame over 200 ms and auto-fire works while held.
- Player used W/A/S/D, mouse, Space; never Ctrl. 224 key presses, all delivered; no fill errors.
- Full redraw cost model refined: ~17 ms window + ~0.5 us per formatted cell record in visible rows (64,000) + small per-run term.

# DOOM in an Excel sheet

The real DOOM engine running inside a single Excel workbook. Every cell is a pixel. No add-ins, no DLLs, no helper programs: one `.xlsm` file and Excel.

![DOOM running in worksheet cells](docs/screenshot.png)

**[Download ExcelDoom.xlsm](https://github.com/NickCason/excel-doom/releases/latest/download/ExcelDoom.xlsm)** (7.7 MB, includes the shareware episode)

## Status

| Platform | State |
|---|---|
| Excel for Mac 16.x, Apple Silicon | Works. About 35 fps standing still, about 10–12 fps while moving. |
| Excel for Windows, 64-bit | Runs, lightly tested on one machine. The engine self-check passes. About 45 fps standing still, about 10 fps while moving. Keyboard only. See [docs/WINDOWS.md](docs/WINDOWS.md). |
| 32-bit Excel, Excel Online, Numbers, LibreOffice | Not supported. |

This is an experiment, not a product. Sound effects on Windows only, one at a time; no music. No saving. It starts directly on E1M1.

## How to play

1. Download the workbook and open it in desktop Excel.
2. Allow macros.
   - **Mac:** choose *Enable Macros* when Excel asks.
   - **Windows:** Office blocks macros in downloaded files. Right-click the file, choose *Properties*, tick *Unblock*, then open it and choose *Enable Content*.
3. The workbook opens on the title screen. Click it. (Or run the macro `PlayDoom`.)
4. Press **Q** to quit back to Excel. Click anywhere on the picture to play again. The Cheat Sheet tab lists the keys and cheats; the Checklist tab notes how each run went.

| Key | Action |
|---|---|
| W / S or Up / Down | Forward / back |
| Left / Right | Turn |
| Mouse (Mac) | Turn; left click fires |
| A / D | Strafe |
| Ctrl | Fire |
| Space or E | Open doors, use |
| Shift | Run |
| 1–7 | Weapons |
| Tab | Map |
| Esc, Enter, Y, N | Menu |
| M | Mouse look on/off |
| P | Pause |
| V | Sound on/off (Windows) |
| 8 / 9 / 0 | Cheats: god mode / all weapons, keys and ammo / walk through walls |
| Q | Quit |

The cheat keys type the original codes (`IDDQD`, `IDKFA`, `IDSPISPOPD`) into the engine for you. They cannot be typed by hand: most of their letters are not game keys here, and Q quits.

On Mac, mouse look pins the pointer to the top of the screen while Excel is in front. Press **M** to free it, or switch apps. Mouse look starts off on Windows.

## What the macros do

Macros are the whole point here, so it is fair to ask what they touch.

- They change cell colours on the game sheet, and the zoom and gridline settings of that window. One macro runs when the workbook opens: it fits the title screen to the window and removes the two lines under it about allowing macros. After each game they tick the box on the checklist, write one line about the run next to it, and mark the workbook as saved so that closing it does not ask about the changed colours.
- While a game is running they ask the operating system which keys and mouse buttons are down (CoreGraphics on Mac, `GetAsyncKeyState` on Windows), and only while Excel is the front application. Some antivirus tools dislike workbooks that declare those calls.
- On Windows they play the game's sound effects through the system's own wave player (`PlaySound` in `winmm`), from memory.
- They do not use the network and do not write files. The optional macro `PlayDoomWithStats` writes one performance report to a text file in your home folder.

All of the VBA is generated from the sources in this repository.

## How it works

1. **The engine is the real one.** [doomgeneric](https://github.com/ozkl/doomgeneric) (the id Software source release with a thin platform layer) is compiled to WebAssembly with `zig cc`.
2. **WebAssembly is translated to VBA** ahead of time by `w2v.py`: about 113,000 instructions become about 45,000 lines of VBA across 30 modules. Linear memory is one `Long` array.
3. **The game data lives in cells.** The shareware WAD is stored as about a million numbers far to the right of the picture and loaded into an array at start-up. The C library's file calls are answered from that array (`rtvba.py`).
4. **The screen is 320 × 200 cells.** Each frame, changed pixels are grouped into same-colour rectangles and painted with a few hundred multi-range `Interior.Color` calls (`hostvba.py`). Small changes are painted directly; big ones use one full redraw.
5. **The frame around the game** is four pictures built from the game's own logo, border texture and status-bar font (`frameart.py`).
6. **The workbook is written from scratch** by `mkxlsm.py`, including the VBA project container, so the build needs no copy of Excel.

Correctness is checked frame by frame against the same WebAssembly module running under wasmtime, both in a small offline emulator of the generated VBA (`vbaemu.py`) and inside Excel itself (`Host.Verify`).

A short account of the method, for anyone thinking of using it for something else, is in [docs/TECHNICAL-OVERVIEW.md](docs/TECHNICAL-OVERVIEW.md). Measurements, dead ends and the things Excel does that cost hours are in [docs/LAB-NOTES.md](docs/LAB-NOTES.md).

## Build from source

```
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python mkdoom.py            # writes out/doom.xlsm
.venv/bin/python mkdoom.py verify     # Mac only: also runs the in-Excel self-check
```

A prebuilt engine (`work/doomxl.wasm`) is checked in, so only Python is needed. To rebuild the engine itself install [zig](https://ziglang.org/) and run `./build_xl.sh` (on Windows: from Git Bash, with `zig` on the `PATH`). The checked-in engine was built with zig 0.17.1.

| File | Purpose |
|---|---|
| `doom/dg_xl.c`, `build_xl.sh` | Platform layer and engine build |
| `wasm.py`, `w2v.py` | WebAssembly parser and WebAssembly-to-VBA translator |
| `rtvba.py` | Hand-written VBA runtime: memory helpers, 64-bit maths, file calls |
| `hostvba.py`, `televba.py` | Excel side: display, keyboard, mouse, main loop, telemetry |
| `frameart.py`, `mkxlsm.py`, `mkdoom.py` | Frame artwork, workbook writer, build script |
| `frontsheets.py` | The cheat sheet and the checklist (the scoreboard) |
| `win/verify.ps1`, `win/playtest.py`, `win/clicktest.ps1` | Windows test drivers: self-check, scripted play with injected keys, first-time-player path with real mouse clicks |
| `tests/check_verify.py`, `tests/picture_check.py` | Compare the self-check's checksums, and the cells with the engine's frame |
| `vbaemu.py`, `emu_run.py`, `ref.py` | Offline VBA emulator and reference comparison |
| `xl.py`, `playtest.py`, `latency.py`, `freezetest.py`, `turntest.py`, `clicktest.py` | Mac test drivers (AppleScript and injected input) |
| `experiments/` | Earlier measurements and prototypes, kept for the record; not maintained |

## Credits and licence

- DOOM engine source © id Software, released under the GNU GPL v2. Platform layer from [doomgeneric](https://github.com/ozkl/doomgeneric) by ozkl.
- The embedded data is the freely distributable shareware episode of DOOM, © id Software. Its graphics are used unmodified for the frame artwork.
- Everything in this repository, and the generated workbook, is licensed under the [GNU General Public License v2](LICENSE).

This is an unofficial fan project. It is not affiliated with or endorsed by id Software, ZeniMax, Bethesda or Microsoft. DOOM is a trademark of id Software LLC.

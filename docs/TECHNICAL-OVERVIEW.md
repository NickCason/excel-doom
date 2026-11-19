# Technical overview: a compiled C program running inside an Excel workbook

This describes how the DOOM workbook is built, for anyone weighing the same method for other work in Excel. It is written to be read on its own.

## What it is

One macro-enabled workbook (`.xlsm`, 7.7 MB) that runs the original DOOM engine. No add-in, no DLL, no helper program, nothing installed. The engine is the real C source, not a rewrite. The game is incidental: what the project shows is that **an existing C program can be compiled, translated to VBA by a script, and run correctly inside a plain workbook**.

## The pipeline

1. **C to WebAssembly.** The C source is compiled with `zig cc` to a WebAssembly module (about 800 functions, about 113,000 instructions, 26 MB of memory).
2. **WebAssembly to VBA, ahead of time.** A Python script (`w2v.py`) translates every function into a VBA procedure. The program's memory is one VBA `Long` array. The result is about 45,000 lines of translated code; the whole VBA project is about 80,000 lines in 34 modules.
3. **A small hand-written runtime** (`rtvba.py`) supplies what VBA lacks: unsigned and 64-bit arithmetic, byte access into the `Long` array, memory copy and fill.
4. **A host layer** (`hostvba.py`) answers the program's calls to the outside world: clock, keyboard, display, sound, and file reads.
5. **The workbook is written by a script, not by Excel.** `mkxlsm.py` writes the `.xlsm` zip and the binary VBA project container directly, so the build needs only Python and runs anywhere.

Data the program expects in a file (4 MB here) is stored as about a million numbers in worksheet cells and read into an array at start-up in 0.1 to 0.3 seconds. The program's file calls are answered from that array.

## How it is known to be correct

The same WebAssembly module runs under a standard runtime (wasmtime) with a fixed clock, and a checksum is taken of every output frame. The workbook has a self-check macro that does the same run inside Excel. All 82 checksums must match. They did on Mac for the first build and do on Windows for the current one. A Python emulator of the generated VBA allows the same comparison without Excel. A lint pass catches undeclared names before Excel sees the code, because a VBA compile error shows a modal dialog and leaves the session stuck.

## Measured numbers

| What | Result | Where |
|---|---|---|
| Translated VBA speed | 72 to 132 million simple operations a second | Mac |
| One VBA statement / one procedure call | about 5 ns / about 220 ns | Mac |
| Interpreting WebAssembly in VBA instead of translating | 13 to 19 million operations a second (about 6 times slower) | Mac |
| Loading a 250,000-line generated VBA project | about 4 seconds | Mac |
| Largest procedure that compiles | about 2,500 lines works, 10,000 fails | Mac |
| Engine time per frame (a full 3D render) | 10 to 24 ms | both |
| Setting the fill colour of one cell, naively | about 19 µs | Mac |
| Same, as same-colour rectangles in multi-area `Range` calls | 3 to 6 µs per rectangle with screen updating off | both |
| Same with screen updating on | 0.4 to 0.8 ms per rectangle | Windows |
| One full-window redraw during play | about 50 to 60 ms, nearly independent of how much changed | both |
| Playing a sound from memory (`PlaySound`) | under 1 ms per call | Windows |

Windows figures are from one machine (Core Ultra 7 155H, 64-bit Excel 16.0, 1920 × 1080).

## What Excel does that shaped the design

- **Procedure calls are expensive, statements are cheap.** Helpers are inlined by the translator.
- **Changing `Application.ScreenUpdating` forces a full redraw**, even when set to its current value. Batches of cell changes are worth it only above a threshold, and that threshold differs by platform (1800 rectangles on Mac, 100 on Windows).
- **Picture objects on a sheet cost redraw time even when hidden.** One extra picture halved the frame rate on Windows.
- **`Application.OnKey` macros do not run while another macro is running.** Input is read by polling the operating system's key state, and only while Excel is the front window. When no macro is running, a click on a cell arrives through the sheet's `SelectionChange` event.
- **Excel's own shortcuts stay live.** Every key the program uses, and its Ctrl combinations on Windows, must be disabled with `Application.OnKey` for the duration. Alt cannot be used at all: it puts the ribbon into KeyTips mode.
- **Platform differences sit in `#If Mac` blocks** in the host layer only. The translated program and runtime are identical on both.

## Limits

- Needs macros. Files from email or the web are blocked until the user unblocks them, and company policy may forbid them outright.
- Declares operating-system calls (key state, wave playback). Some security tools flag workbooks that do.
- 64-bit desktop Excel only: the code uses `LongLong`. No Excel Online, no 32-bit Office.
- Single-threaded, and Excel is busy while the program runs unless it yields with `DoEvents`.
- The translator covers what this one program needed (32- and 64-bit integers, floats, memory operations, a handful of imports). Another program may need more.
- Mac and Windows both run it; only Windows has sound.

## What might carry over to real work

- **Shipping an existing C, C++ or Rust library inside a workbook** where nothing may be installed: a parser, a codec, a compression or hashing routine, a numerical kernel, a simulation, a rules engine. The envelope is roughly a hundred million simple operations a second.
- **Proving the port is right.** The reference-run-and-checksum method works for any deterministic program.
- **Building workbooks from a script.** Sheets, styles, pictures and VBA are all generated, so a workbook can be a reproducible build output under version control.
- **Using the cell grid as a display.** With the batching above, a 320 × 200 grid takes a whole-picture change about 10 times a second and small changes 25 or more times a second: enough for live heat maps, images or dashboards.
- **Keeping bulk data in cells** and loading it into arrays in a fraction of a second.

## Questions worth discussing

1. Which real libraries are small and deterministic enough to be worth carrying this way, and where is "no install, one file" a real requirement rather than a curiosity?
2. Would the same translator target Office Scripts or LibreOffice Basic, and is that easier than the macro-security problem?
3. Is ahead-of-time translation to VBA better than calling a WebAssembly runtime through an add-in, once installation is allowed?
4. How far does the correctness method stretch for programs that are not deterministic or that do real file and network I/O?

## Where things are

| File | Role |
|---|---|
| `build_xl.sh`, `doom/dg_xl.c` | Compile step and the program's platform layer |
| `wasm.py`, `w2v.py` | WebAssembly parser and the translator to VBA |
| `rtvba.py` | VBA runtime and the file-call shim |
| `hostvba.py`, `televba.py` | Display, input, sound, main loop, telemetry |
| `mkxlsm.py`, `mkdoom.py` | Workbook writer and build script |
| `ref.py`, `vbaemu.py`, `emu_run.py`, `vbalint.py` | Reference run, offline VBA emulator, lint |
| `docs/LAB-NOTES.md`, `docs/WINDOWS.md` | Measurements and dead ends in the order they happened |

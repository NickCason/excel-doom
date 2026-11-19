#!/bin/bash
# Build the Excel-target Doom module (8-bit output). OPT and EXTRA can be overridden.
set -e
cd "$(dirname "$0")"
D=vendor/doomgeneric/doomgeneric
SRCS=$(grep "^SRC_DOOM" $D/Makefile | sed 's/SRC_DOOM = //; s/doomgeneric_xlib.o//; s/\.o/.c/g' | tr ' ' '\n' | grep -v '^$' | sed "s|^|$D/|")
zig cc -target wasm32-wasi ${OPT:--O2} -mexec-model=reactor -fno-sanitize=undefined \
  -DCMAP256 -DDOOMGENERIC_RESX=320 -DDOOMGENERIC_RESY=200 -I$D -w \
  -Wno-error=implicit-function-declaration -Wno-error=int-conversion \
  -Wno-error=incompatible-pointer-types -Wno-error=implicit-int \
  ${EXTRA} -o ${OUT:-work/doomxl.wasm} $SRCS doom/dg_xl.c
ls -la ${OUT:-work/doomxl.wasm}

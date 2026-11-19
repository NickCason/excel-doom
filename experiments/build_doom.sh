#!/bin/bash
# Compile doomgeneric + the spike platform layer to a WASI reactor module.
set -e
cd "$(dirname "$0")"
D=vendor/doomgeneric/doomgeneric
SRCS=$(grep "^SRC_DOOM" $D/Makefile | sed 's/SRC_DOOM = //; s/doomgeneric_xlib.o//; s/\.o/.c/g' | tr ' ' '\n' | grep -v '^$' | sed "s|^|$D/|")
zig cc -target wasm32-wasi -O2 -mexec-model=reactor -fno-sanitize=undefined \
  -DDOOMGENERIC_RESX=320 -DDOOMGENERIC_RESY=200 -I$D -w \
  -Wno-error=implicit-function-declaration -Wno-error=int-conversion \
  -Wno-error=incompatible-pointer-types -Wno-error=implicit-int \
  -Wl,--strip-debug ${EXTRA_CFLAGS} -o ${OUT:-build/doom.wasm} $SRCS doom/dg_wasm.c
ls -la ${OUT:-build/doom.wasm}

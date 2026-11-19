/* Platform layer for the spike: virtual clock, no input, frames handed to the host. */
#include <stdint.h>
#include "doomgeneric.h"

__attribute__((import_module("env"), import_name("draw_frame")))
void host_draw_frame(uint32_t *buf, int w, int h);

static uint32_t vclock;

void DG_Init(void) {}
void DG_DrawFrame(void) { host_draw_frame((uint32_t *)DG_ScreenBuffer, DOOMGENERIC_RESX, DOOMGENERIC_RESY); }
void DG_SleepMs(uint32_t ms) { vclock += ms; }
uint32_t DG_GetTicksMs(void) { return vclock; }
int DG_GetKey(int *pressed, unsigned char *key) { (void)pressed; (void)key; return 0; }
void DG_SetWindowTitle(const char *title) { (void)title; }

__attribute__((export_name("dg_start")))
void dg_start(void) {
    static char *argv[] = {"doom", "-iwad", "doom1.wad", 0};
    doomgeneric_Create(3, argv);
}

__attribute__((export_name("dg_tick")))
void dg_tick(void) { vclock += 29; doomgeneric_Tick(); }

/* WASI has no process spawning; Doom only uses this for an error popup. */
int system(const char *cmd) { (void)cmd; return -1; }

/* Spike: let the host choose Doom's view size (3..11) and detail (0 high, 1 low). */
extern int screenblocks, detailLevel;
void R_SetViewSize(int blocks, int detail);
__attribute__((export_name("dg_view")))
void dg_view(int blocks, int detail) { screenblocks = blocks; detailLevel = detail; R_SetViewSize(blocks, detail); }

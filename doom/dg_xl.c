/* Platform layer for the Excel build: 8-bit frames + palette handed to the host,
   clock and keyboard supplied by the host. */
#include <stdint.h>
#include "doomgeneric.h"
#include "d_event.h"
#include "doomstat.h"
#include "d_player.h"

#define IMPORT(n) __attribute__((import_module("env"), import_name(n)))
IMPORT("draw_frame") void host_draw_frame(uint8_t *buf, uint32_t *pal);
IMPORT("ticks_ms") uint32_t host_ticks_ms(void);
IMPORT("sleep_ms") void host_sleep_ms(uint32_t ms);
IMPORT("poll_key") int host_poll_key(void);
IMPORT("mouse_buttons") int host_mouse_buttons(void);
IMPORT("mouse_dx") int host_mouse_dx(void);

struct color { uint32_t b:8; uint32_t g:8; uint32_t r:8; uint32_t a:8; };
extern struct color colors[256];
extern int screenblocks, detailLevel;
void R_SetViewSize(int blocks, int detail);

void DG_Init(void) {}
void DG_DrawFrame(void) { host_draw_frame((uint8_t *)DG_ScreenBuffer, (uint32_t *)colors); }
void DG_SleepMs(uint32_t ms) { host_sleep_ms(ms); }
uint32_t DG_GetTicksMs(void) { return host_ticks_ms(); }
int DG_GetKey(int *pressed, unsigned char *key) {
    int v = host_poll_key();
    if (!v) return 0;
    *pressed = (v >> 8) & 1;
    *key = (unsigned char)(v & 0xff);
    return 1;
}
void DG_SetWindowTitle(const char *title) { (void)title; }
int system(const char *cmd) { (void)cmd; return -1; }

__attribute__((export_name("dg_start")))
void dg_start(int warp) {
    static char *argv_warp[] = {"doom", "-iwad", "doom1.wad", "-warp", "1", "1", 0};
    static char *argv_demo[] = {"doom", "-iwad", "doom1.wad", 0};
    if (warp) doomgeneric_Create(6, argv_warp); else doomgeneric_Create(3, argv_demo);
}

__attribute__((export_name("dg_tick")))
void dg_tick(void) {
    static int lastb;
    int b = host_mouse_buttons();
    int dx = host_mouse_dx();
    if (b != lastb || dx) {
        event_t ev;
        ev.type = ev_mouse; ev.data1 = b; ev.data2 = dx; ev.data3 = 0;
        D_PostEvent(&ev);
        lastb = b;
    }
    doomgeneric_Tick();
}

__attribute__((export_name("dg_view")))
void dg_view(int blocks, int detail) { screenblocks = blocks; detailLevel = detail; R_SetViewSize(blocks, detail); }

/* Telemetry / ground truth. One call per value keeps the host side trivial. */
extern int gametic, leveltime;
__attribute__((export_name("dg_stat")))
int dg_stat(int which) {
    player_t *p = &players[consoleplayer];
    switch (which) {
    case 5: return gametic;
    case 6: return (int)gamestate;          /* 0 level, 1 intermission, 2 finale, 3 demo screen */
    case 7: return gameepisode * 10 + gamemap;
    case 8: return menuactive ? 1 : 0;
    case 14: return paused ? 1 : 0;
    case 12: return leveltime;
    }
    if (!p->mo) return -1;
    switch (which) {
    case 0: return p->ammo[0];
    case 1: return p->health;
    case 2: return (int)(p->mo->angle >> 16);
    case 3: return p->mo->x >> 16;
    case 4: return p->mo->y >> 16;
    case 9: return p->killcount;
    case 10: return p->itemcount;
    case 11: return p->secretcount;
    case 13: return p->armorpoints;
    case 16: return (int)p->playerstate;    /* 0 alive, 1 dead, 2 reborn */
    case 17: return (int)p->readyweapon;
    case 18: return p->ammo[1] + p->ammo[2] * 1000 + p->ammo[3] * 1000000;
    }
    return -1;
}

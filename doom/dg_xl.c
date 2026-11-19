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

/* ---------------------------------------------------------------------------- sound
   One voice. The host is handed the raw 8-bit samples of an effect and plays them on its own
   (Windows: winmm from memory; elsewhere it declines and the game stays silent). There is no
   mixer, so a new effect replaces the one that is playing unless that one matters more
   (lower sfxinfo priority number). Time is counted in game tics, so no host call is needed
   to know whether the voice is free. Music is not played. */
#include "i_sound.h"
#include "w_wad.h"
#include "z_zone.h"
#include "m_misc.h"

IMPORT("play_sound") int host_play_sound(const uint8_t *pcm, int len, int rate, int vol, int sep);

extern int gametic;
int use_libsamplerate = 0;                  /* config variables the sound build expects */
float libsamplerate_scale = 0.65f;

static boolean xl_sfx_prefix;
static int xl_voice = -1;                   /* channel that owns the voice, -1: free */
static int xl_voice_priority;
static int xl_voice_end;                    /* gametic at which the effect is over */

static boolean XL_InitSound(boolean use_sfx_prefix) { xl_sfx_prefix = use_sfx_prefix; return true; }
static void XL_ShutdownSound(void) {}
static void XL_UpdateSound(void) {}
static void XL_UpdateSoundParams(int channel, int vol, int sep) { (void)channel; (void)vol; (void)sep; }

static int XL_GetSfxLumpNum(sfxinfo_t *sfx) {
    char name[9];
    if (sfx->link) sfx = sfx->link;
    M_snprintf(name, sizeof(name), xl_sfx_prefix ? "ds%s" : "%s", sfx->name);
    return W_CheckNumForName(name);
}

static boolean XL_SoundIsPlaying(int channel) {
    return channel >= 0 && channel == xl_voice && gametic < xl_voice_end;
}

static void XL_StopSound(int channel) {
    /* the effect plays out on the host (they are short); the voice is simply up for grabs */
    if (channel >= 0 && channel == xl_voice) xl_voice = -1;
}

static int XL_StartSound(sfxinfo_t *sfx, int channel, int vol, int sep) {
    int lump = sfx->lumpnum, rate, started;
    unsigned int lumplen, len;
    uint8_t *d;
    if (lump < 0 || vol <= 0) return -1;
    if (xl_voice >= 0 && gametic < xl_voice_end && sfx->priority > xl_voice_priority) return -1;
    d = W_CacheLumpNum(lump, PU_STATIC);
    lumplen = W_LumpLength(lump);
    started = 0; len = 0; rate = 11025;
    if (lumplen >= 8 && d[0] == 3 && d[1] == 0) {
        rate = d[2] | (d[3] << 8);
        len = d[4] | (d[5] << 8) | (d[6] << 16) | ((unsigned int)d[7] << 24);
        /* the sample count includes 16 padding bytes at each end */
        if (len <= lumplen - 8 && len > 48 && rate > 0)
            started = host_play_sound(d + 8 + 16, (int)(len - 32), rate, vol, sep);
    }
    W_ReleaseLumpNum(lump);                 /* the host has copied what it needs */
    if (!started) return -1;
    xl_voice = channel;
    xl_voice_priority = sfx->priority;
    xl_voice_end = gametic + (int)((len - 32) * 35 / (unsigned int)rate) + 1;
    return channel;
}

static snddevice_t xl_sound_devices[] = {
    SNDDEVICE_SB, SNDDEVICE_PAS, SNDDEVICE_GUS, SNDDEVICE_WAVEBLASTER, SNDDEVICE_SOUNDCANVAS, SNDDEVICE_AWE32,
};
sound_module_t DG_sound_module = {
    xl_sound_devices, sizeof(xl_sound_devices) / sizeof(*xl_sound_devices),
    XL_InitSound, XL_ShutdownSound, XL_GetSfxLumpNum, XL_UpdateSound, XL_UpdateSoundParams,
    XL_StartSound, XL_StopSound, XL_SoundIsPlaying, NULL,
};

static boolean XL_MusicInit(void) { return true; }
static void XL_MusicNop(void) {}
static void XL_MusicVolume(int volume) { (void)volume; }
static void *XL_MusicRegister(void *data, int len) { (void)data; (void)len; return NULL; }
static void XL_MusicUnregister(void *handle) { (void)handle; }
static void XL_MusicPlay(void *handle, boolean looping) { (void)handle; (void)looping; }
static boolean XL_MusicIsPlaying(void) { return false; }
music_module_t DG_music_module = {
    NULL, 0, XL_MusicInit, XL_MusicNop, XL_MusicVolume, XL_MusicNop, XL_MusicNop,
    XL_MusicRegister, XL_MusicUnregister, XL_MusicPlay, XL_MusicNop, XL_MusicIsPlaying, NULL,
};

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

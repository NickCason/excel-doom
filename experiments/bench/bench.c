/* Spike benchmark: three kernels shaped like Doom's hot code, built with the
   same compiler and flags as the real engine. Each returns a checksum. */
typedef unsigned char byte;
typedef int fixed_t;
typedef long long int64;
#define FRACBITS 16
#define FRACUNIT (1 << FRACBITS)
#define SW 320
#define SH 200
#define VIEWH 168
#define EXPORT(n) __attribute__((export_name(#n)))
#define NOINLINE __attribute__((noinline))

static byte screen[SW * SH];
static byte texture[64][128];
static byte flat[4096];
static byte colormaps[32][256];
byte *ylookup[SH];
int columnofs[SW];
static unsigned rng = 12345;
static NOINLINE unsigned rnd(void) { rng = rng * 1103515245u + 12345u; return (rng >> 16) & 0x7fff; }

NOINLINE fixed_t FixedMul(fixed_t a, fixed_t b) { return ((int64)a * (int64)b) >> FRACBITS; }
NOINLINE fixed_t FixedDiv(fixed_t a, fixed_t b) {
    if (((a < 0 ? -a : a) >> 14) >= (b < 0 ? -b : b)) return (a ^ b) < 0 ? (-2147483647 - 1) : 2147483647;
    return (fixed_t)((((int64)a) << 16) / b);
}

/* ---- renderer kernel: R_DrawColumn / R_DrawSpan as in r_draw.c ---- */
byte *dc_colormap; int dc_x, dc_yl, dc_yh; fixed_t dc_iscale, dc_texturemid; byte *dc_source;
int centery = VIEWH / 2;
byte *ds_colormap, *ds_source; int ds_y, ds_x1, ds_x2; fixed_t ds_xfrac, ds_yfrac, ds_xstep, ds_ystep;

NOINLINE void R_DrawColumn(void) {
    int count; byte *dest; fixed_t frac, fracstep;
    count = dc_yh - dc_yl;
    if (count < 0) return;
    dest = ylookup[dc_yl] + columnofs[dc_x];
    fracstep = dc_iscale;
    frac = dc_texturemid + (dc_yl - centery) * fracstep;
    do {
        *dest = dc_colormap[dc_source[(frac >> FRACBITS) & 127]];
        dest += SW;
        frac += fracstep;
    } while (count--);
}

NOINLINE void R_DrawSpan(void) {
    unsigned int position, step, xtemp, ytemp; byte *dest; int count, spot;
    position = ((ds_xfrac << 10) & 0xffff0000) | ((ds_yfrac >> 6) & 0x0000ffff);
    step = ((ds_xstep << 10) & 0xffff0000) | ((ds_ystep >> 6) & 0x0000ffff);
    dest = ylookup[ds_y] + columnofs[ds_x1];
    count = ds_x2 - ds_x1;
    do {
        ytemp = (position >> 4) & 0x0fc0;
        xtemp = (position >> 26);
        spot = xtemp | ytemp;
        *dest++ = ds_colormap[ds_source[spot]];
        position += step;
    } while (count--);
}

EXPORT(k_init) void k_init(void) {
    int i, j;
    for (i = 0; i < SH; i++) ylookup[i] = screen + i * SW;
    for (i = 0; i < SW; i++) columnofs[i] = i;
    for (i = 0; i < 64; i++) for (j = 0; j < 128; j++) texture[i][j] = (byte)(rnd() ^ (i * 3 + j));
    for (i = 0; i < 4096; i++) flat[i] = (byte)rnd();
    for (i = 0; i < 32; i++) for (j = 0; j < 256; j++) colormaps[i][j] = (byte)((j * (32 - i)) >> 5);
}

EXPORT(k_render) int k_render(int frames) {
    int f, x, y, i; unsigned sum = 0;
    for (f = 0; f < frames; f++) {
        for (y = 0; y < VIEWH; y++) {
            fixed_t dist;
            if (y > 56 && y < 112) continue;
            dist = FixedDiv(40 * FRACUNIT, (y < centery ? centery - y : y - centery + 1) * FRACUNIT);
            ds_y = y; ds_x1 = 0; ds_x2 = SW - 1;
            ds_xstep = FixedMul(dist, 900 + f); ds_ystep = FixedMul(dist, 300 - f);
            ds_xfrac = f * 4096 + dist; ds_yfrac = dist * 3;
            ds_source = flat; ds_colormap = colormaps[(dist >> 18) & 31];
            R_DrawSpan();
        }
        for (x = 0; x < SW; x++) {
            fixed_t dist = FRACUNIT * (8 + ((x * 7 + f * 13) & 63));
            fixed_t scale = FixedDiv(100 * FRACUNIT, dist);
            int h = FixedMul(scale, 6 * FRACUNIT) >> FRACBITS;
            int yl = centery - h, yh = centery + h;
            if (yl < 0) yl = 0;
            if (yh >= VIEWH) yh = VIEWH - 1;
            dc_x = x; dc_yl = yl; dc_yh = yh;
            dc_iscale = FixedDiv(FRACUNIT, scale) * 16;
            dc_texturemid = 64 * FRACUNIT;
            dc_source = texture[(x + f) & 63];
            dc_colormap = colormaps[(dist >> 18) & 31];
            R_DrawColumn();
        }
    }
    for (i = 0; i < SW * SH; i++) sum = sum * 31 + screen[i];
    return (int)sum;
}

/* ---- game-logic kernel: linked structs, function pointers, a switch ---- */
typedef struct mobj_s {
    struct mobj_s *next; fixed_t x, y, z, momx, momy, momz;
    int flags, health, state, tics; void (*action)(struct mobj_s *);
    short type; byte pad; byte frame;
} mobj_t;
static mobj_t mobjs[256];
static mobj_t *mobjhead;
static int blockmap[32 * 32];
static byte rndtable[256]; static int prndindex;
static NOINLINE int P_Random(void) { prndindex = (prndindex + 1) & 0xff; return rndtable[prndindex]; }
static void A_Chase(mobj_t *m) { m->momx += (P_Random() - 128) << 8; m->momy += (P_Random() - 128) << 8; m->frame++; }
static void A_Fall(mobj_t *m) { m->momz -= FRACUNIT; if (m->z < 0) { m->z = 0; m->momz = -m->momz >> 1; } }
static void A_Look(mobj_t *m) { if (P_Random() < 40) m->flags ^= 4; m->type = (short)(m->type + m->frame); }
static void (*const actions[4])(mobj_t *) = {A_Chase, A_Fall, A_Look, A_Chase};

EXPORT(k_logic) int k_logic(int ticks) {
    int i, t; unsigned sum = 0; mobj_t *m;
    for (i = 0; i < 256; i++) rndtable[i] = (byte)rnd();
    for (i = 0; i < 1024; i++) blockmap[i] = (rnd() & 7) == 0;
    mobjhead = 0;
    for (i = 0; i < 256; i++) {
        m = &mobjs[i];
        m->x = (rnd() & 1023) << FRACBITS; m->y = (rnd() & 1023) << FRACBITS; m->z = 0;
        m->momx = m->momy = m->momz = 0; m->flags = i & 3; m->health = 100; m->state = i & 7;
        m->tics = 1 + (i & 3); m->action = actions[i & 3]; m->type = (short)i; m->frame = 0;
        m->next = mobjhead; mobjhead = m;
    }
    for (t = 0; t < ticks; t++) {
        for (m = mobjhead; m; m = m->next) {
            int bx, by;
            if (--m->tics <= 0) {
                switch (m->state & 7) {
                case 0: m->tics = 3; m->state = 1; break;
                case 1: m->tics = 2; m->state = 4; m->health--; break;
                case 2: m->tics = 5; m->state = 3; break;
                case 3: m->tics = 1; m->state = 0; m->flags |= 8; break;
                case 4: m->tics = 4; m->state = 6; break;
                case 5: m->tics = 2; m->state = 2; m->health += 2; break;
                case 6: m->tics = 3; m->state = 7; break;
                default: m->tics = 6; m->state = 5; break;
                }
                m->action(m);
            }
            m->x += m->momx; m->y += m->momy; m->z += m->momz;
            m->momx = FixedMul(m->momx, 0xE800); m->momy = FixedMul(m->momy, 0xE800);
            bx = (m->x >> 21) & 31; by = (m->y >> 21) & 31;
            if (blockmap[by * 32 + bx]) { m->momx = -m->momx; m->momy = -m->momy; m->health--; }
            if (m->health <= 0) { m->health = 100; m->x = (P_Random() << 18); }
        }
    }
    for (m = mobjhead; m; m = m->next)
        sum = sum * 31 + (unsigned)(m->x ^ m->y ^ m->z ^ m->health ^ m->flags ^ m->type ^ m->frame);
    return (int)sum;
}

/* ---- BSP kernel: recursive tree walk with R_PointOnSide ---- */
typedef struct { fixed_t x, y, dx, dy; int children[2]; } node_t;
#define NNODES 1023
static node_t nodes[NNODES];
static fixed_t viewx, viewy; static unsigned bspsum;

static NOINLINE int R_PointOnSide(fixed_t x, fixed_t y, node_t *node) {
    fixed_t dx, dy, left, right;
    if (!node->dx) { if (x <= node->x) return node->dy > 0; return node->dy < 0; }
    if (!node->dy) { if (y <= node->y) return node->dx < 0; return node->dx > 0; }
    dx = (x - node->x); dy = (y - node->y);
    if ((node->dy ^ node->dx ^ dx ^ dy) & 0x80000000) { if ((node->dy ^ dx) & 0x80000000) return 1; return 0; }
    left = FixedMul(node->dy >> FRACBITS, dx); right = FixedMul(dy, node->dx >> FRACBITS);
    if (right < left) return 0;
    return 1;
}
static NOINLINE void R_RenderBSPNode(int bspnum) {
    node_t *bsp; int side;
    if (bspnum & 0x8000) { bspsum = bspsum * 31 + (unsigned)(bspnum & 0x7fff); return; }
    bsp = &nodes[bspnum];
    side = R_PointOnSide(viewx, viewy, bsp);
    R_RenderBSPNode(bsp->children[side]);
    if (((bsp->x ^ viewx) & 0x30000) != 0) R_RenderBSPNode(bsp->children[side ^ 1]);
}
EXPORT(k_bsp) int k_bsp(int iters) {
    int i;
    for (i = 0; i < NNODES; i++) {
        nodes[i].x = (rnd() - 16384) << 12; nodes[i].y = (rnd() - 16384) << 12;
        nodes[i].dx = (i % 5 == 0) ? 0 : (rnd() - 16384) << 10; nodes[i].dy = (i % 7 == 0) ? 0 : (rnd() - 16384) << 10;
        nodes[i].children[0] = (2 * i + 1 < NNODES) ? 2 * i + 1 : (0x8000 | i);
        nodes[i].children[1] = (2 * i + 2 < NNODES) ? 2 * i + 2 : (0x8000 | (i ^ 0x155));
    }
    bspsum = 0;
    for (i = 0; i < iters; i++) { viewx = (i * 7919) << 12; viewy = (i * 104729) << 10; R_RenderBSPNode(0); }
    return (int)bspsum;
}

/* the compiler may synthesize these */
void *memset(void *d, int c, unsigned long n) { byte *p = d; while (n--) *p++ = (byte)c; return d; }
void *memcpy(void *d, const void *s, unsigned long n) { byte *p = d; const byte *q = s; while (n--) *p++ = *q++; return d; }

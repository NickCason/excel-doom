"""Offline simulation of display tricks on real demo frames. For each scheme: how many
fill calls per displayed frame, and how close the sheet would look to the true frame.
Time estimate uses the measured Excel costs: ~9 us per fill call + ~85 ms repaint."""
import os, sys, pickle
import numpy as np
from wasmtime import Engine, Store, Module, Linker, WasiConfig, FuncType, ValType

H, W = 200, 320
CALL_US, PAINT_MS = 9.0, 85.0

def capture(screenblocks, detail, ticks):
    key = 'out/frames_sb%d_d%d.pkl' % (screenblocks, detail)
    if os.path.exists(key):
        return pickle.load(open(key, 'rb'))
    d = 'out/run_sb%d_d%d' % (screenblocks, detail)
    os.makedirs(d, exist_ok=True)
    if not os.path.exists(d + '/doom1.wad'):
        os.link('vendor/wad/doom1.wad', d + '/doom1.wad')
    open(d + '/.default.cfg', 'w').write('screenblocks %d\ndetaillevel %d\n' % (screenblocks, detail))
    engine = Engine(); store = Store(engine)
    wasi = WasiConfig(); wasi.preopen_dir(d, '.'); store.set_wasi(wasi)
    linker = Linker(engine); linker.define_wasi()
    last = {}
    def draw(ptr, w, h):
        last['f'] = bytes(last['mem'].read(store, ptr, ptr + w * h * 4))
    linker.define_func('env', 'draw_frame', FuncType([ValType.i32()] * 3, []), draw)
    inst = linker.instantiate(store, Module.from_file(engine, 'build/doom.wasm'))
    ex = inst.exports(store); last['mem'] = ex['memory']
    ex['_initialize'](store); ex['dg_start'](store)
    ex['dg_view'](store, screenblocks, detail)
    out = []
    for i in range(max(ticks) + 1):
        ex['dg_tick'](store)
        if i in ticks:
            a = np.frombuffer(last['f'], dtype=np.uint8).reshape(H, W, 4)
            out.append(np.ascontiguousarray(a[:, :, [2, 1, 0]]))
    pickle.dump(out, open(key, 'wb'))
    return out

def pack(a):
    a = a.astype(np.uint32)
    return a[..., 0] | (a[..., 1] << 8) | (a[..., 2] << 16)

def unpack(c):
    return (c & 255, (c >> 8) & 255, (c >> 16) & 255)

def exact_runs(disp, new, rows=None):
    pn, pd = pack(new), pack(disp)
    h, w = pn.shape
    out = []
    for y in (range(h) if rows is None else rows):
        idx = np.flatnonzero(pn[y] != pd[y])
        if len(idx) == 0:
            continue
        row = pn[y].tolist(); idx = idx.tolist()
        i = 0; n = len(idx)
        while i < n:
            x0 = idx[i]; c = row[x0]; x1 = x0
            while x1 + 1 < w and row[x1 + 1] == c:
                x1 += 1
            out.append((y, x0, x1, c))
            while i < n and idx[i] <= x1:
                i += 1
    return out

def lossy_runs(disp, new, T, rows=None):
    h, w, _ = new.shape
    ni = new.astype(np.int32); di = disp.astype(np.int32)
    need_all = np.abs(ni - di).max(axis=2) > T
    out = []
    for y in (range(h) if rows is None else rows):
        if not need_all[y].any():
            continue
        need = need_all[y].tolist(); rl = ni[y].tolist()
        x = 0
        while x < w:
            if need[x]:
                r0, g0, b0 = rl[x]; x1 = x
                sr, sg, sb = r0, g0, b0
                while x1 + 1 < w:
                    r, g, b = rl[x1 + 1]
                    if abs(r - r0) > T or abs(g - g0) > T or abs(b - b0) > T:
                        break
                    x1 += 1; sr += r; sg += g; sb += b
                n = x1 - x + 1
                c = int(round(sr / n)) | (int(round(sg / n)) << 8) | (int(round(sb / n)) << 16)
                out.append((y, x, x1, c))
                x = x1 + 1
            else:
                x += 1
    return out

def rect_merge(runs, tol=0):
    """Join runs that line up exactly in consecutive rows (same span, same/near colour)."""
    rects = []; prev = {}; cur = {}; cury = None
    for (y, x0, x1, c) in runs:
        if y != cury:
            prev = cur if (cury is not None and y == cury + 1) else {}
            cur = {}; cury = y
        k = (x0, x1)
        hit = prev.get(k)
        if hit is not None:
            ri, c0 = hit
            if c0 == c or (tol and max(abs(a - b) for a, b in zip(unpack(c0), unpack(c))) <= tol):
                rects[ri][3] = y; cur[k] = (ri, c0); continue
        rects.append([x0, y, x1, y, c]); cur[k] = (len(rects) - 1, c)
    return rects

def apply(disp, rects):
    for x0, y0, x1, y1, c in rects:
        disp[y0:y1 + 1, x0:x1 + 1] = unpack(c)

def psnr(a, b):
    mse = ((a.astype(np.float64) - b.astype(np.float64)) ** 2).mean()
    return 99.0 if mse == 0 else 10 * np.log10(255 * 255 / mse)

def top_budget(disp, new, rects, budget):
    if len(rects) <= budget:
        return rects
    err = np.abs(new.astype(np.int32) - disp.astype(np.int32)).sum(axis=2)
    ii = np.zeros((err.shape[0] + 1, err.shape[1] + 1), dtype=np.int64)
    ii[1:, 1:] = err.cumsum(0).cumsum(1)
    pr = [ii[y1 + 1, x1 + 1] - ii[y0, x1 + 1] - ii[y1 + 1, x0] + ii[y0, x0] for x0, y0, x1, y1, c in rects]
    order = np.argsort(pr)[::-1][:budget]
    return [rects[i] for i in order]

def quad_encode(new):
    """2x2 blocks -> two colours + a 4-bit shape (quadrant block characters)."""
    b = new.reshape(H // 2, 2, W // 2, 2, 3).transpose(0, 2, 1, 3, 4).reshape(H // 2, W // 2, 4, 3).astype(np.float64)
    lum = b @ np.array([0.3, 0.59, 0.11])
    m = lum > lum.mean(axis=2, keepdims=True) + 2.0
    cnt = m.sum(axis=2)
    mean = b.mean(axis=2)
    fg = np.where(cnt[..., None] > 0, (b * m[..., None]).sum(2) / np.maximum(cnt, 1)[..., None], mean)
    bg = (b * (~m)[..., None]).sum(2) / np.maximum(4 - cnt, 1)[..., None]
    fg = fg.round().astype(np.uint8); bg = bg.round().astype(np.uint8)
    return fg, bg, m

def quad_render(fg, bg, m):
    blk = np.where(m[..., None], fg[:, :, None, :], bg[:, :, None, :])
    return blk.reshape(H // 2, W // 2, 2, 2, 3).transpose(0, 2, 1, 3, 4).reshape(H, W, 3)

def simulate(name, frames, mode, **kw):
    T = kw.get('T', 0); budget = kw.get('budget'); rect = kw.get('rect', True)
    disp = frames[0].copy()
    calls = []; q = []; shots = {}; rects_list = []
    if mode == 'half':
        dsmall = frames[0][::2, ::2].copy()
    if mode == 'quad':
        dfg, dbg, dm = quad_encode(frames[0])
    for k, new in enumerate(frames[1:], 1):
        if mode == 'half':
            ns = new[::2, ::2]
            runs = lossy_runs(dsmall, ns, T) if T else exact_runs(dsmall, ns)
            rects = rect_merge(runs, T) if rect else [[x0, y, x1, y, c] for y, x0, x1, c in runs]
            apply(dsmall, rects)
            shown = dsmall.repeat(2, axis=0).repeat(2, axis=1)
        elif mode == 'quad':
            fg, bg, m = quad_encode(new)
            fgt = np.where(m.any(axis=2)[..., None], fg, dfg)      # fg is don't-care where no shape
            r1 = rect_merge(lossy_runs(dfg, fgt, T), T); r2 = rect_merge(lossy_runs(dbg, bg, T), T)
            apply(dfg, r1); apply(dbg, r2); dm = m
            rects = r1 + r2
            shown = quad_render(dfg, dbg, dm)
        else:
            rows = None
            if mode == 'interlace':
                rows = range(k % 2, H, 2)
            runs = lossy_runs(disp, new, T, rows) if T else exact_runs(disp, new, rows)
            rects = rect_merge(runs, T) if rect else [[x0, y, x1, y, c] for y, x0, x1, c in runs]
            if budget:
                rects = top_budget(disp, new, rects, budget)
            apply(disp, rects)
            shown = disp
        rects_list.append([list(r) for r in rects]); calls.append(len(rects)); q.append(psnr(shown, kw.get('truth', frames)[k] if 'truth' in kw else new))
        shots[k] = shown.copy()
    return dict(name=name, calls=float(np.mean(calls)), p95=float(np.percentile(calls, 95)),
                psnr=float(np.mean(q)), worst=float(np.min(q)), shots=shots, calls_list=calls, rects=rects_list,
                median=float(np.median(calls)))

if __name__ == '__main__':
    ticks = list(range(300, 541, 8))
    base = capture(10, 0, ticks)
    low = capture(10, 1, ticks)
    sb8 = capture(8, 0, ticks)
    sb6 = capture(6, 0, ticks)
    sb8low = capture(8, 1, ticks)
    R = []
    R.append(simulate('full res, changed runs (measured baseline)', base, 'full', rect=False))
    R.append(simulate('+ join runs into rectangles', base, 'full'))
    R.append(simulate('Doom low-detail mode', low, 'full', truth=base))
    R.append(simulate('near-colour merging, tolerance 8', base, 'full', T=8))
    R.append(simulate('near-colour merging, tolerance 16', base, 'full', T=16))
    R.append(simulate('near-colour merging, tolerance 24', base, 'full', T=24))
    R.append(simulate('interlaced (alternate rows each frame)', base, 'interlace'))
    R.append(simulate('budget 6000 calls, worst-first (tol 8)', base, 'full', T=8, budget=6000))
    R.append(simulate('budget 3000 calls, worst-first (tol 8)', base, 'full', T=8, budget=3000))
    R.append(simulate('plain half res 160x100', base, 'half'))
    R.append(simulate('half res + 2-colour quadrant glyphs (tol 8)', base, 'quad', T=8))
    R.append(simulate('low detail + tolerance 16', low, 'full', T=16, truth=base))
    R.append(simulate('low detail + tol 16 + budget 3000', low, 'full', T=16, budget=3000, truth=base))
    R.append(simulate('small view (size 8 of 10)', sb8, 'full'))
    R.append(simulate('small view (size 6 of 10)', sb6, 'full'))
    R.append(simulate('small view 8 + low detail + tol 16', sb8low, 'full', T=16))
    print('%-46s %8s %8s %8s %9s %7s %7s' % ('scheme', 'calls', 'median', 'p95', 'est ms', 'PSNR', 'worst'))
    for r in R:
        est = r['calls'] * CALL_US / 1000 + PAINT_MS
        print('%-46s %8.0f %8.0f %8.0f %9.0f %7.1f %7.1f' % (r['name'], r['calls'], r['median'], r['p95'], est, r['psnr'], r['worst']))
    print('baseline calls per frame:', R[0]['calls_list'])
    pickle.dump(dict(R=R, base=base, low=low, sb8=sb8, sb6=sb6, ticks=ticks), open('out/dispsim.pkl', 'wb'))

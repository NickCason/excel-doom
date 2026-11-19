"""Run the translated Doom VBA in the offline emulator and compare every frame with
the wasmtime reference (same deterministic clock, no input)."""
import sys, time, struct, zlib, pickle, traceback, re
import w2v, rtvba, vbaemu, wasm

def build_modules(wasm_path='work/doomxl.wasm', checked=False):
    T = w2v.Translator(wasm_path, checked=checked)
    mods, stats = T.build()
    return T, mods + [('Rt', rtvba.RT), ('Wasi', rtvba.WASI)], stats

def run(nticks=3, checked=True, strict=True, check_index=True, warp=1, verbose=True):
    sys.setrecursionlimit(50000)
    t0 = time.time()
    T, mods, stats = build_modules(checked=checked)
    st = dict(clock=0, frames=[])
    def draw(buf, pal):
        M = emu.ns['M']
        st['frames'].append((struct.pack('<16000i', *M[buf // 4:buf // 4 + 16000]), struct.pack('<256i', *M[pal // 4:pal // 4 + 256])))
    def ticks(): return st['clock'] & 0x7FFFFFFF
    def sleep(ms): st['clock'] += ms
    natives = {'imp_draw_frame': (draw, ['L', 'L'], None), 'imp_ticks_ms': (ticks, [], 'L'),
               'imp_sleep_ms': (sleep, ['L'], None), 'imp_poll_key': (lambda: 0, [], 'L'),
               'imp_mouse_buttons': (lambda: 0, [], 'L'), 'imp_mouse_dx': (lambda: 0, [], 'L')}
    emu = vbaemu.Emu(mods, natives, strict=strict, check_index=check_index)
    if verbose: print('translate+compile %.1fs (python source %.1f MB)' % (time.time() - t0, emu.pysrc_len / 1e6), flush=True)
    wad = open('vendor/wad/doom1.wad', 'rb').read()
    pad = wad + b'\0' * (-len(wad) % 4)
    emu.ns['WD'] = list(struct.unpack('<%di' % (len(pad) // 4), pad)); emu.ns['WDLen'] = len(wad)
    names = T.m.names
    def out_text():
        n = emu.ns['OutN']; b = emu.ns['OutB']
        return bytes(b[i & 65535] for i in range(max(0, n - 65536), n)).decode('latin1')
    try:
        t0 = time.time()
        emu.call('RtInit', 512); emu.call('ProgInit'); emu.call('WasiInit')
        emu.call('x__initialize')
        emu.call('x_dg_start', warp)
        if verbose: print('boot %.1fs, frames so far %d' % (time.time() - t0, len(st['frames'])), flush=True)
        for t in range(nticks):
            st['clock'] += 29
            t1 = time.time()
            emu.call('x_dg_tick')
            if verbose and t < 3: print('tick %d: %.1fs' % (t, time.time() - t1), flush=True)
    except Exception as e:
        tb = traceback.extract_tb(sys.exc_info()[2])
        chain = [f.name for f in tb if f.filename == '<vba>']
        def nm(f):
            m = re.fullmatch(r'f(\d+)', f)
            return '%s(%s)' % (f, names.get(int(m.group(1)), '?')) if m else f
        print('ERROR:', type(e).__name__, str(e)[:300])
        print('call chain (innermost last):', ' > '.join(nm(f) for f in chain[-12:]))
        last = [f for f in tb if f.filename == '<vba>'][-1:]
        print('python line', last[0].lineno if last else None)
        st['error'] = e
    st['out'] = out_text()
    return st, emu

if __name__ == '__main__':
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 3
    flags = sys.argv[2] if len(sys.argv) > 2 else 'csi'
    st, emu = run(n, checked='c' in flags, strict='s' in flags, check_index='i' in flags)
    ref = pickle.load(open('work/ref_frames.pkl', 'rb'))
    fr = st['frames']
    same = 0
    for a, b in zip(fr, ref):
        if a != b: break
        same += 1
    print('frames produced %d, reference has %d, identical prefix %d' % (len(fr), len(ref), same))
    if 'error' in st or same < min(len(fr), len(ref)):
        print('--- program output tail ---'); print(st['out'][-700:])

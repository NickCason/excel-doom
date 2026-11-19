"""Reference run of the Excel-target Doom module under wasmtime with a deterministic
virtual clock and scripted keys. Produces per-frame checksums (and raw frames)."""
import sys, zlib, pickle, struct
from wasmtime import Config, Engine, Store, Module, Linker, WasiConfig, FuncType, ValType

def run(wasm='work/doomxl.wasm', nticks=40, warp=1, keys=None, fuel=False):
    cfg = Config(); cfg.consume_fuel = fuel
    engine = Engine(cfg); store = Store(engine)
    if fuel: store.set_fuel(1 << 62)
    wasi = WasiConfig(); wasi.preopen_dir('vendor/wad', '.'); store.set_wasi(wasi)
    linker = Linker(engine); linker.define_wasi()
    st = dict(clock=0, frames=[], q=[])
    i32 = ValType.i32()
    def draw(buf, pal):
        mem = st['mem']
        st['frames'].append((bytes(mem.read(store, buf, buf + 64000)), bytes(mem.read(store, pal, pal + 1024))))
    def ticks(): return st['clock'] & 0x7FFFFFFF
    def sleep(ms): st['clock'] += ms
    def poll(): return st['q'].pop(0) if st['q'] else 0
    linker.define_func('env', 'draw_frame', FuncType([i32, i32], []), draw)
    linker.define_func('env', 'ticks_ms', FuncType([], [i32]), ticks)
    linker.define_func('env', 'sleep_ms', FuncType([i32], []), sleep)
    linker.define_func('env', 'poll_key', FuncType([], [i32]), poll)
    linker.define_func('env', 'mouse_buttons', FuncType([], [i32]), lambda: 0)
    linker.define_func('env', 'mouse_dx', FuncType([], [i32]), lambda: 0)
    inst = linker.instantiate(store, Module.from_file(engine, wasm))
    ex = inst.exports(store); st['mem'] = ex['memory']
    ex['_initialize'](store)
    ex['dg_start'](store, warp)
    boot_frames = len(st['frames'])
    for t in range(nticks):
        st['clock'] += 29
        for (tick, ev) in (keys or []):
            if tick == t: st['q'].append(ev)
        ex['dg_tick'](store)
    return st['frames'], boot_frames, ex['memory'].size(store)

def crc(fr):
    return zlib.crc32(fr[0] + fr[1])

def png(path, fr, scale=2):
    from PIL import Image
    idx, pal = fr
    rgb = bytearray()
    for i in range(256):
        b, g, r, a = pal[i * 4:i * 4 + 4]; rgb += bytes((r, g, b))
    im = Image.frombytes('P', (320, 200), idx); im.putpalette(bytes(rgb))
    im.convert('RGB').resize((320 * scale, 200 * scale), Image.NEAREST).save(path)

if __name__ == '__main__':
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 40
    frames, boot, pages = run(nticks=n)
    print('frames', len(frames), 'boot frames', boot, 'pages', pages)
    print('crcs:', ' '.join('%08x' % crc(f) for f in frames[:12]))
    pickle.dump(frames, open('work/ref_frames.pkl', 'wb'))
    png('work/ref_last.png', frames[-1])

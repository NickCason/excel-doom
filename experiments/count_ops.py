"""Run the WebAssembly Doom under wasmtime with fuel metering: how many
operations does boot take, and how many per frame of the attract-mode demo?"""
import sys, struct, zlib, os
from wasmtime import Config, Engine, Store, Module, Linker, WasiConfig, FuncType, ValType

wasm = sys.argv[1] if len(sys.argv) > 1 else 'build/doom.wasm'
nticks = int(sys.argv[2]) if len(sys.argv) > 2 else 600
cfg = Config(); cfg.consume_fuel = True
engine = Engine(cfg)
store = Store(engine)
BIG = 1 << 62
store.set_fuel(BIG)
wasi = WasiConfig(); wasi.preopen_dir('vendor/wad', '.'); wasi.inherit_stdout(); wasi.inherit_stderr()
store.set_wasi(wasi)
linker = Linker(engine); linker.define_wasi()
frames = []
state = {}
def draw_frame(ptr, w, h):
    mem = state['mem']
    frames.append(bytes(mem.read(store, ptr, ptr + w * h * (1 if os.environ.get('NOPNG') else 4))))
linker.define_func('env', 'draw_frame', FuncType([ValType.i32(), ValType.i32(), ValType.i32()], []), draw_frame)
module = Module.from_file(engine, wasm)
inst = linker.instantiate(store, module)
ex = inst.exports(store)
state['mem'] = ex['memory']
def used():
    return BIG - store.get_fuel()

def png(path, raw, w=320, h=200):
    rows = b''
    for y in range(h):
        row = raw[y * w * 4:(y + 1) * w * 4]
        rgb = bytearray()
        for x in range(w):
            b, g, r, _ = row[x * 4:x * 4 + 4]
            rgb += bytes((r, g, b))
        rows += b'\0' + bytes(rgb)
    def chunk(t, d):
        return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d))
    open(path, 'wb').write(b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0))
                           + chunk(b'IDAT', zlib.compress(rows)) + chunk(b'IEND', b''))

ex['_initialize'](store)
f0 = used()
ex['dg_start'](store)
boot = used() - f0
print('memory pages: %d (%.1f MB)' % (ex['memory'].size(store), ex['memory'].size(store) * 65536 / 1e6))
print('BOOT ops: %.1f M   (frames drawn during boot: %d)' % (boot / 1e6, len(frames)))
per = []
for i in range(nticks):
    a = used(); nf = len(frames)
    ex['dg_tick'](store)
    per.append((used() - a, len(frames) - nf))
    if i in (5, 120, 300, 450, nticks - 1) and frames and not os.environ.get('NOPNG'):
        png('out/frame%04d.png' % i, frames[-1])
ops = [p for p, _ in per]
drawn = sum(n for _, n in per)
def stats(xs):
    s = sorted(xs)
    return 'min %.2fM  median %.2fM  mean %.2fM  p95 %.2fM  max %.2fM' % (
        s[0] / 1e6, s[len(s) // 2] / 1e6, sum(s) / len(s) / 1e6, s[int(len(s) * .95)] / 1e6, s[-1] / 1e6)
print('ticks: %d, frames drawn: %d' % (nticks, drawn))
print('ops per tick, all        :', stats(ops))
for lo, hi in [(0, 60), (60, 200), (200, 400), (400, nticks)]:
    seg = ops[lo:hi]
    if seg:
        print('ops per tick, %4d-%4d  :' % (lo, hi), stats(seg))
print('first 12 ticks (M ops):', ' '.join('%.2f' % (o / 1e6) for o in ops[:12]))

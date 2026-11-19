"""Run the benchmark workbooks in Excel and compare with the wasmtime reference."""
import sys
from wasmtime import Config, Engine, Store, Module, Instance
import xl, mkbench

cfg = Config(); cfg.consume_fuel = True
e = Engine(cfg); s = Store(e); BIG = 1 << 62; s.set_fuel(BIG)
ex = Instance(s, Module.from_file(e, 'build/bench.wasm'), []).exports(s)
ex['k_init'](s)
ref = {}
for name, n in zip(['render', 'logic', 'bsp'], mkbench.N):
    a = s.get_fuel(); r = ex['k_' + name](s, n); ref[name] = (r, a - s.get_fuel())
print('reference:', {k: 'checksum %d, %.1fM ops' % (v[0], v[1] / 1e6) for k, v in ref.items()})
tags = sys.argv[1:] or ['L', 'B', 'Lh', 'I']
labels = {'L': 'translated, Long-array memory', 'B': 'translated, Byte-array memory',
          'Lh': 'translated, Long-array, byte access via helper calls', 'I': 'emulated (interpreter in VBA)'}
for tag in tags:
    out, dt = xl.session('out/bench%s.xlsm' % tag, 'Driver.Go', timeout=900)
    print('%s  [%s]  wall %.1fs' % (tag, labels[tag], dt))
    if 'ERROR' in out or out.startswith(('DIALOG', 'TIMEOUT')) or '!ERR' in out:
        print('    FAILED:', out[:400]); continue
    for part in out.split(';'):
        if not part: continue
        name, chk, secs = part.split(',')
        rchk, rops = ref[name]
        secs = float(secs)
        ok = 'OK ' if int(chk) == rchk else 'MISMATCH (got %s want %d)' % (chk, rchk)
        print('    %-7s checksum %s %7.3fs  %6.1f M ops/s  %5.1f ns/op' % (name, ok, secs, rops / secs / 1e6 if secs else 0, secs / rops * 1e9))

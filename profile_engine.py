"""Profile the translated engine offline: run it in the VBA emulator and report, per VBA
procedure, how many times it is called per frame and its share of executed time."""
import sys, time, struct, cProfile, pstats, re, io
import w2v, rtvba, vbaemu, emu_run

def main(nticks=20, warp=1):
    sys.setrecursionlimit(50000)
    T, mods, stats = emu_run.build_modules(checked=False)
    st = dict(clock=0)
    natives = {'imp_draw_frame': (lambda b, p: None, ['L', 'L'], None), 'imp_ticks_ms': (lambda: st['clock'] & 0x7FFFFFFF, [], 'L'),
               'imp_sleep_ms': (lambda ms: st.__setitem__('clock', st['clock'] + ms), ['L'], None), 'imp_poll_key': (lambda: 0, [], 'L'),
               'imp_mouse_buttons': (lambda: 0, [], 'L'), 'imp_mouse_dx': (lambda: 0, [], 'L')}
    emu = vbaemu.Emu(mods, natives, strict=False, check_index=False)
    wad = open('vendor/wad/doom1.wad', 'rb').read(); pad = wad + b'\0' * (-len(wad) % 4)
    emu.ns['WD'] = list(struct.unpack('<%di' % (len(pad) // 4), pad)); emu.ns['WDLen'] = len(wad)
    emu.call('RtInit', 512); emu.call('ProgInit'); emu.call('WasiInit'); emu.call('x__initialize'); emu.call('x_dg_start', warp)
    for _ in range(5):
        st['clock'] += 29; emu.call('x_dg_tick')
    pr = cProfile.Profile(); pr.enable()
    for _ in range(nticks):
        st['clock'] += 29; emu.call('x_dg_tick')
    pr.disable()
    ps = pstats.Stats(pr)
    names = T.m.names
    rows = []
    total = 0.0
    for (fn, line, name), (cc, nc, tt, ct, callers) in ps.stats.items():
        if fn != '<vba>':
            continue
        m = re.fullmatch(r'f(\d+)', name)
        label = names.get(int(m.group(1)), name) if m else name
        rows.append((tt, nc, label, name)); total += tt
    rows.sort(reverse=True)
    ncalls = sum(r[1] for r in rows)
    print('procedure calls per frame: %.0f total  (at ~0.22 us each in VBA that is ~%.1f ms of call overhead per frame)' % (ncalls / nticks, ncalls / nticks * 0.22e-3))
    print('%-28s %10s %8s' % ('procedure', 'calls/frame', 'time %'))
    for tt, nc, label, name in rows[:22]:
        print('%-28s %10.0f %7.1f%%' % (label[:28], nc / nticks, 100 * tt / total))
    byc = sorted(rows, key=lambda r: -r[1])[:12]
    print('most-called:', ', '.join('%s %.0f' % (r[2][:20], r[1] / nticks) for r in byc))

if __name__ == '__main__':
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 20)

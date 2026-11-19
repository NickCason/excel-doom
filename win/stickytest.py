"""Windows: are the Sticky Keys and Filter Keys shortcuts off while a game runs, and back afterwards?

Shift is the run key. Five quick taps of it, or eight seconds on the right-hand one, make Windows
offer Sticky Keys or Filter Keys with a prompt over the game, so the game switches both shortcuts
off while it runs. This reads the two settings before, during and after a short game.

It only reads the settings; the game is what changes them. Injected Shift taps do not raise the
prompt (Windows ignores injected keys for it), so the prompt itself is not tested here: tap Shift
five times in a real game for that.

Run from the repo root:  .venv\\Scripts\\python win\\stickytest.py
Ends with STICKY TEST PASSED, or says what was wrong and exits 1.
"""
import ctypes, os, subprocess, sys, time
from ctypes import wintypes
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import playtest as pt

RESULT = os.path.join(pt.OUT, 'sticky_result.txt')
ON, SHORTCUT, ASKS = 1, 4, 8

def flags(get, size):
    b = (wintypes.DWORD * (size // 4))(size)
    if not pt.u32.SystemParametersInfoW(get, size, b, 0):
        raise OSError('SystemParametersInfo(%#x) failed' % get)
    return b[1]

def state(tag):
    s = flags(0x3A, 8), flags(0x32, 24)
    words = lambda f: ', '.join(n for bit, n in ((ON, 'feature ON'), (SHORTCUT, 'shortcut active'), (ASKS, 'asks first')) if f & bit) or 'off'
    print('%-7s Sticky Keys: %-28s | Filter Keys: %s' % (tag, words(s[0]), words(s[1])), flush=True)
    return s

def main():
    pt.release_all()
    before = state('before')
    if os.path.exists(RESULT):
        os.remove(RESULT)
    known = {h for h, c, t, p in pt.windows() if c == 'XLMAIN'}
    ps = subprocess.Popen(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', os.path.join('win', 'verify.ps1'),
                           '-Macro', 'Host.Play', '-Out', RESULT, '-Maximize', '-Linger', '3'])
    hwnd = pid = None; during = None; t0 = time.time()
    while hwnd is None and time.time() - t0 < 90 and ps.poll() is None:
        time.sleep(0.5); hwnd, pid = pt.excel_window(known)
    try:
        if hwnd is None:
            print('no Excel window with the workbook appeared'); return 1
        time.sleep(7); pt.bring_front(hwnd)
        during = state('during')
        for _ in range(8):
            pt.hold(hwnd, ['q'], 0.12); time.sleep(0.5)
            if ps.poll() is not None or os.path.exists(RESULT):
                break
        ps.wait(timeout=60)
    finally:
        print('keys left down:', pt.release_all())
        if ps.poll() is None and pid:
            subprocess.run(['taskkill', '/PID', str(pid), '/F'], capture_output=True)
    after = state('after')
    out = open(RESULT, encoding='utf-8-sig').read() if os.path.exists(RESULT) else ''
    f = dict(p.split('=', 1) for p in out.split('|')[1:] if '=' in p)
    print('game says it took out: Sticky Keys bits %s, Filter Keys bits %s' % (f.get('accsticky', '?'), f.get('accfilter', '?')))
    bad = []
    if not out.startswith('ok'):
        bad.append('the game did not end cleanly: %r' % out[:80])
    for name, b, d in (('Sticky Keys', before[0], during[0]), ('Filter Keys', before[1], during[1])):
        if b & ON:
            if d != b: bad.append('%s is turned on and was changed (%d -> %d)' % (name, b, d))
        elif d & SHORTCUT:
            bad.append('%s shortcut still active during the game' % name)
        if (d | SHORTCUT | ASKS) != (b | SHORTCUT | ASKS):
            bad.append('%s: other settings changed during the game (%d -> %d)' % (name, b, d))
    if after != before:
        bad.append('settings not put back: before %s, after %s' % (before, after))
    if not (before[0] & SHORTCUT or before[1] & SHORTCUT):
        print('NOTE: neither shortcut was active on this PC to begin with, so this run shows little')
    for b in bad:
        print('WRONG:', b)
    print('STICKY TEST FAILED' if bad else 'STICKY TEST PASSED')
    return 1 if bad else 0

if __name__ == '__main__':
    sys.exit(main())

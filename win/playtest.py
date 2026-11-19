"""Windows smoke test of live play with injected keys: fire taps, a turn, a walk, optional
key combinations, screenshots of the Excel window, quit with Q. The Windows counterpart of
playtest.py. Run from the repo root:

    python win\\playtest.py                 basic run
    python win\\playtest.py combos          also holds Shift and Ctrl together with game keys (run and fire while moving)
    python win\\playtest.py Host.PlayB      another entry macro (PlayA: never a full redraw, PlayB: always)

Keys are only ever sent while the workbook's own window is the foreground window. If the
game cannot be made to quit, the Excel instance this script started is closed.

Injected key-ups can be lost: if the workbook window goes away while a key is held and the
window that takes over does not accept injected input, the key stays down for the whole
machine and the next game quits the moment it starts (a held Q). So Q is only tapped, the
workbook is kept open for a few seconds afterwards, and every key is released again at the
end from a small window of our own, where the release is known to register."""
import ctypes, os, subprocess, sys, time
from ctypes import wintypes
from PIL import ImageGrab

u32 = ctypes.windll.user32
k32 = ctypes.windll.kernel32
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)      # window rectangles in real pixels
except Exception:
    pass
u32.GetForegroundWindow.restype = wintypes.HWND
u32.SetForegroundWindow.argtypes = [wintypes.HWND]
u32.IsWindowVisible.argtypes = [wintypes.HWND]
u32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
u32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
u32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
u32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
u32.GetAsyncKeyState.restype = ctypes.c_short
u32.CreateWindowExW.restype = wintypes.HWND
u32.CreateWindowExW.argtypes = [wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD, ctypes.c_int, ctypes.c_int,
                                ctypes.c_int, ctypes.c_int, wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID]
u32.DestroyWindow.argtypes = [wintypes.HWND]
u32.PeekMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT, wintypes.UINT]
u32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
u32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
ENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

VK = dict(ctrl=0x11, shift=0x10, alt=0x12, left=0x25, up=0x26, right=0x27, down=0x28,
          w=0x57, a=0x41, s=0x53, d=0x44, q=0x51, space=0x20, **{'1': 0x31})
EXTENDED = {0x25, 0x26, 0x27, 0x28}
OUT = 'out'
RESULT = os.path.join(OUT, 'play_result.txt')

def windows():
    """[(hwnd, class, title, pid)] for every visible top-level window."""
    found = []
    def cb(h, _):
        if u32.IsWindowVisible(h):
            c = ctypes.create_unicode_buffer(256); t = ctypes.create_unicode_buffer(512); pid = wintypes.DWORD()
            u32.GetClassNameW(h, c, 256); u32.GetWindowTextW(h, t, 512); u32.GetWindowThreadProcessId(h, ctypes.byref(pid))
            found.append((h, c.value, t.value, pid.value))
        return True
    u32.EnumWindows(ENUMPROC(cb), 0)
    return found

def excel_window(known):
    for h, c, t, pid in windows():
        if c == 'XLMAIN' and h not in known and 'doom' in t.lower():
            return h, pid
    return None, None

def dialogs(pid):
    return [t for h, c, t, p in windows() if p == pid and c != 'XLMAIN' and t]

def is_front(hwnd):
    return u32.GetForegroundWindow() == hwnd

def bring_front(hwnd):
    if is_front(hwnd):
        return True
    u32.SetForegroundWindow(hwnd)
    time.sleep(0.4)
    return is_front(hwnd)

def keys_down():
    return [n for n, v in VK.items() if u32.GetAsyncKeyState(v) & 0x8000]

def release_all():
    """Release every key this script uses, from a window of our own. Returns the keys still down."""
    h = u32.CreateWindowExW(0x88, 'STATIC', 'playtest key release', 0x90000000, 0, 0, 260, 40, None, None, None, None)
    msg = wintypes.MSG()
    def pump():
        while u32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
            u32.TranslateMessage(ctypes.byref(msg)); u32.DispatchMessageW(ctypes.byref(msg))
    try:
        u32.SetForegroundWindow(h); time.sleep(0.3); pump()
        if not is_front(h):
            print('  key release: could not bring our own window to the front', flush=True)
        stuck = keys_down()
        if stuck:
            print('  keys still down before release: %s' % ' '.join(stuck), flush=True)
        for v in VK.values():
            send(v, False)
        time.sleep(0.15); pump()
        return keys_down()
    finally:
        u32.DestroyWindow(h)

def send(vk, down):
    flags = (0 if down else 2) | (1 if vk in EXTENDED else 0)
    u32.keybd_event(vk, u32.MapVirtualKeyW(vk, 0), flags, 0)

def hold(hwnd, names, seconds):
    """Hold the named keys together. Returns False (and sends nothing) if Excel is not in front."""
    if not bring_front(hwnd):
        print('  SKIPPED %s: Excel is not the foreground window' % '+'.join(names), flush=True)
        return False
    vks = [VK[n] for n in names]
    try:
        for v in vks:
            send(v, True); time.sleep(0.03)
        time.sleep(seconds)
    finally:
        for v in reversed(vks):
            send(v, False)
    return True

def shot(hwnd, name):
    r = wintypes.RECT()
    u32.GetWindowRect(hwnd, ctypes.byref(r))
    box = (max(r.left, 0), max(r.top, 0), r.right, r.bottom)
    path = os.path.join(OUT, name + '.png')
    try:
        ImageGrab.grab(bbox=box, all_screens=True).save(path)
        print('  %s: window %dx%d, front=%s' % (path, box[2] - box[0], box[3] - box[1], is_front(hwnd)), flush=True)
    except Exception as e:
        print('  %s: screenshot failed: %s' % (path, e), flush=True)

def summary(out):
    print('result:', out[:500])
    f = dict(p.split('=', 1) for p in out.split('|')[1:] if '=' in p)
    try:
        fr = max(1, int(f['frames'])); n = max(1, int(f['ticks'])); run = float(f['run'])
        print('TRUTH ammo=%s angle=%s pos=%s | key downs=%s ups=%s taken by engine=%s | polls=%s, Excel not in front for %s'
              % (f['ammo'], f['angle'], f['pos'], f['downs'], f['ups'], f['pops'], f['polls'], f['notfront']))
        print('PER FRAME ms: engine %.0f | scan %.0f | fills %.0f | redraw %.0f | nothing to paint %s, small %s, of %d frames'
              % ((run - float(f['paint'])) / n * 1000, float(f['tscan']) / fr * 1000, float(f['tfill']) / fr * 1000,
                 float(f['tredraw']) / fr * 1000, f['nzero'], f['nsmall'], fr))
        nb = max(1, int(f['batched'])); nu = max(1, int(f['nunb']))
        print('PATHS: small-frame %s frames, %.0f rectangles and %.1f ms each | full-redraw %s frames, %.0f rectangles and %.0f ms each | settle redraws %s'
              % (f['nunb'], int(f['rectsunb']) / nu, float(f['tunb']) / nu * 1000, f['batched'], int(f['cbatch']) / nb,
                 float(f['tbatch']) / nb * 1000, f['nsettle']))
        print('avg fps %.1f over %.1f s | fill errors %s %s | front-window check %s | telemetry file %s'
              % (n / run, run, f['fillerr'], f['errtext'], f['front'], f.get('telefile', '')))
        tele = {}
        if os.path.exists(f.get('telefile', '')):
            tele = dict(l.rstrip('\n').split('=', 1) for l in open(f['telefile'], encoding='utf-8', errors='replace') if '=' in l)
        for kind in ('state', 'class'):                  # frames per second by what the player / the painter was doing
            if kind + '_names' in tele:
                names = tele[kind + '_names'].split(',')
                ns = [int(x) for x in tele[kind + '_n'].split(',') if x]; ms = [float(x) for x in tele[kind + '_ms'].split(',') if x]
                print(kind.upper() + ': ' + ' | '.join('%s %d frames, %.0f ms each (%.1f fps)' % (nm, k, t / k, 1000 * k / t)
                                                       for nm, k, t in zip(names, ns, ms) if k and t))
    except Exception as e:
        print('stats parse failed:', e)

def main():
    combos = 'combos' in sys.argv[1:]
    macro = ([a for a in sys.argv[1:] if a.startswith('Host.')] or ['Host.Play'])[0]
    os.makedirs(OUT, exist_ok=True)
    if os.path.exists(RESULT):
        os.remove(RESULT)
    release_all()                                        # a key left down by an earlier run would quit this one
    known = {h for h, c, t, p in windows() if c == 'XLMAIN'}
    ps = subprocess.Popen(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', os.path.join('win', 'verify.ps1'),
                           '-Macro', macro, '-Out', RESULT, '-Maximize', '-Linger', '3'])
    hwnd = pid = None
    t0 = time.time()
    while hwnd is None and time.time() - t0 < 90 and ps.poll() is None:
        time.sleep(0.5)
        hwnd, pid = excel_window(known)
    if hwnd is None:
        print('Excel window with the workbook never appeared'); ps.kill(); return 2
    print('workbook window after %.1fs, pid %d' % (time.time() - t0, pid), flush=True)
    quit_sent = False
    try:
        time.sleep(6)                                    # layout, data load, engine start, first frames
        print('  in front: %s' % bring_front(hwnd), flush=True)
        shot(hwnd, 'wpt0')
        time.sleep(3)
        shot(hwnd, 'wpt1')                               # nothing pressed between wpt0 and wpt1
        for _ in range(5):                               # five 80 ms taps of Ctrl (fire)
            hold(hwnd, ['ctrl'], 0.08); time.sleep(0.7)
        shot(hwnd, 'wpt2')
        hold(hwnd, ['left'], 1.0); time.sleep(0.3)       # turn
        shot(hwnd, 'wpt3')
        hold(hwnd, ['w'], 1.5); time.sleep(0.3)          # walk
        shot(hwnd, 'wpt4')
        if combos:
            # Alt is left out on purpose: it switches the ribbon to KeyTips mode whatever the workbook does
            for names in (['shift', 'w'], ['ctrl', 'w'], ['ctrl', 'shift', 'w'], ['ctrl', 'a'], ['ctrl', 'd'], ['ctrl', 's'],
                          ['ctrl', '1'], ['ctrl', 'space'], ['ctrl', 'left']):
                if ps.poll() is not None or not u32.IsWindowVisible(hwnd):
                    print('  workbook window is gone before %s' % '+'.join(names), flush=True); break
                hold(hwnd, names, 0.8); time.sleep(0.6)
                print('  after %s: workbook window visible=%s, other Excel windows=%s'
                      % ('+'.join(names), bool(u32.IsWindowVisible(hwnd)), dialogs(pid)), flush=True)
                shot(hwnd, 'wpt_' + '_'.join(names))
        for _ in range(5):                               # a tap can fall inside one long repaint: retry
            quit_sent = hold(hwnd, ['q'], 0.12) or quit_sent
            time.sleep(0.5)
            if os.path.exists(RESULT) or ps.poll() is not None:
                break
        ps.wait(timeout=60 if quit_sent else 5)
    except subprocess.TimeoutExpired:
        print('game did not end (Q sent: %s); other Excel windows: %s' % (quit_sent, dialogs(pid)), flush=True)
        shot(hwnd, 'wpt_stuck')
    finally:
        left = release_all()
        if left:
            print('  WARNING: keys still reported down after release: %s' % ' '.join(left), flush=True)
        if ps.poll() is None:                            # only ever the instance started above
            subprocess.run(['taskkill', '/PID', str(pid), '/F'], capture_output=True)
            try:
                ps.wait(timeout=20)
            except subprocess.TimeoutExpired:
                ps.kill()
    if os.path.exists(RESULT):
        summary(open(RESULT, encoding='utf-8-sig', errors='replace').read())
        return 0
    print('no result: the macro never returned')
    return 1

if __name__ == '__main__':
    sys.exit(main())

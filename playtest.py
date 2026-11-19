"""Automated smoke test of live play with injected input: quick key taps, mouse turn,
clicks, a double-click, arrow turn; screenshots; quit with Q."""
import subprocess, threading, time, sys
import Quartz
import xl

def key(code, hold):
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, Quartz.CGEventCreateKeyboardEvent(None, code, True)); time.sleep(hold)
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, Quartz.CGEventCreateKeyboardEvent(None, code, False))
def click(pos, hold=0.12):
    for kind in (Quartz.kCGEventLeftMouseDown, Quartz.kCGEventLeftMouseUp):
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, Quartz.CGEventCreateMouseEvent(None, kind, pos, Quartz.kCGMouseButtonLeft)); time.sleep(hold)
def front():
    return xl.osa('tell application "System Events" to get name of first process whose frontmost is true', timeout=20)
def shot(name):
    print('  %s: front app = %s' % (name, front()), flush=True)
    subprocess.run(['screencapture', '-x', 'out/%s.png' % name]); subprocess.run(['sips', '-Z', '1400', 'out/%s.png' % name], capture_output=True)

W = Quartz.CGDisplayPixelsWide(Quartz.CGMainDisplayID()); H = Quartz.CGDisplayPixelsHigh(Quartz.CGMainDisplayID())
C = (W / 2, H / 2)
res = {}
print('open %.1fs' % xl.open_wb('out/doom.xlsm', wait=300), flush=True)
subprocess.run(['open', '-a', xl.APP]); time.sleep(1.5)
shot('pt0')
th = threading.Thread(target=lambda: res.update(out=xl.run('doom.xlsm', 'Host.Play', timeout=600)))
th.start(); time.sleep(float(sys.argv[1]) if len(sys.argv) > 1 else 10)
shot('pt1')
for _ in range(5):                      # five 80 ms taps of Ctrl (fire)
    key(59, 0.08); time.sleep(0.7)
time.sleep(0.5); shot('pt2')
for _ in range(5):                      # five 50 ms clicks
    click(C, 0.05); time.sleep(0.65)
time.sleep(0.5); shot('pt3')
for _ in range(3):                      # mouse: three nudges of 100 px to the right
    Quartz.CGWarpMouseCursorPosition((C[0] + 100, C[1])); time.sleep(0.35)
key(123, 1.0); time.sleep(0.3)          # hold left arrow 1 s
key(13, 1.5); time.sleep(0.3)           # hold W 1.5 s
shot('pt4')
shot('pt_live')
key(12, 0.3)                            # Q
th.join(timeout=60)
out = res.get('out', ('still running', 0))[0]
print('result:', out[:700])
f = dict(p.split('=') for p in out.split('|')[1:] if '=' in p)
try:
    print('TRUTH ammo=%s angle=%s pos=%s | fire keys down=%s clicks=%s | batched: %s frames, %.0f ms each, %.0f fills each' % (f['ammo'], f['angle'], f['pos'], f['downs'], f['clk'], f['batched'], float(f['tbatch']) / max(1, int(f['batched'])) * 1000, int(f['cbatch']) / max(1, int(f['batched']))))
    fr = int(f['frames']); eng = float(f['run']) - float(f['paint'])
    print('PER FRAME ms: engine %.0f | scan %.0f | fills %.0f (%.1f us each) | redraw %.0f | frames with nothing to paint %s, with <200 fills %s of %d'
          % (eng / int(f['ticks']) * 1000, float(f['tscan']) / fr * 1000, float(f['tfill']) / fr * 1000, float(f['tfill']) / max(1, int(f['calls'])) * 1e6, float(f['tredraw']) / fr * 1000, f['nzero'], f['nsmall'], fr))
    nb = max(1, int(f['batched'])); nu = max(1, int(f['nunb']))
    print('FILLS: %.1f rectangles per Excel call | small-frame path: %s frames, %.0f rectangles each, %.1f ms each | full-redraw path: %s frames, %.0f rectangles each, %.0f ms each | settle redraws %s'
          % (int(f['calls']) / max(1, int(f['omcalls'])), f['nunb'], int(f['rectsunb']) / nu, float(f['tunb']) / nu * 1000, f['batched'], int(f['cbatch']) / nb, float(f['tbatch']) / nb * 1000, f['nsettle']))
    n = int(f['ticks']); print('avg fps %.1f | avg paint %.0f ms | avg fills/frame %.0f | batched frames %d of %d | avg tolerance %.1f'
        % (n / float(f['run']), float(f['paint']) / int(f['frames']) * 1000, int(f['calls']) / int(f['frames']), int(f['batched']), int(f['frames']), int(f['tolsum']) / int(f['frames'])))
except Exception as e:
    print('stats parse failed', e)
time.sleep(0.5)
print('A1 value: %r' % xl.osa('tell application "%s" to get value of cell 1 of row 1 of active sheet' % xl.APP, timeout=20)[:60],
      '| active cell: %r' % xl.osa('tell application "%s" to get address of active cell' % xl.APP, timeout=20)[:40])
print('dialog:', xl.dialog()[:200])

"""Mouse-turn test: nudge the mouse steadily for 3 s (no keys), count distinct pictures on
screen and report the cost of the small-frame path versus the full-redraw path."""
import subprocess, threading, time, hashlib, sys
import numpy as np
import Quartz
from PIL import Image
import xl
macro = sys.argv[1] if len(sys.argv) > 1 else 'Play'
W = Quartz.CGDisplayPixelsWide(Quartz.CGMainDisplayID()); H = Quartz.CGDisplayPixelsHigh(Quartz.CGMainDisplayID())
xl.open_wb('out/doom.xlsm', wait=300)
subprocess.run(['open', '-a', xl.APP]); time.sleep(1.5)
res = {}
th = threading.Thread(target=lambda: res.update(out=xl.run('doom.xlsm', 'Host.' + macro, timeout=300)))
th.start(); time.sleep(9)
stop = []
def nudge():
    while not stop:
        Quartz.CGWarpMouseCursorPosition((W / 2 + 14, H / 2)); time.sleep(0.05)
nt = threading.Thread(target=nudge); nt.start()
hashes = []
for i in range(12):
    subprocess.run(['screencapture', '-x', 'out/fz.png'])
    a = np.asarray(Image.open('out/fz.png').convert('L').resize((480, 300)))
    hashes.append(hashlib.md5(a[40:200, 100:360].tobytes()).hexdigest())
stop.append(1); nt.join(); time.sleep(0.5)
for d in (True, False):
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, Quartz.CGEventCreateKeyboardEvent(None, 12, d)); time.sleep(0.3)
th.join(timeout=60)
out = res.get('out', ('none',))[0]
f = dict(p.split('=', 1) for p in out.split('|log=')[0].split('|')[1:] if '=' in p)
nb = max(1, int(f['batched'])); nu = max(1, int(f['nunb']))
log = [tuple(float(v) for v in e.split(':')) for e in out.split('|log=')[1].split(';') if e]
turning = [log[i][0] for i in range(1, len(log)) if log[i][1] != log[i - 1][1]]
fps_turn = (len(turning) - 1) / (turning[-1] - turning[0]) if len(turning) > 2 else 0
print('%s: distinct pictures %d of 12 | fps while turning %.1f | small-frame path: %s frames, %.0f rects, %.1f ms fills each | full-redraw path: %s frames, %.0f rects, %.0f ms each'
      % (macro, len(set(hashes)), fps_turn, f['nunb'], int(f['rectsunb']) / nu, float(f['tunb']) / nu * 1000, f['batched'], int(f['cbatch']) / nb, float(f['tbatch']) / nb * 1000))

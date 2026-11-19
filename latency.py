"""Measure display latency: hold the left-arrow key for 1 s and compare when the engine's
view angle changes (logged per frame inside Excel) with when the picture on screen changes."""
import subprocess, threading, time, datetime, sys
import numpy as np
import Quartz
from PIL import Image
import xl
def tod():
    n = datetime.datetime.now(); return n.hour * 3600 + n.minute * 60 + n.second + n.microsecond / 1e6
def keyev(code, down):
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, Quartz.CGEventCreateKeyboardEvent(None, code, down))
print('open %.1fs' % xl.open_wb('out/doom.xlsm', wait=300), flush=True)
subprocess.run(['open', '-a', xl.APP]); time.sleep(1.5)
res = {}
th = threading.Thread(target=lambda: res.update(out=xl.run('doom.xlsm', 'Host.' + (sys.argv[1] if len(sys.argv) > 1 else 'Play'), timeout=300)))
th.start(); time.sleep(10)
shots = []
def grab():
    t = tod()
    subprocess.run(['screencapture', '-x', 'out/lat.png'])
    a = np.asarray(Image.open('out/lat.png').convert('L').resize((480, 310))).astype(int)
    shots.append((t, a[60:240, 60:230]))
t_end = time.time() + 6.5
pressed = released = None
while time.time() < t_end:
    grab()
    if pressed is None and len(shots) >= 6:
        pressed = tod(); keyev(123, True)
    if pressed is not None and released is None and tod() - pressed >= 1.0:
        released = tod(); keyev(123, False)
keyev(12, True); time.sleep(0.3); keyev(12, False)
th.join(timeout=60)
out = res.get('out', ('none',))[0]
log = [tuple(float(v) for v in e.split(':')) for e in out.split('|log=')[1].split(';') if e] if '|log=' in out else []
print('frames logged:', len(log), '| screenshots:', len(shots), '(every %.2f s)' % ((shots[-1][0] - shots[0][0]) / (len(shots) - 1)))
eng = [(t, a) for t, a, m in log]
ch = [eng[i][0] for i in range(1, len(eng)) if eng[i][1] != eng[i - 1][1]]
scr = [shots[i][0] for i in range(1, len(shots)) if np.abs(shots[i][1] - shots[i - 1][1]).mean() > 1.5]
if ch and scr:
    print('key pressed at 0.00, released at %.2f' % (released - pressed))
    print('engine view started turning at %+.2f s, stopped at %+.2f s' % (ch[0] - pressed, ch[-1] - pressed))
    print('screen started changing by   %+.2f s, last change at %+.2f s' % (scr[0] - pressed, scr[-1] - pressed))
    fr = [t for t, a in eng if pressed <= t <= released + 0.3]
    if len(fr) > 2: print('frame rate while turning: %.1f fps' % ((len(fr) - 1) / (fr[-1] - fr[0])))
    f = dict(p.split('=', 1) for p in out.split('|')[1:] if '=' in p)
    print('whole run: %.1f fps, %s of %s frames used the full redraw, %.0f fills per frame' % (int(f['ticks']) / float(f['run']), f['batched'], f['frames'], int(f['calls']) / int(f['frames'])))
else:
    print('no change detected: engine changes %d, screen changes %d' % (len(ch), len(scr)), out[:200])

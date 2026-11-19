"""Does the screen keep up during small-change frames? Hold fire for 3 s and count how many
distinct pictures appear in rapid screenshots, for several ways of ending a quiet frame."""
import subprocess, threading, time, hashlib
import numpy as np
import Quartz
from PIL import Image
import xl
names = {'Play': 'hybrid', 'PlayB': 'always full redraw'}
def keyev(code, down):
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, Quartz.CGEventCreateKeyboardEvent(None, code, down))
print('open %.1fs' % xl.open_wb('out/doom.xlsm', wait=300), flush=True)
subprocess.run(['open', '-a', xl.APP]); time.sleep(1.5)
import sys
KEYS = [int(k) for k in (sys.argv[1] if len(sys.argv) > 1 else '3').split(',')]
FIRE = KEYS[0]
for macro in (sys.argv[2] if len(sys.argv) > 2 else 'Play',):
    res = {}
    th = threading.Thread(target=lambda: res.update(out=xl.run('doom.xlsm', 'Host.' + macro, timeout=300)))
    th.start(); time.sleep(9)
    for kc in KEYS: keyev(kc, True)
    hashes = []
    for i in range(14):
        subprocess.run(['screencapture', '-x', 'out/fz.png'])
        a = np.asarray(Image.open('out/fz.png').convert('L').resize((480, 300)))
        hashes.append(hashlib.md5(a[40:200, 100:360].tobytes()).hexdigest())
        time.sleep(0.15)
    for kc in KEYS: keyev(kc, False)
    time.sleep(0.3)
    keyev(12, True); time.sleep(0.3); keyev(12, False)
    th.join(timeout=60)
    out = res.get('out', ('none',))[0]
    f = dict(p.split('=', 1) for p in out.split('|')[1:] if '=' in p)
    if 'ticks' not in f or float(f.get('run', 0)) == 0:
        print(macro, 'FAILED:', out[:200]); continue
    print('held keys %-8s ' % sys.argv[1] + '%-18s distinct pictures in 14 screenshots: %2d | shots %d, moved to %s | batched frames %s of %s | fps %.1f'
          % (names[macro], len(set(hashes)), 50 - int(f.get('ammo', 50)), f.get('pos'), f.get('batched'), f.get('frames'), int(f['ticks']) / float(f['run'])))
    time.sleep(2.0)

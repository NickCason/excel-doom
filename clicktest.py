"""Does holding the mouse button stall the game? Hold a click for 0.6 s at a given point
and read the longest frame and shots fired from the telemetry."""
import subprocess, threading, time, sys, glob, os
import Quartz
import xl
W = Quartz.CGDisplayPixelsWide(Quartz.CGMainDisplayID()); H = Quartz.CGDisplayPixelsHigh(Quartz.CGMainDisplayID())
where = sys.argv[1] if len(sys.argv) > 1 else 'centre'
xl.open_wb('out/doom.xlsm', wait=300)
subprocess.run(['open', '-a', xl.APP]); time.sleep(1.5)
res = {}
th = threading.Thread(target=lambda: res.update(out=xl.run('doom.xlsm', 'Host.Play', timeout=300)))
th.start(); time.sleep(9)
P = (W / 2, H / 2) if where == 'centre' else (int(W * 0.54), 4)
for hold in (0.6, 0.6, 0.6):
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, Quartz.CGEventCreateMouseEvent(None, Quartz.kCGEventLeftMouseDown, P, Quartz.kCGMouseButtonLeft)); time.sleep(hold)
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, Quartz.CGEventCreateMouseEvent(None, Quartz.kCGEventLeftMouseUp, P, Quartz.kCGMouseButtonLeft)); time.sleep(1.0)
for d in (True, False):
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, Quartz.CGEventCreateKeyboardEvent(None, 12, d)); time.sleep(0.3)
th.join(timeout=60)
f = max(glob.glob(os.path.expanduser('~/Library/Containers/com.microsoft.Excel/Data/doom_telemetry_*')), key=os.path.getmtime)
kv = dict(l.rstrip('\n').split('=', 1) for l in open(f) if '=' in l and not l.startswith(('tl=', 'slow=')))
slow = [[float(v) for v in l[5:].rstrip(',\n').split(',')] for l in open(f) if l.startswith('slow=')]
print('three 0.6 s clicks: longest frame %s ms | frames over 200 ms: %d (%s) | bullets fired %s | clicks seen %s'
      % (kv['max_frame_ms'], len(slow), ', '.join('%.0f' % s[2] for s in slow), kv['bullets_fired'], kv['mouse_clicks']))

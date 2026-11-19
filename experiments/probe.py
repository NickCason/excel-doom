import subprocess, threading, time
import Quartz
import xl
def se(script):
    return xl.osa('tell application "System Events"\n%s\nend tell' % script, timeout=20)
try: xl.close_wb('doom.xlsm')
except Exception: pass
time.sleep(1)
print('open %.1fs' % xl.open_wb('out/doom.xlsm', wait=300), flush=True)
subprocess.run(['open', '-a', xl.APP]); time.sleep(1.2)
res = {}
th = threading.Thread(target=lambda: res.update(out=xl.run('doom.xlsm', 'Host.Probe', timeout=120)))
th.start(); time.sleep(1.2)
print('left arrow hold:', se('key down (ASCII character 28)\ndelay 0.5\nkey up (ASCII character 28)') or 'ok')
print('right arrow hold:', se('key down (ASCII character 29)\ndelay 0.5\nkey up (ASCII character 29)') or 'ok')
print('j x3:', se('keystroke "j"\ndelay 0.2\nkeystroke "j"\ndelay 0.2\nkeystroke "j"') or 'ok')
Quartz.CGWarpMouseCursorPosition((400, 700)); time.sleep(0.3)
Quartz.CGWarpMouseCursorPosition((1300, 700)); time.sleep(0.3)
for kind in (Quartz.kCGEventLeftMouseDown, Quartz.kCGEventLeftMouseUp):
    ev = Quartz.CGEventCreateMouseEvent(None, kind, (900, 700), Quartz.kCGMouseButtonLeft)
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, ev); time.sleep(0.4)
th.join(timeout=60)
print('probe:', res.get('out', ('no result',))[0][:500])
print('dialog:', xl.dialog()[:150])

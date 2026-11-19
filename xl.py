"""Drive the test copy of Excel (/Applications/Microsoft Excel.app): open a generated
workbook, call a macro, close it. Watches for modal VBA dialogs (compile errors etc.),
reports their text, dismisses them and restarts the test copy so nothing is left stuck."""
import subprocess, sys, time, os

APP = '/Applications/Microsoft Excel.app'
BIN = APP + '/Contents/MacOS/Microsoft Excel'

def pid():
    r = subprocess.run(['pgrep', '-f', BIN], capture_output=True, text=True).stdout.split()
    return int(r[0]) if r else None

def osa(script, timeout=3600):
    r = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=timeout)
    return (r.stdout.strip() + (' !ERR ' + r.stderr.strip() if r.returncode else ''))

DIALOG = '''tell application "System Events"
  tell (first process whose unix id is %d)
    repeat with w in windows
      if (subrole of w) is "AXDialog" then
        set t to ""
        try
          set t to (value of static texts of w) as text
        end try
        if %s then
          try
            click (first button of w whose name is "OK")
          end try
        end if
        return t
      end if
    end repeat
    return ""
  end tell
end tell'''

def dialog(dismiss=False):
    p = pid()
    if not p:
        return ''
    try:
        return osa(DIALOG % (p, 'true' if dismiss else 'false'), timeout=20).replace('\n', ' ')
    except subprocess.TimeoutExpired:
        return ''

def restart():
    p = pid()
    if p:
        subprocess.run(['kill', '-TERM', str(p)])
        for _ in range(20):
            if not pid():
                break
            time.sleep(0.5)
    subprocess.run(['open', '-g', '-a', APP])
    for _ in range(60):
        time.sleep(1)
        try:
            if pid() and '!ERR' not in workbooks():
                break
        except subprocess.TimeoutExpired:
            pass
    time.sleep(2)

def workbooks():
    return osa('tell application "%s" to get name of every workbook' % APP, timeout=30)

def open_wb(path, wait=120):
    path = os.path.abspath(path)
    name = os.path.basename(path)
    try:
        if name in workbooks():
            close_wb(name); time.sleep(1.0)
    except subprocess.TimeoutExpired:
        pass
    t0 = time.time()
    subprocess.run(['open', '-g', '-a', APP, path], check=True)
    tries = 0
    while time.time() - t0 < wait:
        try:
            if name in workbooks():
                return time.time() - t0
        except subprocess.TimeoutExpired:
            pass
        time.sleep(0.5)
        tries += 1
        if tries % 20 == 0:
            subprocess.run(['open', '-g', '-a', APP, path])
    raise RuntimeError('workbook did not open; dialog=%r' % dialog())

def run(name, macro, timeout=3600):
    t0 = time.time()
    script = ('with timeout of %d seconds\ntell application "%s" to run VB macro "%s!%s"\nend timeout'
              % (timeout, APP, name, macro))
    p = subprocess.Popen(['osascript', '-e', script], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    nextcheck = 2.0
    while True:
        try:
            out, err = p.communicate(timeout=nextcheck)
            return out.strip() + (' !ERR ' + err.strip() if p.returncode else ''), time.time() - t0
        except subprocess.TimeoutExpired:
            pass
        d = dialog(dismiss=False)
        if d:
            subprocess.run(['screencapture', '-x', 'out/dialog.png'])
            dialog(dismiss=True)
            p.kill(); p.communicate()
            restart()
            return 'DIALOG: ' + d, time.time() - t0
        if time.time() - t0 > timeout:
            p.kill(); p.communicate()
            restart()
            return 'TIMEOUT', time.time() - t0
        nextcheck = min(nextcheck * 1.5, 15)

def close_wb(name):
    try:
        return osa('tell application "%s" to close workbook "%s" saving no' % (APP, name), timeout=30)
    except subprocess.TimeoutExpired:
        return 'close timed out'

def session(path, macro, timeout=3600):
    name = os.path.basename(path)
    open_wb(path)
    out, dt = run(name, macro, timeout)
    if not out.startswith(('DIALOG', 'TIMEOUT')):
        close_wb(name)
    return out, dt

if __name__ == '__main__':
    out, dt = session(sys.argv[1], sys.argv[2])
    print('%.2fs -> %s' % (dt, out))

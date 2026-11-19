"""Compare the three picture dumps written by  win\\verify.ps1 -Dump out\\frame_  (Host.Dump):
0 the engine's last frame, 1 what the painter believes each cell shows, 2 what the cells hold.
Prints how far apart they are and writes out/picture_*.png to look at.
Usage: python tests/picture_check.py [prefix, default out/frame_]"""
import sys
from PIL import Image

W, H = 320, 200

def load(path):
    t = open(path, encoding='ascii', errors='replace').read().strip()
    px = []
    for i in range(W * H):
        v = int(t[i * 6:i * 6 + 6], 16)
        px.append((v & 255, (v >> 8) & 255, v >> 16))      # VBA colour: red is the low byte
    return px

def diff(a, b):
    return [max(abs(p[0] - q[0]), abs(p[1] - q[1]), abs(p[2] - q[2])) for p, q in zip(a, b)]

def report(name, d):
    n = len(d); bad = [v for v in d if v]
    hist = [sum(1 for v in d if lo < v <= hi) for lo, hi in ((0, 8), (8, 16), (16, 24), (24, 48), (48, 255))]
    print('%-34s differ %5.1f %% | worst %3d | mean of those %5.1f | by size 1-8: %d, 9-16: %d, 17-24: %d, 25-48: %d, over 48: %d'
          % (name, 100.0 * len(bad) / n, max(d), sum(bad) / max(1, len(bad)), *hist))

def save(path, px, scale=3):
    im = Image.new('RGB', (W, H)); im.putdata(px)
    im.resize((W * scale, H * scale), Image.NEAREST).save(path)

def runs(px):
    """Mean length of same-colour runs along rows: smearing shows up as longer runs."""
    n = 0
    for y in range(H):
        row = px[y * W:(y + 1) * W]
        n += 1 + sum(1 for x in range(1, W) if row[x] != row[x - 1])
    return W * H / n

if __name__ == '__main__':
    prefix = sys.argv[1] if len(sys.argv) > 1 else 'out/frame_'
    truth, shown, cells = (load('%s%d.txt' % (prefix, k)) for k in range(3))
    report('engine frame vs painter record', diff(truth, shown))
    report('painter record vs real cells', diff(shown, cells))
    report('engine frame vs real cells', diff(truth, cells))
    print('mean same-colour run along a row: engine %.2f px, cells %.2f px' % (runs(truth), runs(cells)))
    save('out/picture_truth.png', truth); save('out/picture_cells.png', cells)
    d = diff(truth, cells)
    save('out/picture_diff.png', [(min(255, v * 6),) * 3 for v in d])
    print('wrote out/picture_truth.png, out/picture_cells.png, out/picture_diff.png (difference x6)')

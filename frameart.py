"""Compose the tutorial frame around the game picture from the game's own graphics
(border flat, bevel, logo, HUD font). Output: a grid of palette indices + the palette."""
import struct

FW, FH, GX, GY = 592, 224, 136, 12

class Wad:
    def __init__(self, path='vendor/wad/doom1.wad'):
        self.b = open(path, 'rb').read()
        n, off = struct.unpack_from('<ii', self.b, 4)
        self.dir = {}
        for i in range(n):
            p, sz, name = struct.unpack_from('<ii8s', self.b, off + i * 16)
            self.dir.setdefault(name.rstrip(b'\0').decode('ascii').upper(), (p, sz))
    def lump(self, name):
        p, sz = self.dir[name.upper()]
        return self.b[p:p + sz]
    def has(self, name):
        return name.upper() in self.dir
    def palette(self):
        d = self.lump('PLAYPAL')
        return [tuple(d[i * 3:i * 3 + 3]) for i in range(256)]
    def patch(self, name):
        """-> (w, h, leftofs, topofs, {(x, y): index})"""
        d = self.lump(name)
        w, h, lo, to = struct.unpack_from('<hhhh', d, 0)
        px = {}
        for x in range(w):
            p = struct.unpack_from('<i', d, 8 + 4 * x)[0]
            while d[p] != 255:
                top, n = d[p], d[p + 1]
                for k in range(n):
                    px[(x, top + k)] = d[p + 3 + k]
                p += n + 4
        return w, h, lo, to, px

class Canvas:
    def __init__(self, wad):
        self.wad = wad; self.g = [[0] * FW for _ in range(FH)]
        cm = wad.lump('COLORMAP'); self.cmap = [cm[i * 256:(i + 1) * 256] for i in range(34)]
        self.pal = wad.palette()
    def put(self, x, y, idx):
        if 0 <= x < FW and 0 <= y < FH:
            self.g[y][x] = idx
    def draw(self, name, x, y, remap=None, origin=False):
        w, h, lo, to, px = self.wad.patch(name)
        if origin: x -= lo; y -= to
        for (dx, dy), v in px.items():
            self.put(x + dx, y + dy, remap[v] if remap else v)
        return w
    def tile(self, flat, level):
        d = self.wad.lump(flat); m = self.cmap[level]
        for y in range(FH):
            for x in range(FW):
                self.g[y][x] = m[d[(y & 63) * 64 + (x & 63)]]
    def shade(self, x0, y0, x1, y1, level):
        """Darken a rectangle already drawn (panel backing)."""
        m = self.cmap[level]
        for y in range(y0, y1):
            for x in range(x0, x1):
                self.g[y][x] = m[self.g[y][x]]
    def text_width(self, s):
        w = 0
        for ch in s.upper():
            nm = 'STCFN%03d' % ord(ch)
            w += self.wad.patch(nm)[0] if (ch != ' ' and self.wad.has(nm)) else 4
        return w
    def text(self, s, x, y, remap=None):
        for ch in s.upper():
            nm = 'STCFN%03d' % ord(ch)
            if ch == ' ' or not self.wad.has(nm):
                x += 4; continue
            x += self.draw(nm, x, y, remap)
        return x
    def grey_map(self):
        """Remap the red HUD font to light grey using the palette's grey ramp."""
        greys = range(80, 112)
        def lum(c): return 0.3 * c[0] + 0.59 * c[1] + 0.11 * c[2]
        out = list(range(256))
        for i in range(256):
            target = min(255.0, lum(self.pal[i]) * 2.6 + 40)
            out[i] = min(greys, key=lambda g: abs(lum(self.pal[g]) - target))
        return out

def compose():
    wad = Wad(); c = Canvas(wad)
    c.tile('FLOOR7_2', 10)
    # darker backing for the two side panels
    c.shade(4, 4, GX - 12, FH - 4, 14)
    c.shade(GX + 320 + 12, 4, FW - 4, FH - 4, 14)
    # the game window: title picture and the engine's own bevel
    c.draw('TITLEPIC', GX, GY)
    for x in range(0, 320, 8):
        c.draw('BRDR_T', GX + x, GY - 8); c.draw('BRDR_B', GX + x, GY + 200)
    for y in range(0, 200, 8):
        c.draw('BRDR_L', GX - 8, GY + y); c.draw('BRDR_R', GX + 320, GY + y)
    c.draw('BRDR_TL', GX - 8, GY - 8); c.draw('BRDR_TR', GX + 320, GY - 8)
    c.draw('BRDR_BL', GX - 8, GY + 200); c.draw('BRDR_BR', GX + 320, GY + 200)
    grey = c.grey_map()
    def entries(x0, y, items):
        for keys, action in items:
            c.text(keys, x0, y, grey)
            c.text(action, x0 + 8, y + 9)
            y += 21
        return y
    lw = GX - 16
    c.draw('M_DOOM', 4 + (lw - 123) // 2, 8)
    y = entries(10, 62, [('W   UP', 'FORWARD'), ('S   DOWN', 'BACK'), ('LEFT  RIGHT', 'TURN'), ('MOUSE', 'TURN'),
                         ('A   D', 'STRAFE'), ('SHIFT', 'RUN'), ('TAB', 'MAP')])
    rx = GX + 320 + 18
    y = entries(rx, 10, [('CTRL   CLICK', 'FIRE'), ('SPACE   E', 'OPEN DOORS'), ('1 - 7', 'WEAPONS'), ('ESC', 'MENU'),
                         ('ENTER', 'SELECT'), ('M', 'MOUSE ON OFF'), ('Q', 'QUIT')])
    c.text('TO START', rx, y + 6, grey)
    c.text('CLICK HERE', rx + 8, y + 15)
    widths = {s: c.text_width(s) for s in ('CTRL   CLICK', 'LEFT  RIGHT', 'MOUSE ON OFF', 'OPEN DOORS', 'CLICK HERE')}
    return c.g, c.pal, widths, y

def panels(scale=4):
    """-> (title grid 200x320 of palette indices, palette, {name: png bytes}) for the four frame pieces."""
    import io
    from PIL import Image
    g, pal, _, _ = compose()
    im = Image.new('P', (FW, FH)); im.putpalette(bytes(v for c in pal for v in c))
    im.putdata([v for row in g for v in row])
    rgb = im.convert('RGB')
    boxes = {'left': (0, 0, GX, FH), 'right': (GX + 320, 0, FW, FH), 'top': (GX, 0, GX + 320, GY), 'bottom': (GX, GY + 200, GX + 320, FH)}
    out = {}
    for name, box in boxes.items():
        p = rgb.crop(box)
        p = p.resize((p.width * scale, p.height * scale), Image.NEAREST)
        b = io.BytesIO(); p.save(b, 'PNG', optimize=True); out[name] = b.getvalue()
    title = [row[GX:GX + 320] for row in g[GY:GY + 200]]
    return title, pal, out

if __name__ == '__main__':
    from PIL import Image
    g, pal, widths, y = compose()
    im = Image.new('P', (FW, FH)); im.putpalette(bytes(v for c in pal for v in c))
    im.putdata([v for row in g for v in row])
    im.convert('RGB').resize((FW * 2, int(FH * 2 * 8 / 7)), Image.NEAREST).save('out/frame_preview.png')
    print('text widths:', widths, '| panel inner width', GX - 16 - 12, '| right column ends at y', y + 33, 'of', FH)
    print('distinct colours:', len(set(v for row in g for v in row)))

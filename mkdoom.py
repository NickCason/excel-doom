"""Build doom.xlsm: translated engine + runtime + host as VBA, shareware WAD as cell values."""
import os, sys, struct, time, re, pickle
import mkxlsm, w2v, rtvba, hostvba, vbalint, frameart, televba

WADC = 701
WADR = 300      # first sheet row (0-based) holding WAD data: below the game, so game rows stay light
OM = set('''application worksheet range variant thisworkbook worksheets cells interior color screenupdating statusbar
enablecancelkey xldisabled xlinterrupt onkey windows zoom displaygridlines goto select columnwidth rowheight value doevents timer
format chr instr array lbound ubound cstr cint abs err description number declare ptrsafe lib const mac
cgeventsourcekeystate getasynckeystate nothing step boolean me psn byte getfrontprocess getcurrentprocess sameprocess
getforegroundwindow hwnd activecell row column top playing evtcount pointapi getcursorpos setcursorpos getsystemmetrics
x y double iif ev cgeventsourcecounterforeventtype rgb smallscroll len clear visiblerange rows columns count window each in object shapes visible entirerow hidden address mid string entirecolumn windowstate xlnormal xlmaximized skipfills displayformulabar mach_absolute_time queryperformancecounter queryperformancefrequency
vblf vbcr replace right erase now telepath freefile open output print close environ version operatingsystem int bad nohi cgeventsourcebuttonstate cgeventcreate cgeventgetlocation cfrelease cgwarpmousecursorposition
cgassociatemouseandmousecursorposition cgmaindisplayid cgdisplaypixelswide cgdisplaypixelshigh usablewidth usableheight width height displayheadings left msgbox longptr type hi lo byref'''.split())

SHEET_CODE = '''Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Playing Then
        Cancel = True
        EvtCount = EvtCount + 1
    End If
End Sub
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    If Playing Then
        Cancel = True
        EvtCount = EvtCount + 1
    End If
End Sub
'''

def wad_sheet(path='vendor/wad/doom1.wad'):
    """Sheet XML. Column A and the column after the game are wide, row 1 and the row after the
    game are tall: the tutorial frame is four pictures sitting on those few big cells. The game
    is 320x200 small cells starting at B2, pre-coloured with the title picture. WAD far right."""
    title, pal, pngs = frameart.panels()
    GXu, GYu = frameart.GX, frameart.GY
    used = sorted(set(v for row in title for v in row))
    sid = {v: i + 1 for i, v in enumerate(used)}
    styles = mkxlsm.styles_xml([pal[v] for v in used] + [(0, 0, 0)])
    black = len(used) + 1
    gcols = [mkxlsm.col_name(i + 2) for i in range(320)]
    wad = open(path, 'rb').read()
    pad = wad + b'\0' * (-len(wad) % 64)
    words = struct.unpack('<%di' % (len(pad) // 4), pad)
    nrows = len(words) // 16
    cols = [mkxlsm.col_name(WADC + i) for i in range(16)]
    meta = '<c r="%s%d"><v>%d</v></c><c r="%s%d"><v>%d</v></c>' % (mkxlsm.col_name(WADC - 2), WADR + 1, len(wad), mkxlsm.col_name(WADC - 1), WADR + 1, nrows)
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr codeName="Sheet1"/>'
           '<sheetViews><sheetView showGridLines="0" showRowColHeaders="0" zoomScale="100" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>'
           '<sheetFormatPr defaultRowHeight="3" customHeight="1"/>'
           '<cols><col min="1" max="1" width="%.6f" customWidth="1" style="%d"/><col min="2" max="321" width="0.571429" customWidth="1"/>'
           '<col min="322" max="322" width="%.6f" customWidth="1" style="%d"/><col min="323" max="%d" width="40" customWidth="1" style="%d"/></cols><sheetData>'
           % (GXu * 4 / 7.99, black, GXu * 4 / 7.99, black, WADC - 10, black)]
    big = {1: GYu * 3, 202: GYu * 3, 203: 409, 204: 409}
    for r in range(WADR + nrows):
        attrs = ' ht="%d" customHeight="1" s="%d" customFormat="1"' % (big[r + 1], black) if (r + 1) in big else (' ht="3" customHeight="1"' if 1 <= r <= 200 else '')
        fr = ''
        if 1 <= r <= 200:
            g = title[r - 1]
            fr = ''.join('<c r="%s%d" s="%d"/>' % (gcols[x], r + 1, sid[g[x]]) for x in range(320))
        wr = r - WADR
        row = words[wr * 16:wr * 16 + 16] if 0 <= wr < nrows else ()
        if not row and r > 204:
            continue
        cells = ''.join('<c r="%s%d"><v>%d</v></c>' % (cols[i], r + 1, v) for i, v in enumerate(row) if v)
        out.append('<row r="%d"%s>%s%s%s</row>' % (r + 1, attrs, fr, meta if r == WADR else '', cells))
    out.append('</sheetData><drawing r:id="rId1"/></worksheet>')
    pictures = [('Left', pngs['left'], 0, 0, 1, 202), ('Right', pngs['right'], 321, 0, 322, 202, 'PlayDoom'),
                ('Top', pngs['top'], 1, 0, 321, 1), ('Bottom', pngs['bottom'], 1, 201, 321, 202)]
    return ''.join(out), styles, pictures

def build(out='out/doom.xlsm', wasm_path='work/doomxl.wasm'):
    os.makedirs(os.path.dirname(out) or '.', exist_ok=True)
    t = time.time()
    T = w2v.Translator(wasm_path)
    mods, stats = T.build()
    build_id = time.strftime('%Y-%m-%d %H:%M:%S')
    mods += [('Rt', rtvba.RT), ('RtF', rtvba.RTF), ('Wasi', rtvba.WASI), ('Tele', televba.TELE), ('Host', hostvba.HOST.replace('@@BUILD@@', build_id))]
    host_src = re.sub(r'(?m)^#.*$', '', hostvba.HOST)
    host_src = re.sub(r'(?m)^Private Declare .*$', '', host_src)
    host_src = re.sub(r'(?s)Private Type \w+.*?End Type', '', host_src)
    host_src = re.sub(r'(?m)^Private Const (\w+)', r'Private \1', host_src)
    pr = [p for p in vbalint.lint([(n, s) for n, s in mods if n != 'Host'] + [('Host', host_src)])
          if p[0] in ('Host', 'Tele') and p[3].lower() not in OM]
    if pr:
        for p in pr[:12]: print('LINT', p)
        raise SystemExit(1)
    longest = max(len(l) for _, s in mods for l in s.split('\n'))
    sheet, styles, pictures = wad_sheet()
    vba = mkxlsm.build_xlsm(out, mods, sheet_xml=sheet, styles=styles, sheet_code=SHEET_CODE, pictures=pictures)
    print('built %s: %.1f MB, %d modules, %d VBA lines, longest line %d chars, vbaProject %.1f MB, %.1fs'
          % (out, os.path.getsize(out) / 1e6, len(mods), sum(s.count('\n') for _, s in mods), longest, len(vba) / 1e6, time.time() - t))

def ref_sums(n=None):
    frames = pickle.load(open('work/ref_frames.pkl', 'rb'))
    sums = []
    for idx, pal in frames:
        c = 0
        for w in struct.unpack('<16000I', idx) + struct.unpack('<256I', pal):
            c = ((c * 31) & 0xFFFFFFFF) ^ w
        sums.append(c)
    return sums

if __name__ == '__main__':
    build()
    if len(sys.argv) > 1 and sys.argv[1] == 'verify':
        import xl
        t = time.time()
        print('open %.1fs' % xl.open_wb('out/doom.xlsm', wait=300), flush=True)
        out, dt = xl.run('doom.xlsm', 'Host.Verify', timeout=900)
        head, _, sums = out.partition('|sums=')
        print('verify wall %.1fs: %s' % (dt, head[:500]))
        if sums:
            got = [int(x) for x in sums.rstrip(',').split(',') if x]
            ref = ref_sums()
            same = 0
            for a, b in zip(got, ref):
                if a != b: break
                same += 1
            print('frames from Excel: %d, reference: %d, identical prefix: %d' % (len(got), len(ref), same))
        if not out.startswith(('DIALOG', 'TIMEOUT')):
            print('close:', xl.close_wb('doom.xlsm') or 'ok')

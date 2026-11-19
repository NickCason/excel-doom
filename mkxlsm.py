"""Build a macro-enabled workbook (.xlsm) directly from VBA source text.

No Excel needed: writes the OLE compound file + MS-OVBA project by hand, with
no compiled cache, so Excel compiles from source on first open.
"""
import struct, sys, zipfile, os

# ---------------------------------------------------------------- compression
def _compress_chunk(chunk):
    n = len(chunk)
    out = bytearray()
    table = {}
    pos = 0
    while pos < n:
        flag_at = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= n:
                break
            best_len = 0
            best_off = 0
            if pos + 3 <= n and pos > 0:
                bc = 4
                while (1 << bc) < pos:
                    bc += 1
                max_len = min((0xFFFF >> bc) + 3, n - pos)
                cands = table.get(chunk[pos:pos + 3])
                if cands:
                    for c in reversed(cands[-24:]):
                        l = 3
                        while l < max_len and chunk[c + l] == chunk[pos + l]:
                            l += 1
                        if l > best_len:
                            best_len = l
                            best_off = pos - c
                            if l == max_len:
                                break
            if best_len >= 3:
                bc = 4
                while (1 << bc) < pos:
                    bc += 1
                tok = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                out += struct.pack('<H', tok)
                flags |= 1 << bit
                step = best_len
            else:
                out.append(chunk[pos])
                step = 1
            for p in range(pos, pos + step):
                if p + 3 <= n:
                    table.setdefault(chunk[p:p + 3], []).append(p)
            pos += step
        out[flag_at] = flags
    return bytes(out)

def ovba_compress(data):
    out = bytearray(b'\x01')
    for cs in range(0, len(data), 4096):
        chunk = data[cs:cs + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) > 4096:
            out += struct.pack('<H', 0x3000 | 4095) + chunk.ljust(4096, b'\0')
        else:
            out += struct.pack('<H', 0xB000 | (len(comp) - 1)) + comp
    return bytes(out)

def ovba_decompress(buf):
    assert buf[0] == 1
    out = bytearray()
    i = 1
    while i < len(buf):
        hdr, = struct.unpack_from('<H', buf, i)
        size = (hdr & 0xFFF) + 3
        end = i + size
        i += 2
        start = len(out)
        if not hdr & 0x8000:
            out += buf[i:i + 4096]
            i = end
            continue
        while i < end:
            flags = buf[i]; i += 1
            for bit in range(8):
                if i >= end:
                    break
                if flags >> bit & 1:
                    tok, = struct.unpack_from('<H', buf, i); i += 2
                    diff = len(out) - start
                    bc = 4
                    while (1 << bc) < diff:
                        bc += 1
                    ln = (tok & (0xFFFF >> bc)) + 3
                    off = (tok >> (16 - bc)) + 1
                    for _ in range(ln):
                        out.append(out[-off])
                else:
                    out.append(buf[i]); i += 1
    return bytes(out)

# ------------------------------------------------------- PROJECT stream crypto
def _encrypt(project_id, data, seed=0x41):
    key = sum(project_id.encode('ascii')) & 0xFF
    out = [seed, seed ^ 2, seed ^ key]
    unenc1, enc1, enc2 = key, seed ^ key, seed ^ 2
    def push(b):
        nonlocal unenc1, enc1, enc2
        e = b ^ ((enc2 + unenc1) & 0xFF)
        out.append(e)
        enc2, enc1, unenc1 = enc1, e, b
    for _ in range((seed & 6) // 2):
        push(0x07)
    for b in struct.pack('<I', len(data)):
        push(b)
    for b in data:
        push(b)
    return ''.join('%02X' % b for b in out)

# ------------------------------------------------------------------ dir stream
def _rec(rid, data=b''):
    return struct.pack('<HI', rid, len(data)) + data

def _dir_stream(modules):
    d = bytearray()
    d += _rec(0x0001, struct.pack('<I', 3))
    d += _rec(0x0002, struct.pack('<I', 0x409))
    d += _rec(0x0014, struct.pack('<I', 0x409))
    d += _rec(0x0003, struct.pack('<H', 1252))
    d += _rec(0x0004, b'VBAProject')
    d += _rec(0x0005) + _rec(0x0040)
    d += _rec(0x0006) + _rec(0x003D)
    d += _rec(0x0007, struct.pack('<I', 0))
    d += _rec(0x0008, struct.pack('<I', 0))
    d += struct.pack('<HIIH', 0x0009, 4, 0x644D8E4A, 0x0012)
    d += _rec(0x000C) + _rec(0x003C)
    refs = [
        ('stdole', r'*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation'),
        ('Office', r'*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library'),
    ]
    for name, libid in refs:
        n = name.encode('ascii')
        d += _rec(0x0016, n) + _rec(0x003E, name.encode('utf-16-le'))
        lb = libid.encode('ascii')
        d += _rec(0x000D, struct.pack('<I', len(lb)) + lb + struct.pack('<IH', 0, 0))
    d += _rec(0x000F, struct.pack('<H', len(modules)))
    d += _rec(0x0013, struct.pack('<H', 0xFFFF))
    for name, kind, _src in modules:
        n = name.encode('ascii')
        u = name.encode('utf-16-le')
        d += _rec(0x0019, n) + _rec(0x0047, u)
        d += _rec(0x001A, n) + _rec(0x0032, u)
        d += _rec(0x001C) + _rec(0x0048)
        d += _rec(0x0031, struct.pack('<I', 0))
        d += _rec(0x001E, struct.pack('<I', 0))
        d += _rec(0x002C, struct.pack('<H', 0xFFFF))
        d += _rec(0x0022 if kind == 'doc' else 0x0021)
        d += _rec(0x002B)
    d += _rec(0x0010)
    return bytes(d)

# ------------------------------------------------------------------ CFB writer
FREE, EOC, FATSECT, DIFSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFC, 0xFFFFFFFF

def _cfb_key(name):
    return (len(name.encode('utf-16-le')), name.upper())

def write_cfb(tree):
    """tree: nested dict; bytes values are streams, dict values are storages."""
    entries = []   # dicts: name,type,left,right,child,data

    def add(name, node, etype):
        idx = len(entries)
        e = dict(name=name, type=etype, left=NOSTREAM, right=NOSTREAM, child=NOSTREAM,
                 data=node if isinstance(node, bytes) else None, color=1)
        entries.append(e)
        if isinstance(node, dict):
            kids = sorted(node.items(), key=lambda kv: _cfb_key(kv[0]))
            ids = [add(k, v, 1 if isinstance(v, dict) else 2) for k, v in kids]
            def build(lo, hi):
                if lo >= hi:
                    return NOSTREAM
                mid = (lo + hi) // 2
                entries[ids[mid]]['left'] = build(lo, mid)
                entries[ids[mid]]['right'] = build(mid + 1, hi)
                return ids[mid]
            e['child'] = build(0, len(ids))
        return idx

    add('Root Entry', tree, 5)

    SS, MSS, CUTOFF = 512, 64, 4096
    mini = bytearray()
    minifat = []
    big = []        # (entry, data)
    for e in entries:
        if e['type'] != 2:
            e['start'], e['size'] = (EOC if e['type'] == 5 else 0), 0
            continue
        data = e['data']
        e['size'] = len(data)
        if len(data) == 0:
            e['start'] = EOC
        elif len(data) < CUTOFF:
            first = len(minifat)
            nsec = (len(data) + MSS - 1) // MSS
            for i in range(nsec):
                minifat.append(first + i + 1 if i < nsec - 1 else EOC)
            mini += data.ljust(nsec * MSS, b'\0')
            e['start'] = first
        else:
            big.append(e)

    def nsect(nbytes):
        return (nbytes + SS - 1) // SS

    dir_bytes_len = len(entries) * 128
    n_dir = nsect(dir_bytes_len)
    minifat_bytes = b''.join(struct.pack('<I', x) for x in minifat)
    n_minifat = nsect(len(minifat_bytes))
    n_mini = nsect(len(mini))
    n_big = sum(nsect(e['size']) for e in big)
    n_data = n_dir + n_minifat + n_mini + n_big
    n_fat = 0
    n_difat = 0
    while True:
        need_fat = (n_data + n_fat + n_difat + 127) // 128
        need_difat = 0 if need_fat <= 109 else (need_fat - 109 + 126) // 127
        if need_fat == n_fat and need_difat == n_difat:
            break
        n_fat, n_difat = need_fat, need_difat
    total = n_data + n_fat + n_difat
    fat = [FREE] * (n_fat * 128)
    sectors = [b''] * total

    cur = 0
    def alloc_chain(data):
        nonlocal cur
        n = nsect(len(data))
        first = cur
        for i in range(n):
            sectors[cur] = data[i * SS:(i + 1) * SS].ljust(SS, b'\0')
            fat[cur] = cur + 1 if i < n - 1 else EOC
            cur += 1
        return first if n else EOC

    dir_start = cur
    cur += n_dir                      # fill later (needs stream starts)
    for i in range(n_dir):
        fat[dir_start + i] = dir_start + i + 1 if i < n_dir - 1 else EOC
    minifat_start = alloc_chain(minifat_bytes.ljust(n_minifat * SS, b'\xff'))
    mini_start = alloc_chain(bytes(mini))
    entries[0]['start'], entries[0]['size'] = mini_start, len(mini)
    for e in big:
        e['start'] = alloc_chain(e['data'])
    fat_start = cur
    for i in range(n_fat):
        fat[cur] = FATSECT
        cur += 1
    difat_start = cur if n_difat else EOC
    for i in range(n_difat):
        fat[cur] = DIFSECT
        cur += 1
    assert cur == total

    d = bytearray()
    for e in entries:
        nm = e['name'].encode('utf-16-le') + b'\0\0'
        d += nm.ljust(64, b'\0')
        d += struct.pack('<HBB', len(nm), e['type'], e['color'])
        d += struct.pack('<III', e['left'], e['right'], e['child'])
        d += b'\0' * 16 + struct.pack('<I', 0) + b'\0' * 16
        d += struct.pack('<IQ', e['start'], e['size'])
    blank = (b'\0' * 64 + struct.pack('<HBB', 0, 0, 0) + struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM)
             + b'\0' * 36 + struct.pack('<IQ', 0, 0))
    while len(d) < n_dir * SS:
        d += blank
    for i in range(n_dir):
        sectors[dir_start + i] = bytes(d[i * SS:(i + 1) * SS])
    fat_bytes = b''.join(struct.pack('<I', x) for x in fat)
    for i in range(n_fat):
        sectors[fat_start + i] = fat_bytes[i * SS:(i + 1) * SS]
    fat_ids = [fat_start + i for i in range(n_fat)]
    extra = fat_ids[109:]
    for i in range(n_difat):
        part = extra[i * 127:(i + 1) * 127]
        part = part + [FREE] * (127 - len(part))
        nxt = difat_start + i + 1 if i < n_difat - 1 else EOC
        sectors[difat_start + i] = b''.join(struct.pack('<I', x) for x in part + [nxt])

    h = bytearray()
    h += bytes.fromhex('D0CF11E0A1B11AE1') + b'\0' * 16
    h += struct.pack('<HHHHH', 0x003E, 3, 0xFFFE, 9, 6) + b'\0' * 6
    h += struct.pack('<III', 0, n_fat, dir_start)
    h += struct.pack('<II', 0, CUTOFF)
    h += struct.pack('<II', minifat_start if n_minifat else EOC, n_minifat)
    h += struct.pack('<II', difat_start, n_difat)
    head = fat_ids[:109] + [FREE] * (109 - len(fat_ids[:109]))
    h += b''.join(struct.pack('<I', x) for x in head)
    assert len(h) == 512
    return bytes(h) + b''.join(sectors)

# ------------------------------------------------------------------- assembly
DOC_ATTRS = '''Attribute VB_Name = "%s"
Attribute VB_Base = "0{%s}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'''

def build_vba_project(std_modules, sheet_code='', extra_sheets=0, workbook_code=''):
    """std_modules: list of (name, source_without_attribute_header). Every worksheet needs a
    document module; extra_sheets adds empty ones named Sheet2, Sheet3, ..."""
    modules = [
        ('ThisWorkbook', 'doc', DOC_ATTRS % ('ThisWorkbook', '00020819-0000-0000-C000-000000000046') + workbook_code),
        ('Sheet1', 'doc', DOC_ATTRS % ('Sheet1', '00020820-0000-0000-C000-000000000046') + sheet_code),
    ]
    for k in range(extra_sheets):
        modules.append(('Sheet%d' % (k + 2), 'doc', DOC_ATTRS % ('Sheet%d' % (k + 2), '00020820-0000-0000-C000-000000000046')))
    for name, src in std_modules:
        modules.append((name, 'std', 'Attribute VB_Name = "%s"\n%s' % (name, src)))
    pid = '{9E3A6A6B-4F0C-4E4E-9C0B-5B1E7D2A6C11}'
    lines = ['ID="%s"' % pid]
    for name, kind, _ in modules:
        lines.append(('Document=%s/&H00000000' if kind == 'doc' else 'Module=%s') % name)
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              'CMG="%s"' % _encrypt(pid, b'\0\0\0\0'),
              'DPB="%s"' % _encrypt(pid, b'\0'),
              'GC="%s"' % _encrypt(pid, b'\xff'),
              '', '[Host Extender Info]',
              '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '']
    project = '\r\n'.join(lines).encode('ascii')
    wm = b''.join(n.encode('ascii') + b'\0' + n.encode('utf-16-le') + b'\0\0' for n, _, _ in modules) + b'\0\0'
    vba = {'_VBA_PROJECT': bytes.fromhex('CC61FFFF000000'),
           'dir': ovba_compress(_dir_stream(modules))}
    for name, _, src in modules:
        text = src.replace('\r\n', '\n').replace('\n', '\r\n').encode('cp1252')
        vba[name] = ovba_compress(text)
    return write_cfb({'VBA': vba, 'PROJECT': project, 'PROJECTwm': wm})

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
WB = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr codeName="ThisWorkbook"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'''
WBRELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'''
SHEET = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr codeName="Sheet1"/><sheetData/></worksheet>'''

def col_name(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s

def sheet_with_values(rows):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr codeName="Sheet1"/><sheetData>']
    for r, row in enumerate(rows, 1):
        out.append('<row r="%d">' % r)
        out.append(''.join('<c r="%s%d"><v>%d</v></c>' % (col_name(c), r, v) for c, v in enumerate(row, 1)))
        out.append('</row>')
    out.append('</sheetData></worksheet>')
    return ''.join(out)

def styles_xml(colours, extra=None):
    """colours: list of (r, g, b). Cell style id k+1 is a solid fill of colours[k].
    extra: more cell styles for ordinary sheets, each a dict with any of
      font=(size, bold, italic, 'RRGGBB'), fill='RRGGBB', border='RRGGBB' (thin, all sides),
      align='horizontal' or 'horizontal,vertical', wrap=True.
    They get the style ids that follow the colours: len(colours) + 1 + position."""
    extra = extra or []
    fills = ''.join('<fill><patternFill patternType="solid"><fgColor rgb="FF%02X%02X%02X"/><bgColor indexed="64"/></patternFill></fill>' % c for c in colours)
    xfs = ''.join('<xf numFmtId="0" fontId="0" fillId="%d" borderId="0" xfId="0" applyFill="1"/>' % (i + 2) for i in range(len(colours)))
    fonts, borders, efills = ['<font><sz val="11"/><name val="Calibri"/></font>'], ['<border><left/><right/><top/><bottom/><diagonal/></border>'], []
    def index(items, xml):
        if xml not in items: items.append(xml)
        return items.index(xml)
    for st in extra:
        fi = bi = 0; fl = 0
        if 'font' in st:
            sz, bold, ital, rgb = st['font']
            fi = index(fonts, '<font>%s%s<sz val="%s"/><color rgb="FF%s"/><name val="Calibri"/></font>' % ('<b/>' if bold else '', '<i/>' if ital else '', sz, rgb))
        if 'fill' in st:
            fl = len(colours) + 2 + index(efills, '<fill><patternFill patternType="solid"><fgColor rgb="FF%s"/><bgColor indexed="64"/></patternFill></fill>' % st['fill'])
        if 'border' in st:
            side = '<%%s style="thin"><color rgb="FF%s"/></%%s>' % st['border']
            bi = index(borders, '<border>' + ''.join(side % (s, s) for s in ('left', 'right', 'top', 'bottom')) + '<diagonal/></border>')
        al = ''
        if 'align' in st or st.get('wrap'):
            h, _, v = st.get('align', '').partition(',')
            al = '<alignment%s%s%s/>' % (' horizontal="%s"' % h if h else '', ' vertical="%s"' % v if v else '', ' wrapText="1"' if st.get('wrap') else '')
        xfs += ('<xf numFmtId="0" fontId="%d" fillId="%d" borderId="%d" xfId="0" applyFont="1" applyFill="1" applyBorder="1"%s>%s</xf>'
                % (fi, fl, bi, ' applyAlignment="1"' if al else '', al)) if al else \
               '<xf numFmtId="0" fontId="%d" fillId="%d" borderId="%d" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' % (fi, fl, bi)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="%d">%s</fonts>'
            '<fills count="%d"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>%s%s</fills>'
            '<borders count="%d">%s</borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="%d"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>%s</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
            % (len(fonts), ''.join(fonts), len(colours) + 2 + len(efills), fills, ''.join(efills), len(borders), ''.join(borders),
               len(colours) + 1 + len(extra), xfs))

def text_shape(name, c0, r0, c1, r1, text, macro=None, size=11, rgb='000000', font='Calibri', shape_id=2):
    """One anchor for drawing_doc: a borderless text box over the given cells (0-based corners).
    Its fill is fully transparent rather than absent, so the whole box takes the click."""
    return ('<xdr:twoCellAnchor editAs="twoCell"><xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            '<xdr:sp%s textlink=""><xdr:nvSpPr><xdr:cNvPr id="%d" name="%s"/><xdr:cNvSpPr/></xdr:nvSpPr>'
            '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr>'
            '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface="%s"/></a:rPr><a:t>%s</a:t></a:r></a:p></xdr:txBody></xdr:sp>'
            '<xdr:clientData/></xdr:twoCellAnchor>'
            % (c0, r0, c1, r1, ' macro="[0]!%s"' % macro if macro else '', shape_id, name, size * 100, rgb, font, font, font, text))

def drawing_doc(anchors):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            + ''.join(anchors) + '</xdr:wsDr>')

def drawing_xml(pics):
    """pics: list of (name, from_col, from_row, to_col, to_row), 0-based cell corners."""
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
    for i, pic in enumerate(pics):
        name, c0, r0, c1, r1 = pic[:5]
        macro = ' macro="[0]!%s"' % pic[5] if len(pic) > 5 and pic[5] else ''
        out.append('<xdr:twoCellAnchor editAs="twoCell"><xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                   '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                   '<xdr:pic%s><xdr:nvPicPr><xdr:cNvPr id="%d" name="%s"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
                   '<xdr:blipFill><a:blip r:embed="rId%d"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
                   '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
                   '<xdr:clientData/></xdr:twoCellAnchor>' % (c0, r0, c1, r1, macro, i + 2, name, i + 1))
    out.append('</xdr:wsDr>')
    return ''.join(out)

def build_xlsm(path, std_modules, sheet_xml=None, styles=None, sheet_code='', pictures=None,
               sheet_name='Sheet1', sheet_hidden=False, extra_sheets=None, active_tab=0, workbook_code=''):
    """pictures: list of (name, png_bytes, from_col, from_row, to_col, to_row) on the first sheet.
    extra_sheets: list of (tab name, sheet xml, drawing xml or None[, hidden]). Sheet k of them must carry
    codeName Sheet<k+2> and, if it has a drawing, <drawing r:id="rId1"/>. active_tab is 0-based."""
    extra_sheets = extra_sheets or []
    vba = build_vba_project(std_modules, sheet_code, len(extra_sheets), workbook_code)
    ct, wb, wbrels = CT, WB, WBRELS
    wb = wb.replace('<sheet name="Sheet1" sheetId="1" r:id="rId1"/>',
                    '<sheet name="%s" sheetId="1"%s r:id="rId1"/>' % (sheet_name, ' state="hidden"' if sheet_hidden else '')
                    + ''.join('<sheet name="%s" sheetId="%d"%s r:id="rId%d"/>' % (sh[0], k + 2, ' state="hidden"' if len(sh) > 3 and sh[3] else '', k + 10)
                              for k, sh in enumerate(extra_sheets)))
    extra_sheets = [tuple(sh[:3]) for sh in extra_sheets]
    for k, (_, _, dr) in enumerate(extra_sheets):
        ct = ct.replace('</Types>', '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' % (k + 2))
        if dr:
            ct = ct.replace('</Types>', '<Override PartName="/xl/drawings/drawing%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>' % (k + 2))
        wbrels = wbrels.replace('</Relationships>', '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/></Relationships>' % (k + 10, k + 2))
    view = '<bookViews><workbookView%s/></bookViews>' % (' firstSheet="%d" activeTab="%d"' % (active_tab, active_tab) if active_tab else '')
    if extra_sheets and not styles:
        wb = wb.replace('<sheets>', view + '<sheets>')
    if pictures:
        ct = ct.replace('<Default Extension="bin"', '<Default Extension="png" ContentType="image/png"/><Default Extension="bin"')
        ct = ct.replace('</Types>', '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    if styles:
        ct = ct.replace('</Types>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
        wbrels = wbrels.replace('</Relationships>', '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        wb = wb.replace('<sheets>', view + '<sheets>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for k, (_, xml, dr) in enumerate(extra_sheets):
            z.writestr('xl/worksheets/sheet%d.xml' % (k + 2), xml)
            if dr:
                z.writestr('xl/drawings/drawing%d.xml' % (k + 2), dr)
                z.writestr('xl/worksheets/_rels/sheet%d.xml.rels' % (k + 2), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing%d.xml"/></Relationships>' % (k + 2))
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', RELS)
        z.writestr('xl/workbook.xml', wb)
        z.writestr('xl/_rels/workbook.xml.rels', wbrels)
        if styles:
            z.writestr('xl/styles.xml', styles)
        if pictures:
            z.writestr('xl/drawings/drawing1.xml', drawing_xml([(p[0],) + tuple(p[2:]) for p in pictures]))
            z.writestr('xl/drawings/_rels/drawing1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       + ''.join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image%d.png"/>' % (i + 1, i + 1) for i in range(len(pictures)))
                       + '</Relationships>')
            z.writestr('xl/worksheets/_rels/sheet1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>')
            for i, p in enumerate(pictures):
                z.writestr('xl/media/image%d.png' % (i + 1), p[1])
        z.writestr('xl/worksheets/sheet1.xml', sheet_xml or SHEET)
        z.writestr('xl/vbaProject.bin', vba)
    return vba

if __name__ == '__main__':
    out = sys.argv[1]
    mods = []
    for p in sys.argv[2:]:
        mods.append((os.path.splitext(os.path.basename(p))[0], open(p).read()))
    vba = build_xlsm(out, mods)
    print('wrote', out, 'vbaProject.bin', len(vba), 'bytes')

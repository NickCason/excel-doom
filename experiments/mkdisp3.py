"""Real Excel timings for the display tricks: fill-call lists are precomputed here from
real frames and replayed by a trivial VBA loop, so only Excel's cost is being measured."""
import os, subprocess, time, pickle
import numpy as np
import mkxlsm, dispsim
from dispsim import simulate, exact_runs, rect_merge

def packg(r):
    x0, y0, x1, y1, c = r
    return int(x0) | (int(x1) << 9) | (int(y0) << 18) | (int(y1) << 26), int(c)

VBA_HEAD = r'''Option Explicit
Private LG() As LongLong, LC() As Long, LN As Long

Private Sub LoadList(ByVal ws As Worksheet, ByVal col As Long)
    Dim v As Variant, i As Long
    LN = CLng(ws.Cells(1, col).Value)
    If LN = 0 Then Exit Sub
    v = ws.Range(ws.Cells(2, col), ws.Cells(LN + 1, col + 1)).Value
    ReDim LG(1 To LN): ReDim LC(1 To LN)
    For i = 1 To LN
        LG(i) = CLngLng(v(i, 1)): LC(i) = CLng(v(i, 2))
    Next i
End Sub

Private Sub PaintList(ByVal ws As Worksheet)
    Dim i As Long, gq As LongLong, x0 As Long, x1 As Long, y0 As Long, y1 As Long
    For i = 1 To LN
        gq = LG(i)
        x0 = CLng(gq And 511^): x1 = CLng((gq \ 512^) And 511^)
        y0 = CLng((gq \ 262144^) And 255^): y1 = CLng((gq \ 67108864^) And 255^)
        If x0 = x1 And y0 = y1 Then
            ws.Cells(y0 + 1, x0 + 1).Interior.Color = LC(i)
        Else
            ws.Range(ws.Cells(y0 + 1, x0 + 1), ws.Cells(y1 + 1, x1 + 1)).Interior.Color = LC(i)
        End If
    Next i
End Sub

Private Function RunScheme(ByVal ws As Worksheet, ByVal initCol As Long, ByVal c1 As Long, ByVal nf As Long, ByVal batch As Boolean) As String
    Dim f As Long, t As Double, tw As Double, tp As Double, calls As Long
    Application.ScreenUpdating = False
    LoadList ws, initCol: PaintList ws
    Application.ScreenUpdating = True: DoEvents
    For f = 0 To nf - 1
        LoadList ws, c1 + f * 2
        t = Timer
        If batch Then Application.ScreenUpdating = False
        PaintList ws
        tw = tw + (Timer - t): t = Timer
        If batch Then Application.ScreenUpdating = True
        DoEvents
        tp = tp + (Timer - t)
        calls = calls + LN
    Next f
    RunScheme = (calls \ nf) & "," & Format(tw / nf, "0.000") & "," & Format(tp / nf, "0.000")
End Function

Private Function PaintFloor(ByVal ws As Worksheet, ByVal z As Long, ByVal batch As Boolean) As String
    Dim f As Long, t As Double
    ThisWorkbook.Windows(1).Zoom = z
    DoEvents
    t = Timer
    For f = 1 To 5
        If batch Then Application.ScreenUpdating = False
        ws.Cells(100, 160).Interior.Color = f * 40
        If batch Then Application.ScreenUpdating = True
        DoEvents
    Next f
    PaintFloor = Format((Timer - t) / 5, "0.000")
End Function

Public Function Go() As String
    Dim ws As Worksheet, s As String, stp As String, r As Range
    On Error GoTo Fail
    stp = "layout"
    Set ws = ThisWorkbook.Worksheets(1)
    Set r = ws.Range(ws.Cells(1, 1), ws.Cells(200, 320))
    Application.ScreenUpdating = False
    r.ColumnWidth = 1
    r.RowHeight = 6
    ThisWorkbook.Windows(1).Zoom = 60
    Application.ScreenUpdating = True
    Application.Goto ws.Cells(1, 1), True
    DoEvents
'''
VBA_TAIL = r'''    stp = "floor"
    s = s & "floor zoom 100 batched,0," & PaintFloor(ws, 100, True) & ",0;"
    s = s & "floor zoom 60 batched,0," & PaintFloor(ws, 60, True) & ",0;"
    s = s & "floor zoom 30 batched,0," & PaintFloor(ws, 30, True) & ",0;"
    s = s & "floor zoom 60 unbatched,0," & PaintFloor(ws, 60, False) & ",0;"
    ThisWorkbook.Windows(1).Zoom = 60
    stp = "nogrid"
    ThisWorkbook.Windows(1).DisplayGridlines = False
    s = s & "floor zoom 60 no gridlines,0," & PaintFloor(ws, 60, True) & ",0;"
    ThisWorkbook.Windows(1).DisplayHeadings = False
    s = s & "floor zoom 60 no gridlines or headings,0," & PaintFloor(ws, 60, True) & ",0;"
    s = s & "tol 8 again with no gridlines or headings," & RunScheme(ws, %(i)d, %(c8)d, 3, True) & ";"
    s = s & "tol 16 again with no gridlines or headings," & RunScheme(ws, %(i)d, %(c16)d, 3, True) & ";"
    Go = s
    Exit Function
Fail:
    Application.ScreenUpdating = True
    Go = s & "ERROR at " & stp & ": " & Err.Number & " " & Err.Description
End Function
'''

if __name__ == '__main__':
    ticks = list(range(300, 541, 8))
    sets = dict(base=dispsim.capture(10, 0, ticks), low=dispsim.capture(10, 1, ticks), sb8=dispsim.capture(8, 0, ticks),
                sb8low=dispsim.capture(8, 1, ticks))
    K = slice(17, 21)
    cols = {}; nextcol = [401]; cells = {}
    def put(lst):
        c = nextcol[0]; nextcol[0] += 2
        cells.setdefault(1, {})[c] = len(lst)
        for i, r in enumerate(lst):
            g, col = packg(r)
            row = cells.setdefault(i + 2, {}); row[c] = g; row[c + 1] = col
        return c
    init = {}
    for name, fr in sets.items():
        f0 = fr[K][0]
        init[name] = put(rect_merge(exact_runs(f0 ^ 255, f0)))
    plan = [('changed runs (baseline)', 'base', dict(rect=False), True),
            ('rectangles exact', 'base', dict(), True),
            ('near-colour tol 8', 'base', dict(T=8), True),
            ('near-colour tol 16', 'base', dict(T=16), True),
            ('low detail exact', 'low', dict(), True),
            ('low detail + tol 8', 'low', dict(T=8), True),
            ('small view 8 exact', 'sb8', dict(), True),
            ('small view 8 + low detail + tol 8', 'sb8low', dict(T=8), True),
            ('near-colour tol 8 unbatched', 'base', dict(T=8), False)]
    body = []; first = {}
    for label, key, kw, batch in plan:
        res = simulate(label, sets[key][K], 'full', **kw)
        c1 = None
        for lst in res['rects']:
            c = put(lst)
            c1 = c1 or c
        first[label] = c1
        body.append('    stp = "%s"\n    s = s & "%s," & RunScheme(ws, %d, %d, %d, %s) & ";"\n'
                    % (label, label, init[key], c1, len(res['rects']), 'True' if batch else 'False'))
    xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr codeName="Sheet1"/><sheetData>']
    for r in sorted(cells):
        xml.append('<row r="%d">' % r + ''.join('<c r="%s%d"><v>%d</v></c>' % (mkxlsm.col_name(c), r, v) for c, v in sorted(cells[r].items())) + '</row>')
    xml.append('</sheetData></worksheet>')
    mkxlsm.build_xlsm('out/disp3.xlsm', [('Disp', VBA_HEAD + ''.join(body) + VBA_TAIL % dict(i=init['base'], c8=first['near-colour tol 8'], c16=first['near-colour tol 16']))], sheet_xml=''.join(xml))
    print('built out/disp3.xlsm %.1f MB, %d list columns' % (os.path.getsize('out/disp3.xlsm') / 1e6, (nextcol[0] - 401) // 2))
    import xl
    front = xl.osa('tell application "System Events" to get name of first process whose frontmost is true', timeout=20)
    try:
        print('open %.1fs' % xl.open_wb('out/disp3.xlsm'))
        subprocess.run(['open', '-a', xl.APP]); time.sleep(1.5)
        out, dt = xl.run('disp3.xlsm', 'Disp.Go', timeout=400)
        print('wall %.1fs' % dt)
        print('%-36s %7s %9s %9s %9s' % ('scheme', 'calls', 'write ms', 'paint ms', 'total ms'))
        for part in out.split(';'):
            if not part: continue
            f = part.rsplit(',', 3)
            if len(f) == 4 and 'ERROR' not in part:
                print('%-36s %7s %9.0f %9.0f %9.0f' % (f[0], f[1], float(f[2]) * 1000, float(f[3]) * 1000, (float(f[2]) + float(f[3])) * 1000))
            else:
                print('   ', part[:300])
        if not out.startswith(('DIALOG', 'TIMEOUT')):
            print('close:', xl.close_wb('disp3.xlsm') or 'ok')
    finally:
        if front and '!ERR' not in front and 'Excel' not in front:
            xl.osa('tell application "%s" to activate' % front, timeout=20)
            print('refocused', front)
        else:
            print('front app was:', front[:60])

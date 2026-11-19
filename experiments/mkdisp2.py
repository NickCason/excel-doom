"""Display cost comparison: two real Doom frames, several ways of getting them into cells."""
import os, sys, subprocess, time
from wasmtime import Engine, Store, Module, Linker, WasiConfig, FuncType, ValType
import mkxlsm

def grab_frames(ticks):
    engine = Engine(); store = Store(engine)
    wasi = WasiConfig(); wasi.preopen_dir('vendor/wad', '.'); store.set_wasi(wasi)
    linker = Linker(engine); linker.define_wasi()
    last = {}
    def draw(ptr, w, h):
        last['f'] = bytes(last['mem'].read(store, ptr, ptr + w * h * 4))
    linker.define_func('env', 'draw_frame', FuncType([ValType.i32()] * 3, []), draw)
    inst = linker.instantiate(store, Module.from_file(engine, 'build/doom.wasm'))
    ex = inst.exports(store); last['mem'] = ex['memory']
    ex['_initialize'](store); ex['dg_start'](store)
    out = []
    for i in range(max(ticks) + 1):
        ex['dg_tick'](store)
        if i in ticks:
            out.append(last['f'])
    return out

def rgb_rows(raw):
    rows = []
    for y in range(200):
        row = []
        for x in range(320):
            b, g, r, _ = raw[(y * 320 + x) * 4:(y * 320 + x) * 4 + 4]
            row += [r, g, b]
        rows.append(row)
    return rows

VBA = r'''Option Explicit
Private Const PW As Long = 320, PH As Long = 200

Private Sub AddScale(ByVal r As Range, ByVal c As Long)
    Dim cs As ColorScale
    Set cs = r.FormatConditions.AddColorScale(2)
    cs.ColorScaleCriteria(1).Type = xlConditionValueNumber
    cs.ColorScaleCriteria(1).Value = 0
    cs.ColorScaleCriteria(1).FormatColor.Color = 0
    cs.ColorScaleCriteria(2).Type = xlConditionValueNumber
    cs.ColorScaleCriteria(2).Value = 255
    cs.ColorScaleCriteria(2).FormatColor.Color = c
End Sub

Public Function ShowAt(ByVal col As Long) As String
    Application.Goto ThisWorkbook.Worksheets(1).Cells(1, col), True
    DoEvents
    ShowAt = "ok"
End Function
Public Function ShowGrey() As String
    ShowGrey = ShowAt(1001)
End Function

Public Function Go() As String
    Dim ws As Worksheet, r1 As Range, r2 As Range, r3 As Range, rh As Range
    Dim uR As Range, uG As Range, uB As Range
    Dim fa As Variant, fb As Variant
    Dim ga() As Variant, gb() As Variant, ha() As Variant, hb() As Variant
    Dim da() As Double, db() As Double
    Dim ca() As Long, cb() As Long, cc() As Long, pp() As Long
    Dim x As Long, y As Long, f As Long, n As Long, c As Long, x0 As Long
    Dim pr As Long, pg As Long, pb As Long
    Dim t As Double, tw As Double, tp As Double
    Dim s As String, stp As String
    On Error GoTo Fail
    stp = "setup"
    Set ws = ThisWorkbook.Worksheets(1)
    Set r1 = ws.Range(ws.Cells(1, 1), ws.Cells(PH, PW * 3))
    Set rh = ws.Range(ws.Cells(1, 1), ws.Cells(PH \ 2, (PW \ 2) * 3))
    Set r2 = ws.Range(ws.Cells(1, 1001), ws.Cells(PH, 1000 + PW))
    Set r3 = ws.Range(ws.Cells(1, 1401), ws.Cells(PH, 1400 + PW))
    fa = r1.Value
    fb = ws.Range(ws.Cells(301, 1), ws.Cells(300 + PH, PW * 3)).Value
    ReDim ga(1 To PH, 1 To PW): ReDim gb(1 To PH, 1 To PW)
    ReDim ha(1 To PH \ 2, 1 To (PW \ 2) * 3): ReDim hb(1 To PH \ 2, 1 To (PW \ 2) * 3)
    ReDim da(1 To PH, 1 To PW * 3): ReDim db(1 To PH, 1 To PW * 3)
    ReDim ca(1 To PH, 1 To PW): ReDim cb(1 To PH, 1 To PW)
    n = 0
    For y = 1 To PH
        For x = 1 To PW
            pr = CLng(fa(y, x * 3 - 2)): pg = CLng(fa(y, x * 3 - 1)): pb = CLng(fa(y, x * 3))
            ga(y, x) = (pr * 30 + pg * 59 + pb * 11) \ 100
            ca(y, x) = RGB(pr, pg, pb)
            da(y, x * 3 - 2) = pr: da(y, x * 3 - 1) = pg: da(y, x * 3) = pb
            pr = CLng(fb(y, x * 3 - 2)): pg = CLng(fb(y, x * 3 - 1)): pb = CLng(fb(y, x * 3))
            gb(y, x) = (pr * 30 + pg * 59 + pb * 11) \ 100
            cb(y, x) = RGB(pr, pg, pb)
            db(y, x * 3 - 2) = pr: db(y, x * 3 - 1) = pg: db(y, x * 3) = pb
            If ca(y, x) <> cb(y, x) Then n = n + 1
        Next x
    Next y
    s = s & "info_changed_pixels," & n & ";"
    For y = 1 To PH \ 2
        For x = 1 To PW \ 2
            ha(y, x * 3 - 2) = fa(y * 2, x * 6 - 2): ha(y, x * 3 - 1) = fa(y * 2, x * 6 - 1): ha(y, x * 3) = fa(y * 2, x * 6)
            hb(y, x * 3 - 2) = fb(y * 2, x * 6 - 2): hb(y, x * 3 - 1) = fb(y * 2, x * 6 - 1): hb(y, x * 3) = fb(y * 2, x * 6)
        Next x
    Next y

    stp = "layout"
    Application.ScreenUpdating = False
    r1.NumberFormat = ";;;": r2.NumberFormat = ";;;"
    r1.ColumnWidth = 0.33: r2.ColumnWidth = 1: r3.ColumnWidth = 1
    r1.RowHeight = 6
    ThisWorkbook.Windows(1).Zoom = 60
    For x = 0 To PW - 1
        If x = 0 Then
            Set uR = ws.Range(ws.Cells(1, 1), ws.Cells(PH, 1))
            Set uG = ws.Range(ws.Cells(1, 2), ws.Cells(PH, 2))
            Set uB = ws.Range(ws.Cells(1, 3), ws.Cells(PH, 3))
        Else
            Set uR = Union(uR, ws.Range(ws.Cells(1, x * 3 + 1), ws.Cells(PH, x * 3 + 1)))
            Set uG = Union(uG, ws.Range(ws.Cells(1, x * 3 + 2), ws.Cells(PH, x * 3 + 2)))
            Set uB = Union(uB, ws.Range(ws.Cells(1, x * 3 + 3), ws.Cells(PH, x * 3 + 3)))
        End If
    Next x
    AddScale uR, RGB(255, 0, 0)
    AddScale uG, RGB(0, 255, 0)
    AddScale uB, RGB(0, 0, 255)
    AddScale r2, RGB(255, 255, 255)
    Application.ScreenUpdating = True
    Application.Goto ws.Cells(1, 1), True
    DoEvents

    stp = "rgb_plain"
    t = Timer
    For f = 1 To 3
        If (f And 1) = 1 Then r1.Value = fb Else r1.Value = fa
        DoEvents
    Next f
    s = s & "rgb3_values_no_batching," & Format((Timer - t) / 3, "0.000") & ";"

    stp = "rgb_su"
    tw = 0: tp = 0
    For f = 1 To 4
        t = Timer
        Application.ScreenUpdating = False
        If (f And 1) = 1 Then r1.Value = fb Else r1.Value = fa
        tw = tw + (Timer - t): t = Timer
        Application.ScreenUpdating = True
        DoEvents
        tp = tp + (Timer - t)
    Next f
    s = s & "rgb3_values_batched_write," & Format(tw / 4, "0.000") & ";rgb3_values_batched_paint," & Format(tp / 4, "0.000") & ";"

    stp = "rgb_double"
    tw = 0: tp = 0
    For f = 1 To 4
        t = Timer
        Application.ScreenUpdating = False
        If (f And 1) = 1 Then r1.Value2 = db Else r1.Value2 = da
        tw = tw + (Timer - t): t = Timer
        Application.ScreenUpdating = True
        DoEvents
        tp = tp + (Timer - t)
    Next f
    s = s & "rgb3_typed_array_write," & Format(tw / 4, "0.000") & ";rgb3_typed_array_paint," & Format(tp / 4, "0.000") & ";"

    stp = "rgb_half"
    tw = 0: tp = 0
    For f = 1 To 4
        t = Timer
        Application.ScreenUpdating = False
        If (f And 1) = 1 Then rh.Value = hb Else rh.Value = ha
        tw = tw + (Timer - t): t = Timer
        Application.ScreenUpdating = True
        DoEvents
        tp = tp + (Timer - t)
    Next f
    s = s & "rgb3_halfres_write," & Format(tw / 4, "0.000") & ";rgb3_halfres_paint," & Format(tp / 4, "0.000") & ";"
    r1.Value = fa

    stp = "grey"
    Application.Goto ws.Cells(1, 1001), True
    DoEvents
    tw = 0: tp = 0
    For f = 1 To 4
        t = Timer
        Application.ScreenUpdating = False
        If (f And 1) = 1 Then r2.Value = gb Else r2.Value = ga
        tw = tw + (Timer - t): t = Timer
        Application.ScreenUpdating = True
        DoEvents
        tp = tp + (Timer - t)
    Next f
    s = s & "grey1_values_write," & Format(tw / 4, "0.000") & ";grey1_values_paint," & Format(tp / 4, "0.000") & ";"

    stp = "fill_all"
    Application.Goto ws.Cells(1, 1401), True
    DoEvents
    tw = 0: tp = 0
    For f = 1 To 2
        If f = 1 Then cc = cb Else cc = ca
        t = Timer
        Application.ScreenUpdating = False
        For y = 1 To PH
            For x = 1 To PW
                ws.Cells(y, 1400 + x).Interior.Color = cc(y, x)
            Next x
        Next y
        tw = tw + (Timer - t): t = Timer
        Application.ScreenUpdating = True
        DoEvents
        tp = tp + (Timer - t)
    Next f
    s = s & "fill_every_cell_write," & Format(tw / 2, "0.000") & ";fill_every_cell_paint," & Format(tp / 2, "0.000") & ";"

    stp = "fill_diff"
    tw = 0: tp = 0
    For f = 1 To 2
        If f = 1 Then
            cc = cb: pp = ca
        Else
            cc = ca: pp = cb
        End If
        t = Timer
        Application.ScreenUpdating = False
        For y = 1 To PH
            For x = 1 To PW
                If cc(y, x) <> pp(y, x) Then ws.Cells(y, 1400 + x).Interior.Color = cc(y, x)
            Next x
        Next y
        tw = tw + (Timer - t): t = Timer
        Application.ScreenUpdating = True
        DoEvents
        tp = tp + (Timer - t)
    Next f
    s = s & "fill_changed_cells_write," & Format(tw / 2, "0.000") & ";fill_changed_cells_paint," & Format(tp / 2, "0.000") & ";"

    stp = "fill_runs"
    tw = 0: tp = 0: n = 0
    For f = 1 To 2
        If f = 1 Then cc = cb Else cc = ca
        t = Timer
        Application.ScreenUpdating = False
        For y = 1 To PH
            x0 = 1
            For x = 2 To PW
                If cc(y, x) <> cc(y, x0) Then
                    ws.Range(ws.Cells(y, 1400 + x0), ws.Cells(y, 1399 + x)).Interior.Color = cc(y, x0)
                    n = n + 1: x0 = x
                End If
            Next x
            ws.Range(ws.Cells(y, 1400 + x0), ws.Cells(y, 1400 + PW)).Interior.Color = cc(y, x0)
            n = n + 1
        Next y
        tw = tw + (Timer - t): t = Timer
        Application.ScreenUpdating = True
        DoEvents
        tp = tp + (Timer - t)
    Next f
    s = s & "fill_runs_write," & Format(tw / 2, "0.000") & ";fill_runs_paint," & Format(tp / 2, "0.000") & ";info_runs_per_frame," & (n \ 2) & ";"

    stp = "fill_diff_runs"
    tw = 0: tp = 0: n = 0
    For f = 1 To 2
        If f = 1 Then
            cc = cb: pp = ca
        Else
            cc = ca: pp = cb
        End If
        t = Timer
        Application.ScreenUpdating = False
        For y = 1 To PH
            x = 1
            Do While x <= PW
                If cc(y, x) <> pp(y, x) Then
                    x0 = x: c = cc(y, x)
                    Do While x < PW
                        If cc(y, x + 1) <> c Then Exit Do
                        x = x + 1
                    Loop
                    ws.Range(ws.Cells(y, 1400 + x0), ws.Cells(y, 1400 + x)).Interior.Color = c
                    n = n + 1
                End If
                x = x + 1
            Loop
        Next y
        tw = tw + (Timer - t): t = Timer
        Application.ScreenUpdating = True
        DoEvents
        tp = tp + (Timer - t)
    Next f
    s = s & "fill_changed_runs_write," & Format(tw / 2, "0.000") & ";fill_changed_runs_paint," & Format(tp / 2, "0.000") & ";info_changed_runs_per_frame," & (n \ 2) & ";"
    Go = s
    Exit Function
Fail:
    Application.ScreenUpdating = True
    Go = s & "ERROR at " & stp & ": " & Err.Number & " " & Err.Description
End Function
'''

if __name__ == '__main__':
    fa, fb = grab_frames([450, 458])
    rows = rgb_rows(fa) + [[] for _ in range(100)] + rgb_rows(fb)
    mkxlsm.build_xlsm('out/disp2.xlsm', [('Disp', VBA)], sheet_xml=mkxlsm.sheet_with_values(rows))
    print('built out/disp2.xlsm', os.path.getsize('out/disp2.xlsm'), 'bytes')
    import xl
    front = xl.osa('tell application "System Events" to get name of first process whose frontmost is true', timeout=20)
    try:
        print('open %.1fs' % xl.open_wb('out/disp2.xlsm'))
        subprocess.run(['open', '-a', xl.APP]); time.sleep(1.5)
        out, dt = xl.run('disp2.xlsm', 'Disp.Go', timeout=400)
        print('wall %.1fs' % dt)
        for part in out.split(';'):
            if part: print('   ', part[:220])
        if not out.startswith(('DIALOG', 'TIMEOUT')):
            time.sleep(1)
            subprocess.run(['screencapture', '-x', 'out/disp2_fill.png'])
            xl.run('disp2.xlsm', 'Disp.ShowGrey', timeout=30); time.sleep(1)
            subprocess.run(['screencapture', '-x', 'out/disp2_grey.png'])
            print('close:', xl.close_wb('disp2.xlsm') or 'ok')
    finally:
        if front and '!ERR' not in front and 'Excel' not in front:
            xl.osa('tell application "%s" to activate' % front, timeout=20)
            print('refocused', front)

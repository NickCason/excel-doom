"""Display test: a real Doom frame stored as cell values (3 cells per pixel: R, G, B),
coloured by conditional-format colour scales, plus timing of whole-frame rewrites."""
import os, struct
from wasmtime import Config, Engine, Store, Module, Linker, WasiConfig, FuncType, ValType
import mkxlsm, vbalint

def grab_frame(tick=450):
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
    for _ in range(tick):
        ex['dg_tick'](store)
    return last['f']

VBA = r'''Option Explicit
Private Const W As Long = 320, H As Long = 200

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

Public Function Go() As String
    Dim ws As Worksheet, rng As Range, t As Double, s As String, f As Long, x As Long, y As Long
    Dim rR As Range, rG As Range, rB As Range, v As Variant, keep As Variant, n As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(1)
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(H, W * 3))
    t = Timer
    Application.ScreenUpdating = False
    rng.NumberFormat = ";;;"
    rng.ColumnWidth = 0.33
    rng.RowHeight = 6
    ThisWorkbook.Windows(1).Zoom = 60
    s = s & "layout," & Format(Timer - t, "0.000") & "," & Err.Number & ";": Err.Clear
    t = Timer
    For x = 0 To W - 1
        If x = 0 Then
            Set rR = ws.Range(ws.Cells(1, 1), ws.Cells(H, 1))
            Set rG = ws.Range(ws.Cells(1, 2), ws.Cells(H, 2))
            Set rB = ws.Range(ws.Cells(1, 3), ws.Cells(H, 3))
        Else
            Set rR = Union(rR, ws.Range(ws.Cells(1, x * 3 + 1), ws.Cells(H, x * 3 + 1)))
            Set rG = Union(rG, ws.Range(ws.Cells(1, x * 3 + 2), ws.Cells(H, x * 3 + 2)))
            Set rB = Union(rB, ws.Range(ws.Cells(1, x * 3 + 3), ws.Cells(H, x * 3 + 3)))
        End If
    Next x
    AddScale rR, RGB(255, 0, 0)
    AddScale rG, RGB(0, 255, 0)
    AddScale rB, RGB(0, 0, 255)
    s = s & "colour_scales," & Format(Timer - t, "0.000") & "," & Err.Number & " " & Err.Description & ";": Err.Clear
    Application.ScreenUpdating = True
    DoEvents
    t = Timer
    keep = rng.Value
    s = s & "bulk_read_192k," & Format(Timer - t, "0.000") & "," & Err.Number & ";": Err.Clear
    v = keep
    t = Timer
    For f = 1 To 10
        For y = 1 To H
            v(y, ((f * 37 + y) Mod (W * 3)) + 1) = 255
        Next y
        rng.Value = v
        DoEvents
    Next f
    s = s & "bulk_write_192k_x10," & Format(Timer - t, "0.000") & "," & Err.Number & ";": Err.Clear
    t = Timer
    For f = 1 To 10
        Application.ScreenUpdating = False
        rng.Value = v
        Application.ScreenUpdating = True
        DoEvents
    Next f
    s = s & "bulk_write_192k_x10_su," & Format(Timer - t, "0.000") & "," & Err.Number & ";": Err.Clear
    t = Timer
    n = 0
    For y = 1 To 20
        For x = 1 To 250
            ws.Cells(H + 5 + y, x).Interior.Color = (x * 65793) And 16777215
            n = n + 1
        Next x
    Next y
    DoEvents
    s = s & "percell_fill_5000," & Format(Timer - t, "0.000") & "," & Err.Number & ";": Err.Clear
    rng.Value = keep
    DoEvents
    Go = s
End Function
'''

if __name__ == '__main__':
    raw = grab_frame()
    rows = []
    for y in range(200):
        row = []
        for x in range(320):
            b, g, r, _ = raw[(y * 320 + x) * 4:(y * 320 + x) * 4 + 4]
            row += [r, g, b]
        rows.append(row)
    mods = [('Disp', VBA)]
    pr = [p for p in vbalint.lint(mods) if p[3] not in ('Worksheet', 'Range', 'Variant', 'ColorScale', 'ThisWorkbook', 'Application',
          'ScreenUpdating', 'Worksheets', 'Cells', 'NumberFormat', 'ColumnWidth', 'RowHeight', 'Windows', 'Zoom', 'Union', 'RGB',
          'DoEvents', 'Value', 'Interior', 'Color', 'FormatConditions', 'AddColorScale', 'ColorScaleCriteria', 'Type',
          'xlConditionValueNumber', 'FormatColor', 'Nothing', 'Clear', 'Resume')]
    print('lint:', pr[:5])
    mkxlsm.build_xlsm('out/disp.xlsm', mods, sheet_xml=mkxlsm.sheet_with_values(rows))
    print('built out/disp.xlsm', os.path.getsize('out/disp.xlsm'), 'bytes')

"""Build the three benchmark workbooks: translated (Long-array memory), translated
(Byte-array memory), emulated; plus a micro-benchmark of VBA primitives."""
import sys
import mkxlsm, w2vba, w2interp, vbalint

def checked(path, mods):
    pr = vbalint.lint(mods)
    if pr:
        print('LINT PROBLEMS in', path)
        for p in pr[:10]:
            print('   ', p)
        raise SystemExit(1)
    mkxlsm.build_xlsm(path, mods)

DRIVER = r'''Option Explicit
Public Function RunAll(ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long) As String
    Dim t As Double, r As Long, s As String
    On Error GoTo Fail
    RtInit
    x_k_init
    t = Timer: r = x_k_render(n1): s = s & "render," & r & "," & Format(Timer - t, "0.000") & ";"
    t = Timer: r = x_k_logic(n2): s = s & "logic," & r & "," & Format(Timer - t, "0.000") & ";"
    t = Timer: r = x_k_bsp(n3): s = s & "bsp," & r & "," & Format(Timer - t, "0.000") & ";"
    RunAll = s
    Exit Function
Fail:
    RunAll = s & "ERROR " & Err.Number & " " & Err.Description
End Function
Public Function Go() As String
    Go = RunAll(%d, %d, %d)
End Function
'''

MICRO = r'''Option Explicit
Private M(0 To 65535) As Long
Private B(0 To 262143) As Byte
Private MK(0 To 3) As Long, DV(0 To 3) As Long
Private Function Leaf(ByVal a As Long, ByVal b As Long) As Long
    Leaf = a Xor b
End Function
Public Function MicroRun(ByVal n As Long) As String
    Dim i As Long, a As Long, c As Long, t As Double, s As String, w As LongLong
    On Error GoTo Fail
    MK(0) = &HFF&: MK(1) = &HFF00&: MK(2) = &HFF0000: MK(3) = &HFF000000
    DV(0) = 1: DV(1) = 256: DV(2) = 65536: DV(3) = 16777216
    For i = 0 To 65535: M(i) = (i * 31) Xor 12345: Next i
    For i = 0 To 262143: B(i) = i And 255: Next i
    a = 12345
    t = Timer: For i = 1 To n: Next i
    s = s & "empty_loop," & Format(Timer - t, "0.000") & ";"
    t = Timer: For i = 1 To n: c = a Xor i: Next i
    s = s & "xor_assign," & Format(Timer - t, "0.000") & ";"
    t = Timer: For i = 1 To n: c = (a And 65535) + (i And 65535): Next i
    s = s & "raw_add," & Format(Timer - t, "0.000") & ";"
    t = Timer: For i = 1 To n: c = CLng((((CLngLng(c) + i) And 4294967295^) Xor 2147483648^) - 2147483648^): Next i
    s = s & "wrapped_add," & Format(Timer - t, "0.000") & ";"
    t = Timer: For i = 1 To n: c = CLng((((CLngLng(c) * i) And 4294967295^) Xor 2147483648^) - 2147483648^): Next i
    s = s & "wrapped_mul," & Format(Timer - t, "0.000") & ";"
    t = Timer: For i = 1 To n: c = M((i And 262140) \ 4): Next i
    s = s & "word_load_longarr," & Format(Timer - t, "0.000") & ";"
    t = Timer: For i = 1 To n: a = i And 262143: c = ((M(a \ 4) And MK(a And 3)) \ DV(a And 3)) And 255: Next i
    s = s & "byte_load_longarr," & Format(Timer - t, "0.000") & ";"
    t = Timer: For i = 1 To n: c = B(i And 262143): Next i
    s = s & "byte_load_bytearr," & Format(Timer - t, "0.000") & ";"
    t = Timer: For i = 1 To n: a = i And 262140: c = B(a) Or B(a + 1) * 256& Or B(a + 2) * 65536 Or (B(a + 3) And 127) * 16777216 Or ((B(a + 3) > 127) And &H80000000): Next i
    s = s & "word_load_bytearr," & Format(Timer - t, "0.000") & ";"
    t = Timer: For i = 1 To n: c = Leaf(a, i): Next i
    s = s & "function_call," & Format(Timer - t, "0.000") & ";"
    t = Timer: For i = 1 To n
        If (i And 1) = 0 Then GoTo Skip
        c = c + 1
Skip:
    Next i
    s = s & "if_goto," & Format(Timer - t, "0.000") & ";"
    MicroRun = s
    Exit Function
Fail:
    MicroRun = s & "ERROR " & Err.Number & " " & Err.Description
End Function
'''

N = (40, 2000, 4000)

if __name__ == '__main__':
    wasm_path = 'build/bench.wasm'
    DRIVER = DRIVER % N
    for tag, mode, bh in [('L', 'L', False), ('B', 'B', False), ('Lh', 'L', True)]:
        T = w2vba.Translator(wasm_path, mode, byte_helper=bh)
        mods, stats = T.build()
        checked('out/bench%s.xlsm' % tag, mods + [('Driver', DRIVER)])
        print('bench%s: %d wasm instrs -> %d VBA lines, failed=%s' % (tag, stats['wasm_instrs'], stats['vba_lines'], stats['failed']))
    mods, stats = w2interp.build(wasm_path)
    checked('out/benchI.xlsm', mods + [('Driver', DRIVER)])
    print('benchI:', stats)
    mkxlsm.build_xlsm('out/micro.xlsm', [('Micro', MICRO + 'Public Function Go() As String\nGo = MicroRun(5000000)\nEnd Function\n')])

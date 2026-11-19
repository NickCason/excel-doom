"""Hand-written VBA runtime for translated modules, in the subset vbaemu.py runs.
RT = integer/memory helpers. RTF = float bit-casts (LSet; native in the emulator).
WASI = the file/console calls Doom's C library makes, backed by an in-memory WAD."""

RT = r'''Option Explicit
Public M() As Long
Public MK(0 To 3) As Long, DV(0 To 3) As Long, NM(0 To 3) As Long, MH(0 To 2) As Long, NH(0 To 2) As Long
Public BT(0 To 1023) As Long
Public P2(0 To 63) As LongLong
Public MemCap As Long

Public Sub Trap()
    Err.Raise vbObjectError + 77, , "wasm trap"
End Sub

Public Sub RtInit(ByVal pages As Long)
    Dim k As Long, v As Long
    MemCap = pages
    ReDim M(0 To pages * 16384 - 1)
    MK(0) = &HFF&: MK(1) = &HFF00&: MK(2) = &HFF0000: MK(3) = &HFF000000
    DV(0) = 1: DV(1) = 256: DV(2) = 65536: DV(3) = 16777216
    For k = 0 To 3
        NM(k) = Not MK(k)
    Next k
    MH(0) = &HFFFF&: MH(2) = &HFFFF0000
    NH(0) = &HFFFF0000: NH(2) = &HFFFF&
    For k = 0 To 2
        For v = 0 To 255
            BT(k * 256 + v) = v * DV(k)
        Next v
    Next k
    For v = 0 To 255
        BT(768 + v) = ((v And 127) * 16777216) Or ((v > 127) And &H80000000)
    Next v
    P2(0) = 1
    For k = 1 To 62
        P2(k) = P2(k - 1) * 2
    Next k
    P2(63) = -9223372036854775807^ - 1^
End Sub

Public Function W32(ByVal x As LongLong) As Long
    W32 = CLng(((x And 4294967295^) Xor 2147483648^) - 2147483648^)
End Function
Public Function ShrS32(ByVal a As Long, ByVal k As Long) As Long
    Dim x As LongLong
    k = k And 31: x = a
    ShrS32 = CLng((x - (x And (P2(k) - 1^))) \ P2(k))
End Function
Public Function ShrU32(ByVal a As Long, ByVal k As Long) As Long
    ShrU32 = W32((CLngLng(a) And 4294967295^) \ P2(k And 31))
End Function
Public Function DivU32(ByVal a As Long, ByVal b As Long) As Long
    DivU32 = W32((CLngLng(a) And 4294967295^) \ (CLngLng(b) And 4294967295^))
End Function
Public Function RemU32(ByVal a As Long, ByVal b As Long) As Long
    RemU32 = W32((CLngLng(a) And 4294967295^) Mod (CLngLng(b) And 4294967295^))
End Function
Public Function Rotl32(ByVal a As Long, ByVal k As Long) As Long
    Dim u As LongLong
    k = k And 31: u = CLngLng(a) And 4294967295^
    If k = 0 Then
        Rotl32 = a
    Else
        Rotl32 = W32((u * P2(k)) Or (u \ P2(32 - k)))
    End If
End Function
Public Function Rotr32(ByVal a As Long, ByVal k As Long) As Long
    Rotr32 = Rotl32(a, 32 - (k And 31))
End Function
Public Function Clz32(ByVal a As Long) As Long
    Dim u As LongLong, n As Long
    If a = 0 Then
        Clz32 = 32
        Exit Function
    End If
    u = CLngLng(a) And 4294967295^
    Do While u < 2147483648^
        u = u * 2: n = n + 1
    Loop
    Clz32 = n
End Function
Public Function Ctz32(ByVal a As Long) As Long
    Dim u As LongLong, n As Long
    If a = 0 Then
        Ctz32 = 32
        Exit Function
    End If
    u = CLngLng(a) And 4294967295^
    Do While (u And 1^) = 0
        u = u \ 2: n = n + 1
    Loop
    Ctz32 = n
End Function
Public Function Popcnt32(ByVal a As Long) As Long
    Dim u As LongLong, n As Long
    u = CLngLng(a) And 4294967295^
    Do While u <> 0
        n = n + CLng(u And 1^): u = u \ 2
    Loop
    Popcnt32 = n
End Function

Public Function LD8(ByVal a As Long) As Long
    LD8 = ((M(a \ 4) And MK(a And 3)) \ DV(a And 3)) And 255
End Function
Public Function LD16(ByVal a As Long) As Long
    LD16 = LD8(a) Or LD8(a + 1) * 256&
End Function
Public Function LD16A(ByVal a As Long) As Long
    If (a And 1) <> 0 Then Trap
    LD16A = LD16(a)
End Function
Public Function LD32(ByVal a As Long) As Long
    If (a And 3) = 0 Then
        LD32 = M(a \ 4)
    Else
        LD32 = LD8(a) Or LD8(a + 1) * 256& Or LD8(a + 2) * 65536 Or BT(768 + LD8(a + 3))
    End If
End Function
Public Function LD64(ByVal a As Long) As LongLong
    LD64 = (CLngLng(LD32(a + 4)) * 4294967296^) Or (CLngLng(LD32(a)) And 4294967295^)
End Function
Public Function AL4(ByVal a As Long) As Long
    If (a And 3) <> 0 Then Trap
    AL4 = a \ 4
End Function
Public Sub ST8(ByVal a As Long, ByVal v As Long)
    Dim i As Long
    i = a \ 4
    M(i) = (M(i) And NM(a And 3)) Or BT((a And 3) * 256 + (v And 255))
End Sub
Public Sub ST16(ByVal a As Long, ByVal v As Long)
    ST8 a, v
    ST8 a + 1, (v And &HFF00&) \ 256
End Sub
Public Sub ST16A(ByVal a As Long, ByVal v As Long)
    If (a And 1) <> 0 Then Trap
    ST16 a, v
End Sub
Public Sub ST32(ByVal a As Long, ByVal v As Long)
    If (a And 3) = 0 Then
        M(a \ 4) = v
    Else
        ST8 a, v
        ST8 a + 1, (v And &HFF00&) \ 256
        ST8 a + 2, (v And &HFF0000) \ 65536
        ST8 a + 3, ((v And &HFF000000) \ 16777216) And 255
    End If
End Sub
Public Sub ST64(ByVal a As Long, ByVal v As LongLong)
    ST32 a, W32(v)
    ST32 a + 4, CLng((v And (-4294967296^)) \ 4294967296^)
End Sub

Public Function Add64(ByVal a As LongLong, ByVal b As LongLong) As LongLong
    Dim lo As LongLong, hi As LongLong
    lo = (a And 4294967295^) + (b And 4294967295^)
    hi = ((a And (-4294967296^)) \ 4294967296^) + ((b And (-4294967296^)) \ 4294967296^) + (lo \ 4294967296^)
    Add64 = (CLngLng(W32(hi)) * 4294967296^) Or (lo And 4294967295^)
End Function
Public Function Sub64(ByVal a As LongLong, ByVal b As LongLong) As LongLong
    Sub64 = Add64(Add64(a, Not b), 1^)
End Function
Public Function Shl64(ByVal v As LongLong, ByVal k As Long) As LongLong
    Dim r As LongLong
    If k = 0 Then
        Shl64 = v
        Exit Function
    End If
    r = (v And (P2(63 - k) - 1^)) * P2(k)
    If (v And P2(63 - k)) <> 0 Then r = r Or P2(63)
    Shl64 = r
End Function
Public Function ShrU64(ByVal v As LongLong, ByVal k As Long) As LongLong
    If k = 0 Then
        ShrU64 = v
    ElseIf v >= 0 Then
        ShrU64 = v \ P2(k)
    Else
        ShrU64 = ((v And 9223372036854775807^) \ P2(k)) Or P2(63 - k)
    End If
End Function
Public Function ShrS64(ByVal v As LongLong, ByVal k As Long) As LongLong
    If k = 0 Then
        ShrS64 = v
    ElseIf k = 63 Then
        ShrS64 = CLngLng(v < 0)
    Else
        ShrS64 = (v And (Not (P2(k) - 1^))) \ P2(k)
    End If
End Function
Public Function Rotl64(ByVal v As LongLong, ByVal k As Long) As LongLong
    If k = 0 Then
        Rotl64 = v
    Else
        Rotl64 = Shl64(v, k) Or ShrU64(v, 64 - k)
    End If
End Function
Public Function Mul64(ByVal a As LongLong, ByVal b As LongLong) As LongLong
    Dim al As LongLong, ah As LongLong, bl As LongLong, bh As LongLong, x0 As LongLong, x1 As LongLong
    Dim lo As LongLong, cr As LongLong
    al = a And 4294967295^: bl = b And 4294967295^
    ah = (a And (-4294967296^)) \ 4294967296^: bh = (b And (-4294967296^)) \ 4294967296^
    x0 = al And 65535^: x1 = al \ 65536^
    lo = Add64(x0 * bl, Shl64(x1 * bl, 16))
    cr = ((ah * bl) And 4294967295^) + ((al * bh) And 4294967295^)
    Mul64 = Add64(lo, Shl64(cr And 4294967295^, 32))
End Function
Public Function DivU64(ByVal a As LongLong, ByVal b As LongLong) As LongLong
    Dim q As LongLong, r As LongLong
    If b < 0 Then
        If (a Xor P2(63)) >= (b Xor P2(63)) Then
            DivU64 = 1
        Else
            DivU64 = 0
        End If
        Exit Function
    End If
    If a >= 0 Then
        DivU64 = a \ b
        Exit Function
    End If
    q = Shl64(ShrU64(a, 1) \ b, 1)
    r = Sub64(a, Mul64(q, b))
    If (r Xor P2(63)) >= (b Xor P2(63)) Then q = q + 1
    DivU64 = q
End Function
Public Function RemU64(ByVal a As LongLong, ByVal b As LongLong) As LongLong
    RemU64 = Sub64(a, Mul64(DivU64(a, b), b))
End Function
Public Function Clz64(ByVal v As LongLong) As LongLong
    Dim n As LongLong
    If v = 0 Then
        Clz64 = 64
        Exit Function
    End If
    If v < 0 Then
        Clz64 = 0
        Exit Function
    End If
    n = 1
    Do While v < 4611686018427387904^
        v = v * 2: n = n + 1
    Loop
    Clz64 = n
End Function
Public Function Ctz64(ByVal v As LongLong) As LongLong
    Dim n As LongLong
    If v = 0 Then
        Ctz64 = 64
        Exit Function
    End If
    Do While (v And 1^) = 0
        v = ShrU64(v, 1): n = n + 1
    Loop
    Ctz64 = n
End Function
Public Function Popcnt64(ByVal v As LongLong) As LongLong
    Dim n As LongLong
    Do While v <> 0
        n = n + (v And 1^): v = ShrU64(v, 1)
    Loop
    Popcnt64 = n
End Function

Public Sub MemFill(ByVal d As Long, ByVal v As Long, ByVal n As Long)
    Dim w As Long, e As Long
    If n <= 0 Then Exit Sub
    v = v And 255
    e = d + n
    Do While d < e
        If (d And 3) = 0 Then
            If e - d >= 4 Then Exit Do
        End If
        ST8 d, v
        d = d + 1
    Loop
    If d < e Then
        w = v Or v * 256& Or v * 65536 Or BT(768 + v)
        Do While e - d >= 4
            M(d \ 4) = w
            d = d + 4
        Loop
        Do While d < e
            ST8 d, v
            d = d + 1
        Loop
    End If
End Sub
Public Sub MemCopy(ByVal d As Long, ByVal s As Long, ByVal n As Long)
    Dim e As Long
    If n <= 0 Or d = s Then Exit Sub
    If d < s Or d >= s + n Then
        e = s + n
        If ((d Xor s) And 3) = 0 Then
            Do While (s And 3) <> 0 And s < e
                ST8 d, LD8(s)
                d = d + 1: s = s + 1
            Loop
            Do While e - s >= 4
                M(d \ 4) = M(s \ 4)
                d = d + 4: s = s + 4
            Loop
        End If
        Do While s < e
            ST8 d, LD8(s)
            d = d + 1: s = s + 1
        Loop
    Else
        e = n
        Do While e > 0
            e = e - 1
            ST8 d + e, LD8(s + e)
        Loop
    End If
End Sub
Public Function MemGrow(ByVal n As Long) As Long
    If MemPages + n > MemCap Then
        MemGrow = -1
    Else
        MemGrow = MemPages
        MemPages = MemPages + n
    End If
End Function

Public Function TruncL(ByVal x As Double) As Long
    TruncL = CLng(Fix(x))
End Function
Public Function TruncUL(ByVal x As Double) As Long
    TruncUL = W32(CLngLng(Fix(x)))
End Function
Public Function TruncQ(ByVal x As Double) As LongLong
    TruncQ = CLngLng(Fix(x))
End Function
Public Function TruncUQ(ByVal x As Double) As LongLong
    TruncUQ = CLngLng(Fix(x))
End Function
Public Function CvtUQ(ByVal q As LongLong) As Double
    If q >= 0 Then
        CvtUQ = CDbl(q)
    Else
        CvtUQ = CDbl(q) + 18446744073709551616#
    End If
End Function
Public Function FMin(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then
        FMin = a
    Else
        FMin = b
    End If
End Function
Public Function FMax(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then
        FMax = a
    Else
        FMax = b
    End If
End Function
Public Function FCeil(ByVal a As Double) As Double
    FCeil = -Int(-a)
End Function
Public Function FFloor(ByVal a As Double) As Double
    FFloor = Int(a)
End Function
Public Function FNearest(ByVal a As Double) As Double
    FNearest = Int(a + 0.5)
End Function
Public Function FCopySign(ByVal a As Double, ByVal b As Double) As Double
    If b < 0 Then
        FCopySign = -Abs(a)
    Else
        FCopySign = Abs(a)
    End If
End Function
'''

RTF = r'''Option Explicit
Private Type TD
    d As Double
End Type
Private Type TQ
    q As LongLong
End Type
Private Type TS
    s As Single
End Type
Private Type TL
    l As Long
End Type
Public Function F64Bits(ByVal x As Double) As LongLong
    Dim a As TD, b As TQ
    a.d = x
    LSet b = a
    F64Bits = b.q
End Function
Public Function BitsF64(ByVal x As LongLong) As Double
    Dim a As TQ, b As TD
    a.q = x
    LSet b = a
    BitsF64 = b.d
End Function
Public Function F32Bits(ByVal x As Single) As Long
    Dim a As TS, b As TL
    a.s = x
    LSet b = a
    F32Bits = b.l
End Function
Public Function BitsF32(ByVal x As Long) As Single
    Dim a As TL, b As TS
    a.l = x
    LSet b = a
    BitsF32 = b.s
End Function
'''

WASI = r'''Option Explicit
Public WD() As Long, WDLen As Long
Public FPos(0 To 15) As Long, FUsed(0 To 15) As Long
Public OutB(0 To 65535) As Long, OutN As Long
Public ExitCode As Long
Public WN(0 To 8) As Long

Public Sub WasiInit()
    Dim i As Long
    WN(0) = 100: WN(1) = 111: WN(2) = 111: WN(3) = 109: WN(4) = 49: WN(5) = 46: WN(6) = 119: WN(7) = 97: WN(8) = 100
    For i = 0 To 15
        FPos(i) = 0: FUsed(i) = 0
    Next i
    OutN = 0: ExitCode = -1
End Sub

Private Function WB(ByVal p As Long) As Long
    WB = ((WD(p \ 4) And MK(p And 3)) \ DV(p And 3)) And 255
End Function

Private Sub WadCopy(ByVal d As Long, ByVal s As Long, ByVal n As Long)
    Dim e As Long
    e = s + n
    If ((d Xor s) And 3) = 0 Then
        Do While (s And 3) <> 0 And s < e
            ST8 d, WB(s)
            d = d + 1: s = s + 1
        Loop
        Do While e - s >= 4
            M(d \ 4) = WD(s \ 4)
            d = d + 4: s = s + 4
        Loop
    End If
    Do While s < e
        ST8 d, WB(s)
        d = d + 1: s = s + 1
    Loop
End Sub

Private Function IsWad(ByVal p As Long, ByVal n As Long) As Long
    Dim i As Long, c As Long
    IsWad = 0
    If n >= 2 Then
        If LD8(p) = 46 And LD8(p + 1) = 47 Then
            p = p + 2: n = n - 2
        End If
    End If
    If n <> 9 Then Exit Function
    For i = 0 To 8
        c = LD8(p + i)
        If c >= 65 And c <= 90 Then c = c + 32
        If c <> WN(i) Then Exit Function
    Next i
    IsWad = 1
End Function

Public Function imp_fd_prestat_get(ByVal fd As Long, ByVal buf As Long) As Long
    If fd = 3 Then
        ST32 buf, 0
        ST32 buf + 4, 1
        imp_fd_prestat_get = 0
    Else
        imp_fd_prestat_get = 8
    End If
End Function
Public Function imp_fd_prestat_dir_name(ByVal fd As Long, ByVal p As Long, ByVal n As Long) As Long
    If n > 0 Then ST8 p, 46
    imp_fd_prestat_dir_name = 0
End Function
Public Function imp_fd_fdstat_get(ByVal fd As Long, ByVal buf As Long) As Long
    Dim ft As Long
    If fd < 3 Then
        ft = 2
    ElseIf fd = 3 Then
        ft = 3
    Else
        ft = 4
    End If
    ST32 buf, ft
    ST32 buf + 4, 0
    ST32 buf + 8, -1
    ST32 buf + 12, -1
    ST32 buf + 16, -1
    ST32 buf + 20, -1
    imp_fd_fdstat_get = 0
End Function
Public Function imp_fd_fdstat_set_flags(ByVal fd As Long, ByVal fl As Long) As Long
    imp_fd_fdstat_set_flags = 0
End Function
Public Function imp_path_open(ByVal dfd As Long, ByVal dfl As Long, ByVal p As Long, ByVal n As Long, ByVal ofl As Long, ByVal rb As LongLong, ByVal ri As LongLong, ByVal ffl As Long, ByVal pfd As Long) As Long
    Dim i As Long
    imp_path_open = 44
    If (ofl And 9) <> 0 Then Exit Function
    If IsWad(p, n) = 0 Then Exit Function
    For i = 4 To 15
        If FUsed(i) = 0 Then
            FUsed(i) = 1: FPos(i) = 0
            ST32 pfd, i
            imp_path_open = 0
            Exit Function
        End If
    Next i
    imp_path_open = 33
End Function
Public Function imp_fd_close(ByVal fd As Long) As Long
    If fd >= 4 And fd <= 15 Then FUsed(fd) = 0
    imp_fd_close = 0
End Function
Public Function imp_fd_seek(ByVal fd As Long, ByVal ofs As LongLong, ByVal wh As Long, ByVal pr As Long) As Long
    Dim np As LongLong
    If fd < 4 Then
        imp_fd_seek = 70
        Exit Function
    End If
    If wh = 0 Then
        np = ofs
    ElseIf wh = 1 Then
        np = FPos(fd) + ofs
    Else
        np = WDLen + ofs
    End If
    If np < 0 Then
        imp_fd_seek = 28
        Exit Function
    End If
    FPos(fd) = CLng(np)
    ST32 pr, CLng(np)
    ST32 pr + 4, 0
    imp_fd_seek = 0
End Function
Public Function imp_fd_read(ByVal fd As Long, ByVal iovs As Long, ByVal cnt As Long, ByVal pn As Long) As Long
    Dim i As Long, p As Long, ln As Long, total As Long, pos As Long
    imp_fd_read = 0
    If fd < 4 Then
        ST32 pn, 0
        Exit Function
    End If
    pos = FPos(fd)
    For i = 0 To cnt - 1
        p = LD32(iovs + i * 8): ln = LD32(iovs + i * 8 + 4)
        If ln > WDLen - pos Then ln = WDLen - pos
        If ln > 0 Then
            WadCopy p, pos, ln
            pos = pos + ln: total = total + ln
        End If
    Next i
    FPos(fd) = pos
    ST32 pn, total
End Function
Public Function imp_fd_write(ByVal fd As Long, ByVal iovs As Long, ByVal cnt As Long, ByVal pn As Long) As Long
    Dim i As Long, p As Long, ln As Long, total As Long, k As Long
    For i = 0 To cnt - 1
        p = LD32(iovs + i * 8): ln = LD32(iovs + i * 8 + 4)
        For k = 0 To ln - 1
            OutB(OutN And 65535) = LD8(p + k)
            OutN = OutN + 1
        Next k
        total = total + ln
    Next i
    ST32 pn, total
    imp_fd_write = 0
End Function
Public Function imp_path_create_directory(ByVal fd As Long, ByVal p As Long, ByVal n As Long) As Long
    imp_path_create_directory = 0
End Function
Public Function imp_path_remove_directory(ByVal fd As Long, ByVal p As Long, ByVal n As Long) As Long
    imp_path_remove_directory = 44
End Function
Public Function imp_path_rename(ByVal fd As Long, ByVal p As Long, ByVal n As Long, ByVal fd2 As Long, ByVal p2 As Long, ByVal n2 As Long) As Long
    imp_path_rename = 44
End Function
Public Function imp_path_unlink_file(ByVal fd As Long, ByVal p As Long, ByVal n As Long) As Long
    imp_path_unlink_file = 44
End Function
Public Sub imp_proc_exit(ByVal code As Long)
    ExitCode = code
    Err.Raise vbObjectError + 78, , "proc_exit"
End Sub
'''

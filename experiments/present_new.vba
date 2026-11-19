Private Function ColName(ByVal c As Long) As String
    Dim s As String
    Do While c > 0
        s = Chr$(65 + ((c - 1) Mod 26)) & s
        c = (c - 1) \ 26
    Loop
    ColName = s
End Function

Private Sub InitDisplay()
    Dim i As Long
    For i = 0 To 319
        ColS(i) = ColName(GX + i + 1)
    Next i
    For i = 0 To 199
        RowS(i) = CStr(GY + i + 1)
    Next i
    For i = 0 To 255
        BHead(i) = -1: PrevPal(i) = &H7FFFFFFF
    Next i
    For i = 0 To 63999
        Shown(i) = -1
    Next i
    ReDim RX0(0 To 63999): ReDim RY0(0 To 63999): ReDim RX1(0 To 63999): ReDim RY1(0 To 63999): ReDim RC(0 To 63999)
    ReDim RP(0 To 63999): ReDim RNext(0 To 63999)
    ScanAll = True: Dirty = False: UnbStreak = 0: LastTol = TolCur
End Sub

' Paint one engine frame into the cells.
'  - rows whose pixels did not change since last frame are skipped outright
'  - changed pixels are grouped into runs of near-identical colour, runs that line up
'    in consecutive rows become rectangles, and each takes its first pixel's palette colour
'  - rectangles are bucketed by colour and sent to Excel ~20 at a time through one
'    multi-area Range call, which costs a fraction of one call per rectangle
'  - small frames are painted without Excel's full-window redraw; big ones with it
Private Sub Present(ByVal buf As Long, ByVal pal As Long)
    Dim i As Long, w As Long, y As Long, x As Long, x1 As Long, k As Long
    Dim p As Long, q As Long, c As Long, sh As Long, r0 As Long, g0 As Long, b0 As Long
    Dim need As Boolean, tl As Long, cur As Long, prv As Long, pix As Long, rect As Long, rowBase As Long, wb As Long, wr As Long, d As Long
    Dim rowSame As Boolean, a As String, piece As String, batch As Boolean
    Dim tp0 As Double, tp1 As Double, tp2 As Double
    tp0 = Timer
    tl = TolCur
    If tl < LastTol Then ScanAll = True
    LastTol = tl
    wb = pal \ 4
    For i = 0 To 255
        w = M(wb + i)
        If w <> PrevPal(i) Then
            PrevPal(i) = w: ScanAll = True
            PalB(i) = w And 255: PalG(i) = (w And &HFF00&) \ 256: PalR(i) = (w And &HFF0000) \ 65536
            PalC(i) = PalR(i) Or PalG(i) * 256& Or PalB(i) * 65536
        End If
    Next i
    NR = 0: cur = 0: RunN(0) = 0: RunN(1) = 0: NBUsed = 0
    wb = buf \ 4
    For y = 0 To PH - 1
        wr = wb + y * 80
        rowBase = y * 80
        rowSame = Not ScanAll
        If rowSame Then
            For i = 0 To 79
                If M(wr + i) <> PrevW(rowBase + i) Then
                    rowSame = False
                    Exit For
                End If
            Next i
        End If
        prv = cur: cur = 1 - cur: RunN(cur) = 0: pix = 0
        If Not rowSame Then
            For i = 0 To 79
                w = M(wr + i): PrevW(rowBase + i) = w
                RowP(i * 4) = w And 255: RowP(i * 4 + 1) = (w And &HFF00&) \ 256
                RowP(i * 4 + 2) = (w And &HFF0000) \ 65536: RowP(i * 4 + 3) = ((w And &HFF000000) \ 16777216) And 255
            Next i
            rowBase = y * PW
            x = 0
            Do While x < PW
                p = RowP(x): c = PalC(p): sh = Shown(rowBase + x)
                need = False
                If c <> sh Then
                    If sh < 0 Then
                        need = True
                    ElseIf Abs(PalR(p) - (sh And 255)) > tl Then
                        need = True
                    ElseIf Abs(PalG(p) - ((sh And &HFF00&) \ 256)) > tl Then
                        need = True
                    ElseIf Abs(PalB(p) - ((sh And &HFF0000) \ 65536)) > tl Then
                        need = True
                    End If
                End If
                If need Then
                    r0 = PalR(p): g0 = PalG(p): b0 = PalB(p)
                    x1 = x
                    Do While x1 < PW - 1
                        q = RowP(x1 + 1)
                        If q <> p Then
                            If Abs(PalR(q) - r0) > tl Then Exit Do
                            If Abs(PalG(q) - g0) > tl Then Exit Do
                            If Abs(PalB(q) - b0) > tl Then Exit Do
                        End If
                        x1 = x1 + 1
                    Loop
                    rect = -1
                    Do While pix < RunN(prv)
                        If RunA(prv, pix) >= x Then Exit Do
                        pix = pix + 1
                    Loop
                    If pix < RunN(prv) Then
                        If RunA(prv, pix) = x Then
                            If RunB(prv, pix) = x1 Then
                                d = RP(RunR(prv, pix))
                                If d = p Then
                                    rect = RunR(prv, pix)
                                ElseIf Abs(PalR(d) - r0) <= tl Then
                                    If Abs(PalG(d) - g0) <= tl Then
                                        If Abs(PalB(d) - b0) <= tl Then rect = RunR(prv, pix)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If rect >= 0 Then
                        RY1(rect) = y: c = PalC(RP(rect))
                    Else
                        rect = NR: NR = NR + 1
                        RX0(rect) = x: RX1(rect) = x1: RY0(rect) = y: RY1(rect) = y: RP(rect) = p
                        If BHead(p) < 0 Then
                            BUsed(NBUsed) = p: NBUsed = NBUsed + 1
                        End If
                        RNext(rect) = BHead(p): BHead(p) = rect
                    End If
                    k = RunN(cur): RunA(cur, k) = x: RunB(cur, k) = x1: RunR(cur, k) = rect: RunN(cur) = k + 1
                    For k = rowBase + x To rowBase + x1
                        Shown(k) = c
                    Next k
                    x = x1 + 1
                Else
                    x = x + 1
                End If
            Loop
        End If
    Next y
    ScanAll = False
    Calls = Calls + NR: TolSum = TolSum + tl
    If NR > 4500 Then
        If TolCur < 24 Then TolCur = TolCur + 4
    ElseIf NR < 2000 Then
        If TolCur > 8 Then TolCur = TolCur - 2
    End If
    tp1 = Timer: TScan = TScan + (tp1 - tp0)
    If NR = 0 Then
        NZero = NZero + 1
        If Dirty Then
            ' the picture has gone still: one full redraw guarantees the screen is exactly right
            Application.ScreenUpdating = False
            Application.ScreenUpdating = True
            Dirty = False: UnbStreak = 0: NSettle = NSettle + 1
        End If
        DoEvents
        TRedraw = TRedraw + (Timer - tp1)
        Exit Sub
    End If
    If NR < 200 Then NSmall = NSmall + 1
    batch = (NR >= BATCHMIN) Or (UnbStreak >= 20)
    On Error Resume Next
    If batch Then Application.ScreenUpdating = False
    For k = 0 To NBUsed - 1
        p = BUsed(k): c = PalC(p): rect = BHead(p): a = ""
        Do While rect >= 0
            If RX0(rect) = RX1(rect) And RY0(rect) = RY1(rect) Then
                piece = ColS(RX0(rect)) & RowS(RY0(rect))
            Else
                piece = ColS(RX0(rect)) & RowS(RY0(rect)) & ":" & ColS(RX1(rect)) & RowS(RY1(rect))
            End If
            If Len(a) = 0 Then
                a = piece
            ElseIf Len(a) + Len(piece) > 240 Then
                WS.Range(a).Interior.Color = c: OMCalls = OMCalls + 1
                a = piece
            Else
                a = a & "," & piece
            End If
            rect = RNext(rect)
        Loop
        If Len(a) > 0 Then
            WS.Range(a).Interior.Color = c: OMCalls = OMCalls + 1
        End If
        BHead(p) = -1
        If (k And 7) = 7 Then PollKeys
    Next k
    If Err.Number <> 0 Then
        DbgFillErr = DbgFillErr + 1
        If Len(DbgErrText) = 0 Then DbgErrText = Err.Number & " " & Err.Description
        Err.Clear
    End If
    tp2 = Timer: TFill = TFill + (tp2 - tp1)
    If batch Then
        Application.ScreenUpdating = True
        Batched = Batched + 1: CallsBatch = CallsBatch + NR
        Dirty = False: UnbStreak = 0
    Else
        Dirty = True: UnbStreak = UnbStreak + 1
        NUnb = NUnb + 1: RectsUnb = RectsUnb + NR: TUnb = TUnb + (tp2 - tp1)
    End If
    On Error GoTo 0
    DoEvents
    TRedraw = TRedraw + (Timer - tp2)
    If batch Then TBatch = TBatch + (Timer - tp1)
    PollKeys
End Sub


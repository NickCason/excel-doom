"""Telemetry module (VBA): accumulates per-frame statistics during play and writes them
as key=value text to a file in Excel's sandbox folder, at quit and periodically."""

TELE = r'''Option Explicit
' frame classes: 0 still, 1 direct (no full redraw), 2 full redraw (big frame), 3 full redraw (key held),
'                4 full redraw (every 30th small frame), 5 settle redraw
Private CN(0 To 5) As Long, CMs(0 To 5) As Double, CEng(0 To 5) As Double, CScan(0 To 5) As Double, CFill(0 To 5) As Double, CRed(0 To 5) As Double
Private CR(0 To 5) As Double, CR2(0 To 5) As Double, CP(0 To 5) As Double, CRP(0 To 5) As Double
Private FH(0 To 9) As Long
Private SH(0 To 4, 0 To 9) As Long, SN(0 To 4) As Long, SMs(0 To 4) As Double, SRects(0 To 4) As Double, SBat(0 To 4) As Long, SEng(0 To 4) As Double
Private RH(0 To 8) As Long, RHMs(0 To 8) As Double, RHPaint(0 To 8) As Double
Private TolH(0 To 12) As Long
Private TicH(0 To 8) As Long, ETn As Double, ETt As Double, ETtt As Double, ETe As Double, ETte As Double
Private MapN(0 To 99) As Long, MapMs(0 To 99) As Double, GsN(0 To 7) As Long, GsMs(0 To 7) As Double
Private MenuN As Long, MenuMs As Double, PauseN As Long
Private WaitMsSum As Double, WaitFrames As Long
Private SlowN As Long, Slow(0 To 63, 0 To 12) As Double
Private TL(0 To 1439, 0 To 10) As Double, TLMax As Long
Private LastX As Long, LastY As Long, LastA As Long, LastAmmo As Long, LastTic As Long, LastPS As Long, LastMap As Long, LastHealth As Long
Private Deaths As Long, LevelChanges As Long, Shots As Long, DamageTaken As Long
Private NFrames As Long, TotalMs As Double, MaxMs As Double
Public TeleStamp As String, TeleStartNow As String

Public Sub TeleInit()
    Erase CN, CMs, CEng, CScan, CFill, CRed, CR, CR2, CP, CRP, FH, SH, SN, SMs, SRects, SBat, SEng
    Erase RH, RHMs, RHPaint, TolH, TicH, MapN, MapMs, GsN, GsMs, Slow, TL
    ETn = 0: ETt = 0: ETtt = 0: ETe = 0: ETte = 0
    MenuN = 0: MenuMs = 0: PauseN = 0: WaitMsSum = 0: WaitFrames = 0: SlowN = 0: TLMax = -1
    LastX = 0: LastY = 0: LastA = 0: LastAmmo = -1: LastTic = 0: LastPS = 0: LastMap = -1: LastHealth = -1
    Deaths = 0: LevelChanges = 0: Shots = 0: DamageTaken = 0
    NFrames = 0: TotalMs = 0: MaxMs = 0
    TeleStamp = Format(Now, "yyyymmdd_hhnnss")
    TeleStartNow = Format(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Function FB(ByVal ms As Double) As Long
    If ms < 16 Then
        FB = 0
    ElseIf ms < 25 Then
        FB = 1
    ElseIf ms < 33 Then
        FB = 2
    ElseIf ms < 50 Then
        FB = 3
    ElseIf ms < 75 Then
        FB = 4
    ElseIf ms < 100 Then
        FB = 5
    ElseIf ms < 150 Then
        FB = 6
    ElseIf ms < 250 Then
        FB = 7
    ElseIf ms < 500 Then
        FB = 8
    Else
        FB = 9
    End If
End Function

Private Function RB(ByVal r As Long) As Long
    If r <= 0 Then
        RB = 0
    ElseIf r <= 50 Then
        RB = 1
    ElseIf r <= 200 Then
        RB = 2
    ElseIf r <= 800 Then
        RB = 3
    ElseIf r <= 2000 Then
        RB = 4
    ElseIf r <= 4000 Then
        RB = 5
    ElseIf r <= 8000 Then
        RB = 6
    ElseIf r <= 16000 Then
        RB = 7
    Else
        RB = 8
    End If
End Function

Public Sub TeleFrame(ByVal tSess As Double, ByVal ms As Double, ByVal eng As Double, ByVal sc As Double, ByVal fl As Double, ByVal rd As Double, ByVal rects As Long, ByVal cls As Long, ByVal tol As Long, ByVal wt As Double)
    Dim paint As Double, st As Long, tics As Long, b As Long, gs As Long, mp As Long, mn As Long, ps As Long, tb As Long
    Dim x As Long, y As Long, a As Long, ammo As Long, tic As Long, hp As Long, moved As Boolean, turned As Boolean, k As Long, full As Long
    If cls < 0 Or cls > 5 Then cls = 0
    paint = sc + fl + rd
    NFrames = NFrames + 1: TotalMs = TotalMs + ms
    If ms > MaxMs Then MaxMs = ms
    CN(cls) = CN(cls) + 1: CMs(cls) = CMs(cls) + ms: CEng(cls) = CEng(cls) + eng: CScan(cls) = CScan(cls) + sc
    CFill(cls) = CFill(cls) + fl: CRed(cls) = CRed(cls) + rd
    CR(cls) = CR(cls) + rects: CR2(cls) = CR2(cls) + CDbl(rects) * rects: CP(cls) = CP(cls) + paint: CRP(cls) = CRP(cls) + rects * paint
    b = FB(ms): FH(b) = FH(b) + 1
    If cls >= 2 Then full = 1
    gs = x_dg_stat(6): mp = x_dg_stat(7): mn = x_dg_stat(8): tic = x_dg_stat(5)
    x = x_dg_stat(3): y = x_dg_stat(4): a = x_dg_stat(2): ammo = x_dg_stat(0): hp = x_dg_stat(1): ps = x_dg_stat(16)
    moved = (x <> LastX Or y <> LastY): turned = (a <> LastA)
    If gs <> 0 Or mn <> 0 Then
        st = 4
    ElseIf moved And turned Then
        st = 3
    ElseIf moved Then
        st = 2
    ElseIf turned Then
        st = 1
    Else
        st = 0
    End If
    SN(st) = SN(st) + 1: SMs(st) = SMs(st) + ms: SRects(st) = SRects(st) + rects: SBat(st) = SBat(st) + full: SEng(st) = SEng(st) + eng
    SH(st, b) = SH(st, b) + 1
    k = RB(rects): RH(k) = RH(k) + 1: RHMs(k) = RHMs(k) + ms: RHPaint(k) = RHPaint(k) + paint
    k = (tol - 8) \ 2
    If k < 0 Then k = 0
    If k > 12 Then k = 12
    TolH(k) = TolH(k) + 1
    tics = tic - LastTic
    If tics < 0 Then tics = 0
    If tics > 8 Then tics = 8
    TicH(tics) = TicH(tics) + 1
    If st < 4 Then
        ETn = ETn + 1: ETt = ETt + tics: ETtt = ETtt + tics * tics: ETe = ETe + eng: ETte = ETte + tics * eng
    End If
    If wt > 0 Then
        WaitMsSum = WaitMsSum + wt: WaitFrames = WaitFrames + 1
    End If
    If mp < 0 Then mp = 0
    If mp > 99 Then mp = 99
    MapN(mp) = MapN(mp) + 1: MapMs(mp) = MapMs(mp) + ms
    If gs >= 0 And gs <= 7 Then
        GsN(gs) = GsN(gs) + 1: GsMs(gs) = GsMs(gs) + ms
    End If
    If mn <> 0 Then
        MenuN = MenuN + 1: MenuMs = MenuMs + ms
    End If
    If x_dg_stat(14) <> 0 Then PauseN = PauseN + 1
    If LastAmmo >= 0 And ammo >= 0 And ammo < LastAmmo Then Shots = Shots + (LastAmmo - ammo)
    If LastHealth >= 0 And hp >= 0 And hp < LastHealth Then DamageTaken = DamageTaken + (LastHealth - hp)
    If ps = 1 And LastPS <> 1 Then Deaths = Deaths + 1
    If LastMap >= 0 And mp <> LastMap Then LevelChanges = LevelChanges + 1
    If ms >= 200 Then
        k = SlowN And 63
        Slow(k, 0) = tSess: Slow(k, 1) = NFrames: Slow(k, 2) = ms: Slow(k, 3) = eng: Slow(k, 4) = sc: Slow(k, 5) = fl: Slow(k, 6) = rd
        Slow(k, 7) = rects: Slow(k, 8) = cls: Slow(k, 9) = tics: Slow(k, 10) = st: Slow(k, 11) = mp: Slow(k, 12) = tol
        SlowN = SlowN + 1
    End If
    tb = Int(tSess / 5)
    If tb < 0 Then tb = 0
    If tb > 1439 Then tb = 1439
    If tb > TLMax Then TLMax = tb
    TL(tb, 0) = TL(tb, 0) + 1: TL(tb, 1) = TL(tb, 1) + ms: TL(tb, 2) = TL(tb, 2) + full: TL(tb, 3) = TL(tb, 3) + rects
    TL(tb, 4) = TL(tb, 4) + eng: TL(tb, 5) = TL(tb, 5) + paint: TL(tb, 7) = mp: TL(tb, 9) = TL(tb, 9) + tics
    If st = 2 Or st = 3 Then TL(tb, 6) = TL(tb, 6) + 1
    If st = 1 Or st = 3 Then TL(tb, 10) = TL(tb, 10) + 1
    If ms > TL(tb, 8) Then TL(tb, 8) = ms
    LastX = x: LastY = y: LastA = a: LastAmmo = ammo: LastTic = tic: LastPS = ps: LastMap = mp: LastHealth = hp
End Sub

Private Function F1(ByVal v As Double) As String
    F1 = Format(v, "0.0")
End Function

Public Function TeleText(ByVal hostText As String) As String
    Dim s As String, i As Long, j As Long, t As String, nl As String, k As Long
    nl = vbLf
    s = "doom_telemetry=1" & nl & "stamp=" & TeleStamp & nl & "start=" & TeleStartNow & nl & "end=" & Format(Now, "yyyy-mm-dd hh:nn:ss") & nl
    s = s & "excel=" & Application.Version & nl & "os=" & Application.OperatingSystem & nl
    s = s & "frames=" & NFrames & nl & "total_s=" & F1(TotalMs / 1000) & nl & "max_frame_ms=" & F1(MaxMs) & nl
    If TotalMs > 0 Then s = s & "avg_fps=" & Format(NFrames / (TotalMs / 1000), "0.00") & nl
    s = s & "class_names=still,direct,full_big,full_keyheld,full_streak,settle" & nl
    t = "": For i = 0 To 5: t = t & CN(i) & ",": Next i: s = s & "class_n=" & t & nl
    t = "": For i = 0 To 5: t = t & F1(CMs(i)) & ",": Next i: s = s & "class_ms=" & t & nl
    t = "": For i = 0 To 5: t = t & F1(CEng(i)) & ",": Next i: s = s & "class_engine_ms=" & t & nl
    t = "": For i = 0 To 5: t = t & F1(CScan(i)) & ",": Next i: s = s & "class_scan_ms=" & t & nl
    t = "": For i = 0 To 5: t = t & F1(CFill(i)) & ",": Next i: s = s & "class_fill_ms=" & t & nl
    t = "": For i = 0 To 5: t = t & F1(CRed(i)) & ",": Next i: s = s & "class_redraw_ms=" & t & nl
    t = "": For i = 0 To 5: t = t & Format(CR(i), "0") & ",": Next i: s = s & "class_rects=" & t & nl
    t = "": For i = 0 To 5: t = t & Format(CR2(i), "0") & ",": Next i: s = s & "class_rects_sq=" & t & nl
    t = "": For i = 0 To 5: t = t & F1(CP(i)) & ",": Next i: s = s & "class_paint_ms=" & t & nl
    t = "": For i = 0 To 5: t = t & Format(CRP(i), "0") & ",": Next i: s = s & "class_rects_x_paint=" & t & nl
    s = s & "frame_ms_edges=16,25,33,50,75,100,150,250,500" & nl
    t = "": For i = 0 To 9: t = t & FH(i) & ",": Next i: s = s & "frame_hist=" & t & nl
    s = s & "state_names=idle,turning,moving,moving_turning,not_in_level" & nl
    t = "": For i = 0 To 4: t = t & SN(i) & ",": Next i: s = s & "state_n=" & t & nl
    t = "": For i = 0 To 4: t = t & F1(SMs(i)) & ",": Next i: s = s & "state_ms=" & t & nl
    t = "": For i = 0 To 4: t = t & F1(SEng(i)) & ",": Next i: s = s & "state_engine_ms=" & t & nl
    t = "": For i = 0 To 4: t = t & Format(SRects(i), "0") & ",": Next i: s = s & "state_rects=" & t & nl
    t = "": For i = 0 To 4: t = t & SBat(i) & ",": Next i: s = s & "state_full_redraws=" & t & nl
    For i = 0 To 4
        t = "": For j = 0 To 9: t = t & SH(i, j) & ",": Next j: s = s & "state_hist_" & i & "=" & t & nl
    Next i
    s = s & "rect_edges=0,50,200,800,2000,4000,8000,16000" & nl
    t = "": For i = 0 To 8: t = t & RH(i) & ",": Next i: s = s & "rect_hist=" & t & nl
    t = "": For i = 0 To 8: t = t & F1(RHMs(i)) & ",": Next i: s = s & "rect_frame_ms=" & t & nl
    t = "": For i = 0 To 8: t = t & F1(RHPaint(i)) & ",": Next i: s = s & "rect_paint_ms=" & t & nl
    t = "": For i = 0 To 12: t = t & TolH(i) & ",": Next i: s = s & "tolerance_hist_from_8_step_2=" & t & nl
    t = "": For i = 0 To 8: t = t & TicH(i) & ",": Next i: s = s & "tics_per_frame_hist=" & t & nl
    s = s & "engine_fit_n_sumtics_sumtics2_sumeng_sumticseng=" & Format(ETn, "0") & "," & Format(ETt, "0") & "," & Format(ETtt, "0") & "," & F1(ETe) & "," & F1(ETte) & nl
    s = s & "tic_wait_ms=" & F1(WaitMsSum) & nl & "tic_wait_frames=" & WaitFrames & nl
    t = ""
    For i = 0 To 99
        If MapN(i) > 0 Then t = t & i & ":" & MapN(i) & ":" & F1(MapMs(i)) & ";"
    Next i
    s = s & "maps_n_ms=" & t & nl
    t = "": For i = 0 To 7: t = t & GsN(i) & ":" & F1(GsMs(i)) & ";": Next i: s = s & "gamestate_n_ms=" & t & nl
    s = s & "menu_frames=" & MenuN & nl & "menu_ms=" & F1(MenuMs) & nl & "paused_frames=" & PauseN & nl
    s = s & "game_end=map " & x_dg_stat(7) & " kills " & x_dg_stat(9) & " items " & x_dg_stat(10) & " secrets " & x_dg_stat(11) & " leveltime " & x_dg_stat(12)
    s = s & " health " & x_dg_stat(1) & " armor " & x_dg_stat(13) & " bullets " & x_dg_stat(0) & " otherammo " & x_dg_stat(18) & " weapon " & x_dg_stat(17) & " gametic " & x_dg_stat(5) & nl
    s = s & "deaths=" & Deaths & nl & "level_changes=" & LevelChanges & nl & "bullets_fired=" & Shots & nl & "damage_taken=" & DamageTaken & nl
    s = s & "slow_frame_fields=t_s,frame,ms,engine,scan,fill,redraw,rects,class,tics,state,map,tol" & nl
    k = SlowN
    If k > 64 Then k = 64
    s = s & "slow_frames_total=" & SlowN & nl
    For i = 0 To k - 1
        t = ""
        For j = 0 To 12
            t = t & F1(Slow(i, j)) & ","
        Next j
        s = s & "slow=" & t & nl
    Next i
    s = s & "timeline_fields=bucket5s,frames,ms,full_redraws,rects,engine_ms,paint_ms,moving_frames,map,max_ms,tics,turning_frames" & nl
    For i = 0 To TLMax
        If TL(i, 0) > 0 Then
            t = i & ","
            For j = 0 To 10
                t = t & Format(TL(i, j), "0") & ","
            Next j
            s = s & "tl=" & t & nl
        End If
    Next i
    TeleText = s & hostText
End Function

Public Function TelePath() As String
#If Mac Then
    TelePath = Environ("HOME") & "/doom_telemetry_" & TeleStamp & ".txt"
#Else
    TelePath = Environ("USERPROFILE") & "\doom_telemetry_" & TeleStamp & ".txt"
#End If
End Function

Public Function TeleSave(ByVal txt As String) As String
    Dim f As Integer, p As String
    On Error GoTo Bad
    p = TelePath()
    f = FreeFile
    Open p For Output As #f
    Print #f, txt
    Close #f
    TeleSave = p
    Exit Function
Bad:
    TeleSave = "SAVE FAILED " & Err.Number & " " & Err.Description
End Function
'''

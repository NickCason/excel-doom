"""Excel-side host module: display, keyboard, mouse, main loop.
Uses the Excel object model plus a handful of built-in OS calls for input state."""

HOST = r'''Option Explicit
Private Type PSN
    hi As Long
    lo As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type
#If Mac Then
Private Declare PtrSafe Function CGEventSourceKeyState Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" (ByVal stateID As Long, ByVal vk As Integer) As Long
Private Declare PtrSafe Function CGEventSourceButtonState Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" (ByVal stateID As Long, ByVal button As Long) As Long
Private Declare PtrSafe Function CGEventCreate Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" (ByVal src As LongPtr) As LongPtr
Private Declare PtrSafe Function CGEventGetLocation Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" (ByVal ev As LongPtr) As Double
Private Declare PtrSafe Sub CFRelease Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" (ByVal p As LongPtr)
Private Declare PtrSafe Function CGWarpMouseCursorPosition Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" (ByVal x As Double, ByVal y As Double) As Long
Private Declare PtrSafe Function CGAssociateMouseAndMouseCursorPosition Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" (ByVal connected As Long) As Long
Private Declare PtrSafe Function CGMainDisplayID Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" () As Long
Private Declare PtrSafe Function CGDisplayPixelsWide Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" (ByVal d As Long) As LongPtr
Private Declare PtrSafe Function CGDisplayPixelsHigh Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" (ByVal d As Long) As LongPtr
Private Declare PtrSafe Function mach_absolute_time Lib "/usr/lib/libSystem.B.dylib" () As LongLong
Private Declare PtrSafe Function GetFrontProcess Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" (ByRef p As PSN) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" (ByRef p As PSN) As Long
Private Declare PtrSafe Function SameProcess Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" (ByRef p1 As PSN, ByRef p2 As PSN, ByRef res As Byte) As Long
#Else
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vk As Long) As Integer
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef c As LongLong) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef c As LongLong) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef p As POINTAPI) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal n As Long) As Long
Private Declare PtrSafe Function PlaySoundMem Lib "winmm" Alias "PlaySoundA" (ByRef snd As Byte, ByVal hmod As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function PlaySoundStop Lib "winmm" Alias "PlaySoundA" (ByVal snd As LongPtr, ByVal hmod As LongPtr, ByVal flags As Long) As Long
#End If

Private Const PW As Long = 320
Private Const PH As Long = 200
Private Const GX As Long = 1
Private Const GY As Long = 1
Private Const FW As Long = 322
Private Const FH As Long = 202
Private Const WADC As Long = 701
Private Const WADR As Long = 301
Private Const BATCHMIN As Long = 1800
Private Const BATCHWIN As Long = 100
Private Const CHKROW As Long = 10
Private Const STILLN As Long = 10
Private Const BANX As Long = @@BANX@@
Private Const BANY As Long = @@BANY@@
Private Const BANW As Long = @@BANW@@
Private Const BANH As Long = @@BANH@@
Private Const MSENS As Double = 6#

Private Const BUILD_ID As String = "@@BUILD@@"
Private HiOK As Long, HiScale As Double, HiBase As LongLong, THiStart As Double
Private PScan As Double, PFill As Double, PRedraw As Double, PRects As Long, PClass As Long, TolNow As Long
Private WaitOpen As Boolean
Private WaitFirst As Double, WaitLast As Double, SleepCalls As Long, LastPollT As Double, PollSkips As Long, LastFullT As Double
Private KName(0 To 47) As String, KHits(0 To 47) As Long, KHeld(0 To 47) As Long, KDefer(0 To 47) As Long
Private DeferRel As Long, QMax As Long, MouseFrames As Long, MouseAbs As Double, Clicks As Long, MToggles As Long, LastBtn As Long
Private PalChanges As Long, ScanAllN As Long, RowsScanned As Long, RowsSkipped As Long
Public Playing As Boolean
Private TeleToFile As Boolean, LastEndT As Double
Public EvtCount As Long
Private FrontOK As Long
Private NK As Long
Private KCode(0 To 47) As Long, KDoom(0 To 47) As Long, KDown(0 To 47) As Long, KDeliv(0 To 47) As Long
Private KQ(0 To 255) As Long, KQHead As Long, KQTail As Long
Private MouseOn As Boolean, MArmed As Boolean, MouseDX As Double, MBtn As Long, MSticky As Long
Private WarpX As Double, WarpY As Double, ClickCount As Long
Private DbgPolls As Long, DbgNotFront As Long, DbgDowns As Long, DbgUps As Long, DbgPops As Long, DbgClk As Long, DbgBtn As Long
Private Shown(0 To 63999) As Long
Private PalC(0 To 255) As Long, PalR(0 To 255) As Long, PalG(0 To 255) As Long, PalB(0 To 255) As Long
Private RowP(0 To 319) As Long
Private RunA(0 To 1, 0 To 319) As Long, RunB(0 To 1, 0 To 319) As Long, RunR(0 To 1, 0 To 319) As Long, RunN(0 To 1) As Long
Private RX0() As Long, RY0() As Long, RX1() As Long, RY1() As Long, RC() As Long, NR As Long
Private ColS(0 To 319) As String, RowS(0 To 199) As String
Private RP() As Long, RNext() As Long, BHead(0 To 255) As Long, BUsed(0 To 255) As Long, NBUsed As Long
Private PrevW(0 To 15999) As Long, PrevPal(0 To 255) As Long, ScanAll As Boolean, Dirty As Boolean, UnbStreak As Long, LastTol As Long
Private OMCalls As Long, TUnb As Double, RectsUnb As Long, NUnb As Long, NSettle As Long, StillFrames As Long
Private HeldNow As Boolean, HoldOff As Long, NHeldBatch As Long
Private WS As Worksheet
Private Running As Boolean
Private Frames As Long, TStart As Double, TPaint As Double, Calls As Long, Batched As Long, TolCur As Long, TolSum As Long
Private TBatch As Double, CallsBatch As Long
Private TScan As Double, TFill As Double, TRedraw As Double, NZero As Long, NSmall As Long
Public PresentMode As Long
Private LogT(0 To 4095) As Double, LogA(0 To 4095) As Long, LogM(0 To 4095) As Long, LogN As Long
Private DbgFillErr As Long, DbgErrText As String, DbgReadBad As Long, DbgReadOk As Long
Private Virt As Boolean, VClock As Long
Private Sums(0 To 199) As LongLong, NSums As Long
Private WavA() As Byte, WavB() As Byte, WavCapA As Long, WavCapB As Long, WavUseB As Boolean
Private SndVol(0 To 255) As Byte, SndVolNow As Long
Private SndAsked As Long, SndPlayed As Long, SndFailed As Long, SndSamples As Long, TSnd As Double
Private SndMute As Boolean, Cheats As Long, RunCount As Long
Private LastBuf As Long, LastPal As Long, SelR As Long, SelC As Long
Private TolLo As Long, TolHi As Long, TolSettle As Long
Private ViewA As Long, ViewX As Long, ViewY As Long, ViewStill As Long, PalSeen As Long

' ---------------------------------------------------------------- telemetry clock
' Timer only ticks 64 times a second on Mac; this counts in nanoseconds. Used for measurement only.
Public Function Hi() As Double
#If Mac Then
#Else
    Dim c As LongLong
#End If
    If HiOK <= 0 Then
        Hi = Timer
        Exit Function
    End If
#If Mac Then
    Hi = CDbl(mach_absolute_time() - HiBase) * HiScale
#Else
    QueryPerformanceCounter c
    Hi = CDbl(c - HiBase) * HiScale
#End If
End Function

Private Sub InitHi()
    Dim t As Double, a As LongLong, b As LongLong, f As LongLong
    HiOK = 0
    On Error GoTo NoHi
#If Mac Then
    t = Timer
    Do While Timer = t
    Loop
    a = mach_absolute_time(): t = Timer
    Do While Timer - t < 0.25 And Timer >= t
    Loop
    b = mach_absolute_time()
    If b > a Then HiScale = (Timer - t) / CDbl(b - a)
    HiBase = a
#Else
    QueryPerformanceFrequency f
    QueryPerformanceCounter a
    HiBase = a
    If f > 0 Then HiScale = 1# / CDbl(f)
#End If
    If HiScale > 0.0000000001 And HiScale < 0.00001 Then HiOK = 1
    Exit Sub
NoHi:
    HiOK = -1
End Sub

' ------------------------------------------------------------------ imports
Public Sub imp_draw_frame(ByVal buf As Long, ByVal pal As Long)
    Dim t As Double
    WaitOpen = False
    If Virt Then
        FrameSum buf, pal
    Else
        t = Timer
        Present buf, pal
        TPaint = TPaint + (Timer - t)
        LastBuf = buf: LastPal = pal
    End If
    Frames = Frames + 1
End Sub

Public Function imp_ticks_ms() As Long
    Dim d As Double
    If Virt Then
        imp_ticks_ms = VClock And &H7FFFFFFF
    Else
        d = Hi()
        If d - LastPollT > 0.004 Or d < LastPollT Then
            LastPollT = d
            PollKeys
        Else
            PollSkips = PollSkips + 1
        End If
        d = Timer - TStart
        If d < 0 Then d = d + 86400#
        imp_ticks_ms = CLng(d * 1000#)
    End If
End Function

Public Sub imp_sleep_ms(ByVal ms As Long)
    If Virt Then
        VClock = VClock + ms
    Else
        SleepCalls = SleepCalls + 1
        WaitLast = Hi()
        If WaitFirst < 0 Then WaitFirst = WaitLast
    End If
End Sub

Public Function imp_poll_key() As Long
    Dim v As Long
    If KQHead = KQTail Then
        imp_poll_key = 0
    Else
        v = KQ(KQHead)
        KQHead = (KQHead + 1) And 255
        If (v And 256) <> 0 Then KDeliv(v \ 65536) = 1
        DbgPops = DbgPops + 1
        imp_poll_key = v And 511
    End If
End Function

Public Function imp_mouse_buttons() As Long
    imp_mouse_buttons = MBtn Or MSticky
    MSticky = 0
End Function

Public Function imp_mouse_dx() As Long
    Dim d As Double
    d = MouseDX * MSENS
    MouseDX = 0
    If d > 8000 Then d = 8000
    If d < -8000 Then d = -8000
    imp_mouse_dx = CLng(d)
End Function

' ---------------------------------------------------------------------- sound
' The engine hands over the 8-bit samples of one effect. Windows plays a wave file straight from
' memory, so the samples are copied behind a 44-byte header and started without waiting. One
' effect at a time: starting the next one ends the last. Two buffers take turns, because the one
' that is playing must not be touched. Returns 1 if the effect was started.
Private Sub WavFill(ByVal p As Long, ByVal n As Long, ByVal rate As Long, ByRef b() As Byte)
    Dim i As Long, a As Long, w As Long, j As Long, e As Long
    b(0) = 82: b(1) = 73: b(2) = 70: b(3) = 70                              ' RIFF
    w = n + 36
    b(4) = w And 255: b(5) = (w \ 256) And 255: b(6) = (w \ 65536) And 255: b(7) = (w \ 16777216) And 255
    b(8) = 87: b(9) = 65: b(10) = 86: b(11) = 69                            ' WAVE
    b(12) = 102: b(13) = 109: b(14) = 116: b(15) = 32                       ' fmt
    b(16) = 16: b(17) = 0: b(18) = 0: b(19) = 0
    b(20) = 1: b(21) = 0: b(22) = 1: b(23) = 0                              ' PCM, one channel
    b(24) = rate And 255: b(25) = (rate \ 256) And 255: b(26) = (rate \ 65536) And 255: b(27) = 0
    b(28) = b(24): b(29) = b(25): b(30) = b(26): b(31) = 0                  ' bytes a second
    b(32) = 1: b(33) = 0: b(34) = 8: b(35) = 0                              ' 8 bits
    b(36) = 100: b(37) = 97: b(38) = 116: b(39) = 97                        ' data
    b(40) = n And 255: b(41) = (n \ 256) And 255: b(42) = (n \ 65536) And 255: b(43) = (n \ 16777216) And 255
    a = p: e = p + n: j = 44
    Do While a < e And (a And 3) <> 0                                       ' up to the next whole word
        w = M(a \ 4)
        Select Case a And 3
            Case 1: b(j) = SndVol((w And &HFF00&) \ 256)
            Case 2: b(j) = SndVol((w And &HFF0000) \ 65536)
            Case 3: b(j) = SndVol(((w And &HFF000000) \ 16777216) And 255)
        End Select
        a = a + 1: j = j + 1
    Loop
    i = a \ 4
    Do While a + 3 < e
        w = M(i)
        b(j) = SndVol(w And 255): b(j + 1) = SndVol((w And &HFF00&) \ 256)
        b(j + 2) = SndVol((w And &HFF0000) \ 65536): b(j + 3) = SndVol(((w And &HFF000000) \ 16777216) And 255)
        i = i + 1: a = a + 4: j = j + 4
    Loop
    Do While a < e
        w = M(a \ 4)
        Select Case a And 3
            Case 0: b(j) = SndVol(w And 255)
            Case 1: b(j) = SndVol((w And &HFF00&) \ 256)
            Case 2: b(j) = SndVol((w And &HFF0000) \ 65536)
        End Select
        a = a + 1: j = j + 1
    Loop
End Sub

Public Function imp_play_sound(ByVal p As Long, ByVal n As Long, ByVal rate As Long, ByVal vol As Long, ByVal sep As Long) As Long
    Dim i As Long, t As Double, ok As Long
    imp_play_sound = 0
#If Mac Then
#Else
    If Virt Or SndMute Then Exit Function
    If n <= 0 Or n > 4000000 Or vol <= 0 Then Exit Function
    SndAsked = SndAsked + 1
    On Error GoTo NoSound
    t = Hi()
    If vol > 127 Then vol = 127
    If vol <> SndVolNow Then
        ' samples are unsigned around 128; the engine's volume runs to 127
        For i = 0 To 255
            SndVol(i) = 128 + ((i - 128) * vol) \ 127
        Next i
        SndVolNow = vol
    End If
    WavUseB = Not WavUseB
    If WavUseB Then
        If WavCapB < n + 44 Then
            WavCapB = n + 44 + 4096: ReDim WavB(0 To WavCapB - 1)
        End If
        WavFill p, n, rate, WavB
        ok = PlaySoundMem(WavB(0), 0, 7)                                    ' SND_ASYNC, SND_NODEFAULT, SND_MEMORY
    Else
        If WavCapA < n + 44 Then
            WavCapA = n + 44 + 4096: ReDim WavA(0 To WavCapA - 1)
        End If
        WavFill p, n, rate, WavA
        ok = PlaySoundMem(WavA(0), 0, 7)
    End If
    If ok <> 0 Then
        SndPlayed = SndPlayed + 1: SndSamples = SndSamples + n: imp_play_sound = 1
    Else
        SndFailed = SndFailed + 1
    End If
    TSnd = TSnd + (Hi() - t)
    Exit Function
NoSound:
    SndFailed = SndFailed + 1
#End If
End Function

' Nothing may still be playing from the buffers when they go away with the workbook.
Private Sub SoundOff()
#If Mac Then
#Else
    On Error Resume Next
    PlaySoundStop 0, 0, 0
    On Error GoTo 0
#End If
End Sub

' ------------------------------------------------------------ keyboard, mouse
Private Sub AddKey(ByVal macCode As Long, ByVal winCode As Long, ByVal doomKey As Long, ByVal nm As String)
#If Mac Then
    KCode(NK) = macCode
#Else
    KCode(NK) = winCode
#End If
    KDoom(NK) = doomKey: KDown(NK) = 0: KDeliv(NK) = 0: KName(NK) = nm: KHits(NK) = 0: KHeld(NK) = 0: KDefer(NK) = 0: NK = NK + 1
End Sub

Private Sub InitKeys()
    NK = 0: KQHead = 0: KQTail = 0
    DeferRel = 0: QMax = 0: MouseFrames = 0: MouseAbs = 0: Clicks = 0: MToggles = 0: LastBtn = 0
    AddKey 126, &H26, &HAD, "up": AddKey 125, &H28, &HAF, "down": AddKey 123, &H25, &HAC, "left": AddKey 124, &H27, &HAE, "right"
    AddKey 13, &H57, &HAD, "w": AddKey 1, &H53, &HAF, "s": AddKey 0, &H41, &HA0, "a": AddKey 2, &H44, &HA1, "d"
    AddKey 59, &H11, &HA3, "ctrl": AddKey 49, &H20, &HA2, "space": AddKey 14, &H45, &HA2, "e"
    AddKey 56, &H10, &HB6, "shift"
#If Mac Then
    AddKey 58, &H12, &HB8, "option"
#Else
    ' no Alt on Windows: it puts the ribbon into KeyTips mode, and the next game keys then drive the ribbon
#End If
    AddKey 36, &HD, 13, "enter": AddKey 53, &H1B, 27, "esc": AddKey 48, &H9, 9, "tab"
    AddKey 16, &H59, 121, "y": AddKey 45, &H4E, 110, "n"
    AddKey 18, &H31, 49, "1": AddKey 19, &H32, 50, "2": AddKey 20, &H33, 51, "3": AddKey 21, &H34, 52, "4"
    AddKey 23, &H35, 53, "5": AddKey 22, &H36, 54, "6": AddKey 26, &H37, 55, "7"
    AddKey 12, &H51, 0, "q"
    AddKey 46, &H4D, -1, "m"
    ' not game keys: the host acts on these itself (see PollKeys)
    AddKey 9, &H56, -2, "v"
    AddKey 28, &H38, -3, "8": AddKey 25, &H39, -4, "9": AddKey 29, &H30, -5, "0"
    AddKey 35, &H50, &HFF, "p"
End Sub

' The engine has the original cheat codes, but they cannot be typed: the letters are not game keys
' here and Q quits. So the host types a code for the player, as key presses nobody made.
Private Sub TypeCheat(ByVal code As String)
    Dim i As Long, c As Long
    For i = 1 To Len(code)
        c = Asc(Mid$(code, i, 1))
        KQ(KQTail) = 47 * 65536 + 256 + c: KQTail = (KQTail + 1) And 255
        KQ(KQTail) = 47 * 65536 + c: KQTail = (KQTail + 1) And 255
    Next i
End Sub

Private Function KeyIsDown(ByVal code As Long) As Long
#If Mac Then
    If (CGEventSourceKeyState(0, CInt(code)) And 255) <> 0 Then KeyIsDown = 1 Else KeyIsDown = 0
#Else
    If (GetAsyncKeyState(code) And &H8000) <> 0 Then KeyIsDown = 1 Else KeyIsDown = 0
#End If
End Function

Private Function LeftButtonDown() As Long
#If Mac Then
    If (CGEventSourceButtonState(0, 0) And 255) <> 0 Then LeftButtonDown = 1 Else LeftButtonDown = 0
#Else
    If (GetAsyncKeyState(1) And &H8000) <> 0 Then LeftButtonDown = 1 Else LeftButtonDown = 0
#End If
End Function

Private Function ExcelIsFront() As Boolean
    Dim a As PSN, b As PSN, same As Byte
    ExcelIsFront = True
    If FrontOK < 0 Then Exit Function
    On Error GoTo NoApi
#If Mac Then
    If GetFrontProcess(a) <> 0 Then GoTo NoApi
    If GetCurrentProcess(b) <> 0 Then GoTo NoApi
    If SameProcess(a, b, same) <> 0 Then GoTo NoApi
    ExcelIsFront = (same <> 0)
#Else
    ExcelIsFront = (GetForegroundWindow() = Application.Hwnd)
#End If
    FrontOK = 1
    Exit Function
NoApi:
    FrontOK = -1
    ExcelIsFront = True
End Function

' Called many times per frame. A key release is only passed on once the engine has
' taken the matching press, so a quick tap is never lost between two frames.
Private Sub PollKeys()
    Dim i As Long, d As Long, front As Boolean
    front = ExcelIsFront()
    DbgPolls = DbgPolls + 1
    If Not front Then DbgNotFront = DbgNotFront + 1
    HeldNow = False
    For i = 0 To NK - 1
        If front Then d = KeyIsDown(KCode(i)) Else d = 0
        If d = 1 Then
            ' Ctrl, Shift and Option never reach Excel as keystrokes; every other key does
            If KDoom(i) <> &HA3 And KDoom(i) <> &HB6 And KDoom(i) <> &HB8 Then HeldNow = True
        End If
        If d = 1 Then
            If KDown(i) = 0 Then
                KDown(i) = 1: KDeliv(i) = 0: KHits(i) = KHits(i) + 1: KDefer(i) = 0
                If KDoom(i) = 0 Then
                    Running = False
                ElseIf KDoom(i) = -1 Then
                    MouseOn = Not MouseOn: MArmed = False: MToggles = MToggles + 1
                ElseIf KDoom(i) = -2 Then
                    SndMute = Not SndMute
                    If SndMute Then SoundOff
                ElseIf KDoom(i) = -3 Then
                    TypeCheat "iddqd": Cheats = Cheats + 1
                ElseIf KDoom(i) = -4 Then
                    TypeCheat "idkfa": Cheats = Cheats + 1
                ElseIf KDoom(i) = -5 Then
                    TypeCheat "idspispopd": Cheats = Cheats + 1
                Else
                    KQ(KQTail) = i * 65536 + 256 + KDoom(i)
                    KQTail = (KQTail + 1) And 255
                    DbgDowns = DbgDowns + 1
                End If
            End If
        ElseIf KDown(i) = 1 Then
            If KDoom(i) <= 0 Then
                KDown(i) = 0
            ElseIf KDeliv(i) = 1 Then
                KDown(i) = 0
                KQ(KQTail) = i * 65536 + KDoom(i)
                KQTail = (KQTail + 1) And 255
                DbgUps = DbgUps + 1
                If KDefer(i) = 1 Then DeferRel = DeferRel + 1
            Else
                KDefer(i) = 1
            End If
        End If
    Next i
    i = (KQTail - KQHead) And 255
    If i > QMax Then QMax = i
    If front And MouseOn Then
        MBtn = LeftButtonDown()
        MSticky = MSticky Or MBtn
        DbgBtn = DbgBtn + MBtn
        If MBtn = 1 And LastBtn = 0 Then Clicks = Clicks + 1
        LastBtn = MBtn
    Else
        MBtn = 0
    End If
End Sub

Private Sub InitMouse()
    Dim sw As Double, sh As Double
#If Mac Then
    MouseOn = True
#Else
    MouseOn = False
#End If
    MArmed = False: MouseDX = 0: MBtn = 0: MSticky = 0
#If Mac Then
    sw = CDbl(CGDisplayPixelsWide(CGMainDisplayID())): sh = CDbl(CGDisplayPixelsHigh(CGMainDisplayID()))
#Else
    sw = GetSystemMetrics(0): sh = GetSystemMetrics(1)
#End If
    ' Park the cursor on the empty stretch of the menu bar. A click on the grid makes Excel run its own
    ' drag-tracking loop, which freezes the game until the button is released; up here Excel never sees it.
    WarpX = Int(sw * 0.54): WarpY = 4
End Sub

' Once per frame: how far the mouse moved sideways since last frame, then put the
' cursor back so it can never run off the screen.
Private Sub MouseSample()
    Dim x As Double
#If Mac Then
    Dim ev As LongPtr
#Else
    Dim p As POINTAPI
#End If
    If Not MouseOn Then
        MArmed = False
        Exit Sub
    End If
    If Not ExcelIsFront() Then
        MArmed = False
        Exit Sub
    End If
#If Mac Then
    ev = CGEventCreate(0)
    x = CGEventGetLocation(ev)
    CFRelease ev
    If MArmed Then
        MouseDX = MouseDX + (x - WarpX)
        If x <> WarpX Then
            MouseFrames = MouseFrames + 1: MouseAbs = MouseAbs + Abs(x - WarpX)
        End If
    End If
    CGWarpMouseCursorPosition WarpX, WarpY
    CGAssociateMouseAndMouseCursorPosition 1
#Else
    GetCursorPos p
    x = p.x
    If MArmed Then MouseDX = MouseDX + (x - WarpX)
    SetCursorPos CLng(WarpX), CLng(WarpY)
#End If
    MArmed = True
End Sub

Private Sub HookKeys(ByVal hook As Boolean)
    Dim ks As Variant, i As Long
#If Mac Then
#Else
    Dim ms As Variant, bs As Variant, j As Long
#End If
    ks = Array("{UP}", "{DOWN}", "{LEFT}", "{RIGHT}", "+{UP}", "+{DOWN}", "+{LEFT}", "+{RIGHT}", " ", "~", "{ESC}", "{TAB}", _
               "w", "a", "s", "d", "e", "f", "y", "n", "q", "m", "+w", "+a", "+s", "+d", "+e", "+f", "1", "2", "3", "4", "5", "6", "7", _
               "v", "p", "8", "9", "0", "+v", "+p")
    On Error Resume Next
    For i = LBound(ks) To UBound(ks)
        If hook Then
            Application.OnKey CStr(ks(i)), ""
        Else
            Application.OnKey CStr(ks(i))
        End If
    Next i
#If Mac Then
#Else
    ' Ctrl is fire, and on Windows nearly every Ctrl combination with a game key is an Excel
    ' shortcut: Ctrl+W (fire while walking) closes the workbook, Ctrl+1 opens a dialog.
    ms = Array("^", "^+")
    bs = Array("{UP}", "{DOWN}", "{LEFT}", "{RIGHT}", " ", "~", "{TAB}", "w", "a", "s", "d", "e", "f", "y", "n", "q", "m", _
               "1", "2", "3", "4", "5", "6", "7", "v", "p", "8", "9", "0")
    For j = LBound(ms) To UBound(ms)
        For i = LBound(bs) To UBound(bs)
            If hook Then
                Application.OnKey CStr(ms(j)) & CStr(bs(i)), ""
            Else
                Application.OnKey CStr(ms(j)) & CStr(bs(i))
            End If
        Next i
    Next j
#End If
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ display
Private Sub FrameSum(ByVal buf As Long, ByVal pal As Long)
    Dim i As Long, c As LongLong, wb As Long
    wb = buf \ 4
    For i = 0 To 15999
        c = ((c * 31) And 4294967295^) Xor (CLngLng(M(wb + i)) And 4294967295^)
    Next i
    wb = pal \ 4
    For i = 0 To 255
        c = ((c * 31) And 4294967295^) Xor (CLngLng(M(wb + i)) And 4294967295^)
    Next i
    If NSums <= 199 Then
        Sums(NSums) = c: NSums = NSums + 1
    End If
End Sub

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
    ScanAll = True: Dirty = False: UnbStreak = 0: LastTol = TolCur: StillFrames = 0
End Sub

' Merging near colours saves paint calls, and it is worth it while the whole view moves. It also
' leaves streaks: a cell that is near enough to the right colour is not repainted, so pieces of
' old frames stay. So when the player has neither moved nor turned for STILLN frames the tolerance
' drops to TolSettle, which makes that frame rescan and repaint everything exactly, and it stays
' there while only parts of the picture change (doors, monsters, animated floors).
Private Sub ViewCheck()
    Dim a As Long, x As Long, y As Long
    a = x_dg_stat(2): x = x_dg_stat(3): y = x_dg_stat(4)
    If a <> ViewA Or x <> ViewX Or y <> ViewY Then
        ViewA = a: ViewX = x: ViewY = y: ViewStill = 0
        If TolCur < TolLo Then TolCur = TolLo
    Else
        ViewStill = ViewStill + 1
        If ViewStill = STILLN Then TolCur = TolSettle
    End If
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
    Dim BatchAt As Long
    BatchAt = BATCHMIN
#If Mac Then
#Else
    ' On Windows a fill made with screen updating on costs about 0.4 ms per rectangle and a full
    ' redraw about 50 ms, so the full redraw wins from about a hundred rectangles up.
    BatchAt = BATCHWIN
#End If
    If (PresentMode And 4) <> 0 Then BatchAt = 1000000
    If (PresentMode And 128) <> 0 Then BatchAt = 1
    tp0 = Hi()
    If TolSettle < TolLo Then ViewCheck
    tl = TolCur
    If tl < LastTol Then ScanAll = True
    LastTol = tl
    wb = pal \ 4
    d = 0
    For i = 0 To 255
        w = M(wb + i)
        If w <> PrevPal(i) Then
            PrevPal(i) = w: ScanAll = True: d = 1
            PalB(i) = w And 255: PalG(i) = (w And &HFF00&) \ 256: PalR(i) = (w And &HFF0000) \ 65536
            PalC(i) = PalR(i) Or PalG(i) * 256& Or PalB(i) * 65536
        End If
    Next i
    If d = 1 Then PalChanges = PalChanges + 1
    If ScanAll Then ScanAllN = ScanAllN + 1
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
        If rowSame Then RowsSkipped = RowsSkipped + 1 Else RowsScanned = RowsScanned + 1
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
    If NR > 4500 And (ViewStill < STILLN Or NR > 9000) Then
        If TolCur < TolLo Then
            TolCur = TolLo
        ElseIf TolCur < TolHi Then
            TolCur = TolCur + 4
        End If
        If TolCur > TolHi Then TolCur = TolHi
    ElseIf ViewStill >= STILLN Then
        ' the view is at rest: exact again as soon as a frame is not huge (a palette fade can be)
        If NR < 6000 Then TolCur = TolSettle
    ElseIf NR < 2000 Then
        If TolCur > TolLo Then TolCur = TolCur - 2
    End If
    tp1 = Hi(): TScan = TScan + (tp1 - tp0)
    PScan = PScan + (tp1 - tp0) * 1000: PRects = PRects + NR: TolNow = tl
    If NR = 0 Then
        NZero = NZero + 1
        StillFrames = StillFrames + 1
        If Dirty And StillFrames >= 4 And (tp1 - LastFullT > 2 Or tp1 < LastFullT) Then
            ' the picture has been still for a moment: one full redraw guarantees the screen is exactly right
            Application.ScreenUpdating = False
            Application.ScreenUpdating = True
            Dirty = False: UnbStreak = 0: NSettle = NSettle + 1: LastFullT = tp1
            PClass = 5
        End If
        DoEvents
        tp2 = Hi(): TRedraw = TRedraw + (tp2 - tp1): PRedraw = PRedraw + (tp2 - tp1) * 1000
        Exit Sub
    End If
    StillFrames = 0
    If NR < 200 Then NSmall = NSmall + 1
    ' While a key that Excel swallows is held (or was a moment ago), Excel does not repaint small
    ' changes, so those frames must take the full redraw. Ctrl, Shift, Option and the mouse are exempt.
    If HeldNow Then
        HoldOff = 3
    ElseIf HoldOff > 0 Then
        HoldOff = HoldOff - 1
    End If
    batch = (NR >= BatchAt) Or (UnbStreak >= 30) Or (HoldOff > 0)
    If NR >= BatchAt Then
        PClass = 2
    ElseIf HoldOff > 0 Then
        PClass = 3
    ElseIf UnbStreak >= 30 Then
        PClass = 4
    Else
        PClass = 1
    End If
    If HoldOff > 0 And NR < BatchAt Then NHeldBatch = NHeldBatch + 1
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
        If (k And 7) = 7 Then
            If (PresentMode And 1) = 0 Then PollKeys
        End If
    Next k
    If Err.Number <> 0 Then
        DbgFillErr = DbgFillErr + 1
        If Len(DbgErrText) = 0 Then DbgErrText = Err.Number & " " & Err.Description
        Err.Clear
    End If
    tp2 = Hi(): TFill = TFill + (tp2 - tp1): PFill = PFill + (tp2 - tp1) * 1000
    If batch Then
        Application.ScreenUpdating = True
        Batched = Batched + 1: CallsBatch = CallsBatch + NR
        Dirty = False: UnbStreak = 0: LastFullT = tp2
    Else
        Dirty = True: UnbStreak = UnbStreak + 1
        NUnb = NUnb + 1: RectsUnb = RectsUnb + NR: TUnb = TUnb + (tp2 - tp1)
    End If
    On Error GoTo 0
    DoEvents
    tp0 = Hi(): TRedraw = TRedraw + (tp0 - tp2): PRedraw = PRedraw + (tp0 - tp2) * 1000
    If batch Then TBatch = TBatch + (tp0 - tp1)
    If (PresentMode And 1) = 0 Then PollKeys
End Sub

Private Sub Layout()
    Dim zs As Variant, i As Long, best As Long, cell As Double, fit As Double, r As Range, k As Long, w As Window
    Set w = ThisWorkbook.Windows(1)
    Application.ScreenUpdating = False
    w.DisplayGridlines = False
    w.DisplayHeadings = False
    On Error Resume Next
    ' game cells are square: rows as tall as the columns are wide; side panels are 136 cells wide
    WS.Range(WS.Cells(2, 1), WS.Cells(PH + 1, 1)).RowHeight = 3
    cell = WS.Columns(2).Width
    WS.Rows(1).RowHeight = 12 * cell: WS.Rows(PH + 2).RowHeight = 12 * cell
    For k = 1 To 3
        WS.Columns(1).ColumnWidth = WS.Columns(1).ColumnWidth * (136 * cell) / WS.Columns(1).Width
        WS.Columns(FW).ColumnWidth = WS.Columns(FW).ColumnWidth * (136 * cell) / WS.Columns(FW).Width
    Next k
    ' biggest whole-pixel zoom at which the entire frame fits the window
    Set r = WS.Range(WS.Cells(1, 1), WS.Cells(FH, FW))
    fit = w.UsableWidth / r.Width
    If w.UsableHeight / r.Height < fit Then fit = w.UsableHeight / r.Height
    fit = fit * 100 * 0.985
    zs = Array(50, 75, 100, 125, 150, 175, 200, 250, 300)
    best = 50
    For i = LBound(zs) To UBound(zs)
        If zs(i) <= fit Then best = zs(i)
    Next i
    w.Zoom = best
    On Error GoTo 0
    Application.Goto WS.Cells(1, 1), True
    WS.Cells(1, 1).Select
    Application.ScreenUpdating = True
    DoEvents
End Sub

' ---------------------------------------------------------------- game data
Private Sub LoadWad()
    Dim v As Variant, r As Long, c As Long, nrows As Long, k As Long
    WDLen = CLng(WS.Cells(WADR, WADC - 2).Value)
    nrows = CLng(WS.Cells(WADR, WADC - 1).Value)
    v = WS.Range(WS.Cells(WADR, WADC), WS.Cells(WADR + nrows - 1, WADC + 15)).Value
    ReDim WD(0 To nrows * 16)
    For r = 1 To nrows
        For c = 1 To 16
            WD(k) = CLng(v(r, c)): k = k + 1
        Next c
    Next r
End Sub

Private Function OutTail() As String
    Dim i As Long, s As String, a As Long
    a = OutN - 400
    If a < 0 Then a = 0
    For i = a To OutN - 1
        s = s & Chr$(OutB(i And 65535))
    Next i
    OutTail = s
End Function

' ------------------------------------------------------------- around a game
' The workbook opens on the game sheet: the title screen in the cells with the click-to-play sign
' over it, and under the frame two lines about allowing macros. Those lines only matter to
' someone whose macros are off, so the first macro that runs takes them away.
Private Sub GameOn()
    On Error Resume Next
    WS.Activate
    If Len(WS.Cells(FH + 1, 2).Value) > 0 Then WS.Range(WS.Cells(FH + 1, 2), WS.Cells(FH + 2, 2)).Value = ""
    On Error GoTo 0
End Sub

' Runs when the workbook is opened with macros allowed (ThisWorkbook, Workbook_Open): fit the
' frame to the window, as a game would on starting.
Public Sub Splash()
    On Error Resume Next
    Set WS = ThisWorkbook.Worksheets(1)
    GameOn
    Layout
    ThisWorkbook.Saved = True
    On Error GoTo 0
End Sub

' The click-to-play sign, one hex string of palette indices per row (written by mkdoom.py).
Private Sub BannerRows(ByRef r() As String)
    ReDim r(0 To BANH - 1)
    '@@BANNERROWS@@
End Sub

' When a game ends the sign goes back over the picture. It is not a picture object, because one
' more picture on this sheet halves the frame rate on Windows even while it is hidden: it is
' written into the engine's last frame, and the normal painter puts it in the cells.
Private Sub GameOff()
    Dim r() As String, x As Long, y As Long, a As Long
    On Error Resume Next
    If LastBuf <> 0 Then
        BannerRows r
        For y = 0 To BANH - 1
            a = LastBuf + (BANY + y) * PW + BANX
            For x = 0 To BANW - 1
                ST8 a + x, CLng("&H" & Mid$(r(y), x * 2 + 1, 2))
            Next x
        Next y
        TolCur = 4
        Present LastBuf, LastPal
    End If
    On Error GoTo 0
End Sub

' A mouse click on the picture while no game is running starts one (the sheet's SelectionChange
' event calls this). Arrow keys move the selection as well, but only ever to the next cell, so
' a selection that lands in the picture from further away is taken to be a click.
Public Sub CellClick(ByVal r As Long, ByVal c As Long)
    Dim jump As Boolean
    jump = Abs(r - SelR) + Abs(c - SelC) > 1
    SelR = r: SelC = c
    If Playing Or Not jump Then Exit Sub
    If r < GY + 1 Or r > GY + PH Or c < GX + 1 Or c > GX + PW Then Exit Sub
    PlayDoom
End Sub

' Tick the checklist's last task and note how the game went.
Private Sub LogRun()
    Dim cs As Worksheet, kills As Long, items As Long, secrets As Long, secs As Long, map As Long, note As String
    On Error Resume Next
    Set cs = ThisWorkbook.Worksheets("Checklist")
    kills = x_dg_stat(9): items = x_dg_stat(10): secrets = x_dg_stat(11): secs = x_dg_stat(12) \ 35: map = x_dg_stat(7)
    If kills < 0 Then kills = 0
    If items < 0 Then items = 0
    If secrets < 0 Then secrets = 0
    RunCount = RunCount + 1
    note = "Run " & RunCount & ": " & kills & IIf(kills = 1, " demon", " demons") & " dispatched, " & items & IIf(items = 1, " item, ", " items, ") _
         & secrets & IIf(secrets = 1, " secret, ", " secrets, ") & (secs \ 60) & ":" & Format(secs Mod 60, "00") & " on E" & (map \ 10) & "M" & (map Mod 10) & "."
    If x_dg_stat(16) = 1 Then note = note & " Did not make it."
    If Cheats > 0 Then note = note & " Cheated."
    cs.Cells(CHKROW, 4).Value = note
    cs.Shapes("Tick").TextFrame.Characters.Text = ChrW(9745)
    ' the cell colours change with every frame; nobody wants to be asked about saving them
    ThisWorkbook.Saved = True
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------- main loop
Private Function RunDoom(ByVal virtMode As Boolean, ByVal nticks As Long) As String
    Dim t As Double, n As Long, msg As String, i As Long, tLoad As Double, tBoot As Double, tRun As Double, s As String
    Dim tA As Double, tB As Double, tPrev As Double, t16 As Double, frameMs As Double, waitMs As Double, engMs As Double, telefile As String
    Dim sbx As String
    On Error GoTo Fail
    Set WS = ThisWorkbook.Worksheets(1)
    Virt = virtMode: VClock = 0: Frames = 0: Calls = 0: TPaint = 0: NSums = 0: Batched = 0: TolSum = 0: EvtCount = 0
    ' How far apart two colours may be and still be painted as one. It rises while frames are
    ' busy and falls while they are quiet; cells within it of the right colour are left alone.
    ' TolLo..TolHi is the range while things move; TolSettle is what a quiet picture is brought to
    ' (equal to TolLo: never). Windows: a lower ceiling, because 24 turns dark areas into flat
    ' patches, and an exact picture at rest.
#If Mac Then
    TolLo = 8: TolHi = 24: TolSettle = 8
#Else
    TolLo = @@TOLLO@@: TolHi = @@TOLHI@@: TolSettle = @@TOLSETTLE@@
#End If
    TolCur = TolLo: ViewStill = 0: ViewA = -2
    HeldNow = False: HoldOff = 0: NHeldBatch = 0
    TScan = 0: TFill = 0: TRedraw = 0: NZero = 0: NSmall = 0: OMCalls = 0: TUnb = 0: RectsUnb = 0: NUnb = 0: NSettle = 0
    TBatch = 0: CallsBatch = 0: DbgFillErr = 0: DbgErrText = "": DbgReadBad = 0: DbgReadOk = 0
    LogN = 0
    DbgPolls = 0: DbgNotFront = 0: DbgDowns = 0: DbgUps = 0: DbgPops = 0: DbgClk = 0: DbgBtn = 0
    SndAsked = 0: SndPlayed = 0: SndFailed = 0: SndSamples = 0: TSnd = 0: SndVolNow = -1
    Cheats = 0: LastBuf = 0
    If (PresentMode And 32) = 0 Then Application.EnableCancelKey = xlDisabled
    If Not Virt Then
        GameOn
        Layout
    End If
    TeleInit
    InitHi
    PalChanges = 0: ScanAllN = 0: RowsScanned = 0: RowsSkipped = 0: SleepCalls = 0: LastPollT = 0: PollSkips = 0: LastFullT = 0
    Application.StatusBar = "DOOM: loading game data..."
    DoEvents
    t = Timer
    LoadWad
    RtInit 512
    ProgInit
    WasiInit
    TouchAll
    tLoad = Timer - t
    InitKeys
    InitMouse
    If (PresentMode And 64) <> 0 Then MouseOn = False
    InitDisplay
    If Not Virt Then
        If (PresentMode And 2) = 0 Then HookKeys True
        Playing = True
    End If
    Application.StatusBar = "DOOM: starting..."
    TStart = Timer
    t = Timer
    x__initialize
    x_dg_start 1
    tBoot = Timer - t
    Running = True
    t = Timer
    tPrev = Hi(): THiStart = tPrev: t16 = tPrev
    Do While Running
        If Virt Then
            If n >= nticks Then Exit Do
            VClock = VClock + 29
        Else
            PollKeys
            MouseSample
        End If
        PScan = 0: PFill = 0: PRedraw = 0: PRects = 0: PClass = 0: WaitFirst = -1: WaitOpen = True
        tA = Hi()
        x_dg_tick
        n = n + 1
        If Not Virt Then
            tB = Hi()
            frameMs = (tB - tPrev) * 1000: tPrev = tB
            waitMs = 0
            If WaitFirst >= 0 Then waitMs = (WaitLast - WaitFirst) * 1000
            engMs = (tB - tA) * 1000 - PScan - PFill - PRedraw - waitMs
            For i = 0 To NK - 1
                If KDown(i) = 1 Then KHeld(i) = KHeld(i) + 1
            Next i
            TeleFrame tB - THiStart, frameMs, engMs, PScan, PFill, PRedraw, PRects, PClass, TolNow, waitMs
            LogT(LogN And 4095) = Timer: LogA(LogN And 4095) = x_dg_stat(2): LogM(LogN And 4095) = x_dg_stat(0): LogN = LogN + 1
            ' picture diagnostics: per frame, rectangles painted and the tolerance they were painted at
            If (PresentMode And 256) <> 0 Then LogA((LogN - 1) And 4095) = PRects: LogM((LogN - 1) And 4095) = TolNow * 10000 + IIf(ViewStill > 99, 99, ViewStill) * 100 + (PalChanges - PalSeen): PalSeen = PalChanges
            If MouseOn And (PresentMode And 16) = 0 Then
                If ActiveCell.Row <> 1 Then WS.Cells(1, 1).Select
            End If
            If (n And 15) = 0 And (PresentMode And 8) = 0 Then
                sbx = ""
#If Mac Then
#Else
                sbx = "    V: sound " & IIf(SndMute, "off", "on")
#End If
                Application.StatusBar = "DOOM   " & Format(16 / (tB - t16 + 0.000001), "0") & " fps    Ctrl: fire    M: mouse " & IIf(MouseOn, "on", "off") & sbx & "    8 9 0: cheats    P: pause    Q: quit"
                t16 = tB
            End If
            If TeleToFile And (n Mod 3000) = 0 Then TeleSave TeleText(HostTele())
        End If
    Loop
    tRun = Timer - t
    msg = "ok"
Cleanup:
    On Error Resume Next
    SoundOff
    LastEndT = Timer
    ' still "playing" while the sign is painted, so a stray click cannot start a second game inside this one
    If Not Virt And (PresentMode And 256) = 0 Then GameOff
    Playing = False
    Application.ScreenUpdating = True
    HookKeys False
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
    If Not Virt And TeleToFile Then telefile = TeleSave(TeleText(HostTele() & "exit=" & msg & vbLf))
    If Not Virt Then LogRun
    LastEndT = Timer
    On Error GoTo 0
    s = msg & "|load=" & Format(tLoad, "0.00") & "|boot=" & Format(tBoot, "0.00") & "|ticks=" & n & "|run=" & Format(tRun, "0.00") & "|frames=" & Frames & "|paint=" & Format(TPaint, "0.00") & "|calls=" & Calls
    s = s & "|polls=" & DbgPolls & "|notfront=" & DbgNotFront & "|downs=" & DbgDowns & "|ups=" & DbgUps & "|pops=" & DbgPops & "|clk=" & DbgClk & "|btn=" & DbgBtn
    s = s & "|ammo=" & x_dg_stat(0) & "|angle=" & x_dg_stat(2) & "|pos=" & x_dg_stat(3) & "," & x_dg_stat(4) & "|tbatch=" & Format(TBatch, "0.00") & "|cbatch=" & CallsBatch
    s = s & "|fillerr=" & DbgFillErr & "|errtext=" & DbgErrText & "|readok=" & DbgReadOk & "|readbad=" & DbgReadBad
    s = s & "|tscan=" & Format(TScan, "0.00") & "|tfill=" & Format(TFill, "0.00") & "|tredraw=" & Format(TRedraw, "0.00") & "|nzero=" & NZero & "|nsmall=" & NSmall
    s = s & "|omcalls=" & OMCalls & "|nunb=" & NUnb & "|rectsunb=" & RectsUnb & "|tunb=" & Format(TUnb, "0.00") & "|nsettle=" & NSettle
    s = s & "|nheldbatch=" & NHeldBatch
    s = s & "|telefile=" & telefile
    s = s & "|batched=" & Batched & "|tolsum=" & TolSum & "|evt=" & EvtCount & "|front=" & FrontOK & "|warp=" & Format(WarpX, "0") & "," & Format(WarpY, "0")
    s = s & "|sndasked=" & SndAsked & "|sndplayed=" & SndPlayed & "|sndfailed=" & SndFailed & "|sndsamples=" & SndSamples & "|tsnd=" & Format(TSnd, "0.000")
    If Not Virt Then
        s = s & "|log="
        i = LogN - 400
        If i < 0 Then i = 0
        Do While i < LogN
            s = s & Format(LogT(i And 4095), "0.00") & ":" & LogA(i And 4095) & ":" & LogM(i And 4095) & ";"
            i = i + 1
        Loop
    End If
    If Virt Then
        s = s & "|sums="
        For i = 0 To NSums - 1
            s = s & Sums(i) & ","
        Next i
    End If
    RunDoom = s
    Exit Function
Fail:
    If Err.Description = "proc_exit" Then
        msg = "exit " & ExitCode
    Else
        msg = "ERROR " & Err.Number & " " & Err.Description & " | out: " & OutTail()
    End If
    Resume Cleanup
End Function

Public Function SizeProbe() As String
    Dim s As String, z As Variant, w As Window
    On Error Resume Next
    Set WS = ThisWorkbook.Worksheets(1)
    Set w = ThisWorkbook.Windows(1)
    s = "colA.Width=" & WS.Columns(1).Width & " colA.ColumnWidth=" & WS.Columns(1).ColumnWidth & " colB.Width=" & WS.Columns(2).Width & " colB.ColumnWidth=" & WS.Columns(2).ColumnWidth
    s = s & " row1.Height=" & WS.Rows(1).Height & " row2.Height=" & WS.Rows(2).Height
    s = s & " frame.Width=" & WS.Range(WS.Cells(1, 1), WS.Cells(FH, FW)).Width & " frame.Height=" & WS.Range(WS.Cells(1, 1), WS.Cells(FH, FW)).Height
    s = s & " usableW=" & w.UsableWidth & " usableH=" & w.UsableHeight & " winW=" & w.Width & " winH=" & w.Height & " appUsableW=" & Application.UsableWidth & " appUsableH=" & Application.UsableHeight
    For Each z In Array(25, 38, 50, 63, 75, 100)
        w.Zoom = z
        Application.Goto WS.Cells(1, 1), True
        DoEvents
        s = s & " | zoom " & z & ": visible " & w.VisibleRange.Rows.Count & " rows x " & w.VisibleRange.Columns.Count & " cols"
    Next z
    SizeProbe = s
End Function

Private Function Floor1(ByVal n As Long) As Double
    Dim i As Long, t As Double
    t = Timer
    For i = 1 To n
        Application.ScreenUpdating = False
        WS.Cells(GY + 50, GX + 50 + (i And 7)).Interior.Color = RGB(255, i * 9 And 255, 0)
        Application.ScreenUpdating = True
        DoEvents
    Next i
    Floor1 = (Timer - t) / n * 1000
End Function

Public Function Audit() As String
    Dim s As String, i As Long, k As Long, t As Double, w As Window, n As Long, a As String, sh As Object
    On Error GoTo Fail
    Set WS = ThisWorkbook.Worksheets(1)
    Set w = ThisWorkbook.Windows(1)
    GameOn
    Layout
    s = "zoom=" & w.Zoom
    s = s & "|redraw_base=" & Format(Floor1(12), "0")
    For Each sh In WS.Shapes
        sh.Visible = False
    Next sh
    s = s & "|redraw_no_pictures=" & Format(Floor1(12), "0")
    WS.Range(WS.Cells(102, 1), WS.Cells(201, 1)).EntireRow.Hidden = True
    s = s & "|redraw_half_rows_no_pictures=" & Format(Floor1(12), "0")
    WS.Range(WS.Cells(102, 1), WS.Cells(201, 1)).EntireRow.Hidden = False
    For Each sh In WS.Shapes
        sh.Visible = True
    Next sh
    w.Zoom = 75
    s = s & "|redraw_zoom75=" & Format(Floor1(12), "0")
    w.Zoom = 50
    s = s & "|redraw_zoom50=" & Format(Floor1(12), "0")
    w.Zoom = 200
    Application.Goto WS.Cells(60, 100), True
    s = s & "|redraw_zoom200_partial=" & Format(Floor1(12), "0")
    Layout
    WS.Range(WS.Cells(GY + 1, GX + 1), WS.Cells(GY + 200, GX + 320)).Interior.Color = RGB(40, 40, 40)
    s = s & "|redraw_uniform_picture=" & Format(Floor1(12), "0")
    Application.ScreenUpdating = False
    For i = 0 To 79
        WS.Range(WS.Cells(GY + 1, GX + 1 + i * 4), WS.Cells(GY + 200, GX + 2 + i * 4)).Interior.Color = RGB(200, 40, 40)
    Next i
    Application.ScreenUpdating = True
    s = s & "|redraw_80_wide_stripes=" & Format(Floor1(12), "0")
    Application.ScreenUpdating = False
    For i = 0 To 199 Step 2
        WS.Range(WS.Cells(GY + 1 + i, GX + 1), WS.Cells(GY + 1 + i, GX + 320)).Interior.Color = RGB(40, 200, 40)
    Next i
    Application.ScreenUpdating = True
    s = s & "|redraw_stripes_x_rows_16k_runs=" & Format(Floor1(12), "0")
    WS.Range(WS.Cells(GY + 1, GX + 1), WS.Cells(GY + 200, GX + 320)).Interior.Color = RGB(40, 40, 40)
    WS.Range(WS.Cells(2, 1), WS.Cells(201, 1)).EntireRow.Hidden = True
    s = s & "|redraw_all_game_rows_hidden=" & Format(Floor1(12), "0")
    WS.Range(WS.Cells(2, 1), WS.Cells(201, 1)).EntireRow.Hidden = False
    WS.Range(WS.Cells(1, 2), WS.Cells(1, 161)).EntireColumn.Hidden = True
    s = s & "|redraw_uniform_half_columns=" & Format(Floor1(12), "0")
    WS.Range(WS.Cells(1, 2), WS.Cells(1, 161)).EntireColumn.Hidden = False
    s = s & "|U_base=" & Format(Floor1(20), "0")
    For Each sh In WS.Shapes
        sh.Visible = False
    Next sh
    s = s & "|U_no_pictures=" & Format(Floor1(20), "0")
    WS.Range(WS.Cells(102, 1), WS.Cells(201, 1)).EntireRow.Hidden = True
    s = s & "|U_no_pictures_half_rows=" & Format(Floor1(20), "0")
    WS.Range(WS.Cells(2, 1), WS.Cells(201, 1)).EntireRow.Hidden = True
    s = s & "|U_no_pictures_no_rows=" & Format(Floor1(20), "0")
    WS.Range(WS.Cells(2, 1), WS.Cells(201, 1)).EntireRow.Hidden = False
    For Each sh In WS.Shapes
        sh.Visible = True
    Next sh
    s = s & "|U_base_again=" & Format(Floor1(20), "0")
    GoTo SkipFills
    WS.Range(WS.Cells(1, 323), WS.Cells(1, 690)).EntireColumn.Hidden = True
    s = s & "|redraw_uniform_right_columns_hidden=" & Format(Floor1(12), "0")
    WS.Range(WS.Cells(1, 323), WS.Cells(1, 690)).EntireColumn.Hidden = False
    GoTo SkipFills
    ' cost of fills: 6000 two-cell runs, one call each
    Application.ScreenUpdating = False
    t = Timer
    For k = 1 To 3
        For i = 0 To 5999
            WS.Range(WS.Cells(GY + 1 + (i \ 100), GX + 1 + (i Mod 100) * 3), WS.Cells(GY + 1 + (i \ 100), GX + 2 + (i Mod 100) * 3)).Interior.Color = (i * 37 + k * 5000) And &HFFFFFF
        Next i
    Next k
    s = s & "|fill_single_us=" & Format((Timer - t) / 18000 * 1000000#, "0.0")
    ' same number of runs, 20 per call through one multi-area address, one colour per call
    t = Timer
    For k = 1 To 3
        For i = 0 To 299
            a = ""
            For n = 0 To 19
                a = a & "," & Cells(GY + 1 + i \ 5, GX + 1 + ((i Mod 5) * 20 + n) * 3).Address(False, False) & ":" & Cells(GY + 1 + i \ 5, GX + 2 + ((i Mod 5) * 20 + n) * 3).Address(False, False)
            Next n
            WS.Range(Mid$(a, 2)).Interior.Color = (i * 977 + k * 5000) And &HFFFFFF
        Next i
    Next k
    s = s & "|fill_multi20_us_per_run=" & Format((Timer - t) / 18000 * 1000000#, "0.0")
    ' address building alone
    t = Timer
    For k = 1 To 3
        For i = 0 To 299
            a = ""
            For n = 0 To 19
                a = a & "," & Cells(GY + 1 + i \ 5, GX + 1 + ((i Mod 5) * 20 + n) * 3).Address(False, False) & ":" & Cells(GY + 1 + i \ 5, GX + 2 + ((i Mod 5) * 20 + n) * 3).Address(False, False)
            Next n
        Next i
    Next k
    s = s & "|address_build_us_per_run=" & Format((Timer - t) / 18000 * 1000000#, "0.0")
    ' same colour for many runs in ONE call vs many calls
    t = Timer
    For k = 1 To 3
        For i = 0 To 5999
            WS.Range(WS.Cells(GY + 1 + (i \ 100), GX + 1 + (i Mod 100) * 3), WS.Cells(GY + 1 + (i \ 100), GX + 2 + (i Mod 100) * 3)).Interior.Color = k * 70000
        Next i
    Next k
    s = s & "|fill_single_samecolour_us=" & Format((Timer - t) / 18000 * 1000000#, "0.0")
    t = Timer
    For k = 1 To 30
        WS.Range(WS.Cells(GY + 1, GX + 1), WS.Cells(GY + 200, GX + 320)).Interior.Color = k * 50000
    Next k
    s = s & "|fill_whole_picture_ms=" & Format((Timer - t) / 30 * 1000, "0.0")
SkipFills:
    Application.ScreenUpdating = True
    DoEvents
    Audit = s
    Exit Function
Fail:
    Application.ScreenUpdating = True
    Audit = s & "|ERROR " & Err.Number & " " & Err.Description
End Function

Private Function HostTele() As String
    Dim s As String, i As Long, t As String, nl As String
    nl = vbLf
    On Error Resume Next
    s = "build=" & BUILD_ID & nl
    s = s & "zoom=" & ThisWorkbook.Windows(1).Zoom & nl & "usable=" & ThisWorkbook.Windows(1).UsableWidth & "x" & ThisWorkbook.Windows(1).UsableHeight & nl
    s = s & "screen=" & Format(WarpX * 2, "0") & "x" & Format(WarpY * 2, "0") & nl
    s = s & "hires_clock=" & HiOK & " scale_ns=" & Format(HiScale * 1000000000#, "0.000") & nl
    t = ""
    For i = 0 To NK - 1
        t = t & KName(i) & ":" & KHits(i) & ":" & KHeld(i) & ";"
    Next i
    s = s & "keys_name_presses_framesheld=" & t & nl
    s = s & "key_events_down=" & DbgDowns & nl & "key_events_up=" & DbgUps & nl & "key_events_taken_by_engine=" & DbgPops & nl
    s = s & "releases_held_back_for_engine=" & DeferRel & nl & "key_queue_max=" & QMax & nl
    s = s & "polls_skipped_by_throttle=" & PollSkips & nl
    s = s & "polls=" & DbgPolls & nl & "polls_excel_not_front=" & DbgNotFront & nl
    s = s & "mouse_frames_moved=" & MouseFrames & nl & "mouse_abs_dx=" & Format(MouseAbs, "0") & nl & "mouse_clicks=" & Clicks & nl
    s = s & "mouse_toggles=" & MToggles & nl & "mouse_on_at_end=" & MouseOn & nl
    s = s & "palette_change_frames=" & PalChanges & nl & "full_rescan_frames=" & ScanAllN & nl & "rows_scanned=" & RowsScanned & nl & "rows_skipped=" & RowsSkipped & nl
    s = s & "rects_total=" & Calls & nl & "excel_fill_calls=" & OMCalls & nl & "fill_errors=" & DbgFillErr & nl & "first_fill_error=" & DbgErrText & nl
    s = s & "settle_redraws=" & NSettle & nl & "keyheld_small_frames=" & NHeldBatch & nl & "tic_sleep_calls=" & SleepCalls & nl
    s = s & "frames_presented=" & Frames & nl & "mem_pages=" & MemPages & nl & "stdout_bytes=" & OutN & nl
    s = s & "stdout_tail=" & Replace(Replace(Right$(OutTail(), 200), vbLf, " / "), vbCr, "") & nl
    HostTele = s
End Function

Public Function Play() As String
    PresentMode = 0
    TeleToFile = True
    Play = RunDoom(False, 0)
End Function

Public Function PlayA() As String
    PresentMode = 4
    PlayA = RunDoom(False, 0)
End Function
Public Function PlayB() As String
    PresentMode = 128
    PlayB = RunDoom(False, 0)
End Function

' The macro players run (also wired to the "click here" panel on the sheet).
Public Sub PlayDoom()
    Dim s As String
    If Playing Then Exit Sub
    If Abs(Timer - LastEndT) < 1.5 And LastEndT <> 0 Then Exit Sub
    PresentMode = 0
    TeleToFile = False
    s = RunDoom(False, 0)
    If Left$(s, 2) <> "ok" And Left$(s, 4) <> "exit" Then MsgBox s
End Sub

' Same, but also writes a performance report to a text file in your home folder when you quit.
Public Sub PlayDoomWithStats()
    Dim s As String
    If Playing Then Exit Sub
    PresentMode = 0
    TeleToFile = True
    s = RunDoom(False, 0)
    MsgBox "Stats saved to " & TelePath()
End Sub

Public Function Verify() As String
    Verify = RunDoom(True, 40)
End Function

' Diagnostics for the picture. PlayRaw is a normal game that leaves the cells exactly as the last
' frame left them (no click-to-play sign). Dump then returns 64,000 colours, six hex digits each,
' row by row: 0 the engine's last frame through its palette, 1 what the painter believes each cell
' shows, 2 the fill colour each cell really has.
Public Function PlayRaw() As String
    PresentMode = 256
    TeleToFile = True
    PlayRaw = RunDoom(False, 0)
End Function

Public Function Dump(ByVal what As Long) As String
    Dim s As String, i As Long, c As Long, x As Long, y As Long
    On Error Resume Next
    s = String$(PW * PH * 6, "0")
    For y = 0 To PH - 1
        For x = 0 To PW - 1
            i = y * PW + x
            If what = 0 Then
                c = PalC(LD8(LastBuf + i))
            ElseIf what = 1 Then
                c = Shown(i)
            Else
                c = WS.Cells(GY + 1 + y, GX + 1 + x).Interior.Color
            End If
            Mid$(s, i * 6 + 1, 6) = Right$("00000" & Hex$(c And &HFFFFFF), 6)
        Next x
    Next y
    Dump = s
End Function
'''

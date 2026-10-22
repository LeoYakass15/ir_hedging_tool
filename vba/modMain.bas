Attribute VB_Name = "modMain"
'==============================================================================
' modMain - "Run Monte Carlo" button
'
' Everything with a closed form is a live worksheet formula (Schedule, Pricer,
' Scenarios, Risk, Curves, Pitch). The Monte Carlo is the only part run by a
' macro: it simulates Euribor 3M fixings and measures the all-in cost of the loan
' with each hedge.
'
' Scenario model: one Gaussian factor (forward-rate view of Hull-White),
'   X(t) Ornstein-Uhlenbeck with mean reversion a,
'   L_i = centre_i + sigma * sqrt(t_i) * X(t_i) / sd(X(t_i)),
' so each fixing has exactly the normal vol used to price the caps.
' centre = forward Euribor, or today's fixing ("unchanged" view).
'==============================================================================
Option Explicit

Private Const NSTRAT As Long = 5   ' unhedged, payer swap, floored swap, cap, collar

Public Function NamedCell(ByVal nm As String) As Range
    Set NamedCell = ThisWorkbook.Names(nm).RefersToRange
End Function

Public Function NamedValue(ByVal nm As String) As Variant
    NamedValue = ThisWorkbook.Names(nm).RefersToRange.Value
End Function

Private Sub SetStatus(ByVal msg As String)
    NamedCell("Status").Value = msg
    On Error Resume Next
    Application.StatusBar = msg
    On Error GoTo 0
End Sub

Private Function Pos(ByVal x As Double) As Double
    If x > 0 Then Pos = x Else Pos = 0
End Function

Public Sub RunMonteCarlo()
    Dim t0 As Double, msg As String
    t0 = Timer
    On Error GoTo Fail
    SetStatus "Monte Carlo running..."
    On Error Resume Next
    Application.Calculate
    On Error GoTo Fail
    msg = DoMonteCarlo()
    SetStatus "Monte Carlo " & Format(Now, "yyyy-mm-dd hh:mm") & " | " & msg & " | " & _
              Format(Timer - t0, "0.0") & " s"
    Exit Sub
Fail:
    SetStatus "Error in Monte Carlo: " & Err.Description
End Sub

Private Function DoMonteCarlo() As String
    Dim vF As Variant, vN As Variant, vA As Variant, vT As Variant
    Dim n As Long, nRows As Long, i As Long, r As Long, s As Long, p As Long, j As Long
    Dim F() As Double, bse() As Double, tf() As Double, centre() As Double
    Dim ex() As Double, sdStep() As Double, scl() As Double, z() As Double
    Dim sumBase As Double, margin As Double, floorOn As Boolean, budget As Double, ratio As Double
    Dim kS As Double, kFS As Double, kC As Double, kF As Double, prem As Double, ann1 As Double
    Dim sigma As Double, nPaths As Long, seed As Long, a As Double, fwdView As Boolean, fixNow As Double
    Dim premRun(1 To 5) As Double
    Dim allIn() As Double, worstQ() As Double, Lst() As Double
    Dim tot(1 To 5) As Double, wq(1 To 5) As Double
    Dim x As Double, L As Double, idx As Double, interest As Double, pay As Double, cost As Double
    Dim per As Double, sgn As Long, dt As Double, vr As Double
    Dim v() As Double, w() As Double, res(1 To 5, 1 To 5) As Variant
    Dim dist() As Variant, fan() As Variant, sumR As Double, nAbove As Long

    vF = NamedCell("Fwd").Value
    vN = NamedCell("LoanN").Value
    vA = NamedCell("Accr").Value
    vT = NamedCell("TFix").Value
    nRows = UBound(vF, 1)
    n = 0
    For r = 1 To nRows
        If CDbl(vN(r, 1)) > 0 Then n = n + 1
    Next r
    If n = 0 Then DoMonteCarlo = "no active period": Exit Function
    ReDim F(1 To n): ReDim bse(1 To n): ReDim tf(0 To n): ReDim centre(1 To n)
    ReDim ex(1 To n): ReDim sdStep(1 To n): ReDim scl(1 To n): ReDim z(1 To n)
    i = 0
    For r = 1 To nRows
        If CDbl(vN(r, 1)) > 0 Then
            i = i + 1
            F(i) = CDbl(vF(r, 1))
            bse(i) = CDbl(vN(r, 1)) * CDbl(vA(r, 1))
            tf(i) = CDbl(vT(r, 1))
            sumBase = sumBase + bse(i)
        End If
    Next r

    margin = CDbl(NamedValue("Margin"))
    floorOn = (UCase(Left(CStr(NamedValue("FloorYN")), 1)) = "Y")
    budget = CDbl(NamedValue("Budget"))
    ratio = CDbl(NamedValue("HedgeRatio"))
    kS = CDbl(NamedValue("SwapRate"))
    kFS = CDbl(NamedValue("FSwapRate"))
    kC = CDbl(NamedValue("CapStrike"))
    kF = CDbl(NamedValue("CollarFloor"))
    prem = CDbl(NamedValue("CapPremium"))
    ann1 = CDbl(NamedValue("Annuity1"))
    sigma = CDbl(NamedValue("CapVolBp")) / 10000#
    nPaths = CLng(NamedValue("NPaths"))
    If nPaths < 1000 Then nPaths = 1000
    seed = CLng(NamedValue("Seed"))
    a = CDbl(NamedValue("MeanRev"))
    If a < 0.00000001 Then a = 0.00000001
    fwdView = (InStr(1, CStr(NamedValue("ScenView")), "orward", vbTextCompare) > 0)
    fixNow = CDbl(NamedValue("FixingNow"))
    If prem > 0 And ann1 > 0 Then premRun(4) = prem / ann1

    tf(0) = 0
    For i = 1 To n
        If fwdView Then centre(i) = F(i) Else centre(i) = fixNow
        dt = tf(i) - tf(i - 1)
        ex(i) = Exp(-a * dt)
        sdStep(i) = Sqr((1# - Exp(-2# * a * dt)) / (2# * a))
        vr = (1# - Exp(-2# * a * tf(i))) / (2# * a)
        If vr > 1E-14 Then scl(i) = sigma * Sqr(tf(i)) / Sqr(vr) Else scl(i) = 0
    Next i

    ReDim allIn(1 To nPaths, 1 To NSTRAT): ReDim worstQ(1 To nPaths, 1 To NSTRAT)
    ReDim Lst(1 To nPaths, 1 To n)
    RngSeed seed
    p = 0
    Do While p < nPaths
        For i = 1 To n
            z(i) = RngNormal()
        Next i
        For sgn = 1 To -1 Step -2           ' antithetic pair
            p = p + 1
            If p > nPaths Then Exit For
            x = 0
            For s = 1 To NSTRAT
                tot(s) = 0: wq(s) = -1E+30
            Next s
            For i = 1 To n
                x = ex(i) * x + sdStep(i) * sgn * z(i)
                L = centre(i) + scl(i) * x
                Lst(p, i) = L
                idx = L
                If floorOn And idx < 0 Then idx = 0
                interest = bse(i) * (idx + margin)
                For s = 1 To NSTRAT
                    If s = 1 Then
                        pay = 0
                    ElseIf s = 2 Then
                        pay = L - kS
                    ElseIf s = 3 Then
                        pay = Pos(L) - kFS
                    ElseIf s = 4 Then
                        pay = Pos(L - kC)
                    Else
                        pay = Pos(L - kC) - Pos(kF - L)
                    End If
                    cost = interest - ratio * bse(i) * pay
                    tot(s) = tot(s) + cost
                    per = cost / bse(i) + premRun(s)
                    If per > wq(s) Then wq(s) = per
                Next s
            Next i
            For s = 1 To NSTRAT
                allIn(p, s) = (tot(s) + premRun(s) * sumBase) / sumBase
                worstQ(p, s) = wq(s)
            Next s
        Next sgn
    Loop

    ' statistics and distribution
    ReDim v(1 To nPaths): ReDim w(1 To nPaths)
    ReDim dist(1 To 99, 1 To NSTRAT + 1)
    For j = 1 To 99
        dist(j, 1) = j / 100#
    Next j
    For s = 1 To NSTRAT
        sumR = 0: nAbove = 0
        For p = 1 To nPaths
            v(p) = allIn(p, s): w(p) = worstQ(p, s)
            sumR = sumR + v(p)
            If v(p) > budget + 0.000000001 Then nAbove = nAbove + 1
        Next p
        QuickSort v, 1, nPaths
        QuickSort w, 1, nPaths
        res(s, 1) = sumR / nPaths
        res(s, 2) = PercentileSorted(v, 0.95)
        res(s, 3) = PercentileSorted(v, 0.05)
        res(s, 4) = PercentileSorted(w, 0.95)
        res(s, 5) = nAbove / nPaths
        For j = 1 To 99
            dist(j, s + 1) = PercentileSorted(v, j / 100#)
        Next j
    Next s

    ' fan of simulated fixings
    ReDim fan(1 To nRows, 1 To 5)
    For i = 1 To n
        For p = 1 To nPaths
            v(p) = Lst(p, i)
        Next p
        QuickSort v, 1, nPaths
        fan(i, 1) = PercentileSorted(v, 0.05)
        fan(i, 2) = PercentileSorted(v, 0.25)
        fan(i, 3) = PercentileSorted(v, 0.5)
        fan(i, 4) = PercentileSorted(v, 0.75)
        fan(i, 5) = PercentileSorted(v, 0.95)
    Next i
    For i = n + 1 To nRows
        For j = 1 To 5
            fan(i, j) = Empty
        Next j
    Next i

    NamedCell("McTop").Resize(NSTRAT, 5).Value = res
    NamedCell("DistTop").Resize(99, NSTRAT + 1).Value = dist
    NamedCell("FanTop").Resize(nRows, 5).Value = fan
    DoMonteCarlo = nPaths & " paths"
End Function

Attribute VB_Name = "modPricing"
'==============================================================================
' modPricing - hedges of a floating-rate loan, valued off the curves of modCurves
'
' With L the Euribor 3M fixing of a period, N the hedged notional and tau the
' ACT/360 accrual, the client receives per period:
'   SWAP     payer swap            N tau (L - K)
'   FSWAP    floored payer swap    N tau (max(L, Kf) - K)     (matches a floored loan)
'   CAP      cap                   N tau max(L - K, 0)
'   FLOOR    floor                 N tau max(K - L, 0)
'   COLLAR   collar                N tau [max(L - Kc, 0) - max(Kf - L, 0)]
' Caplets and floorlets: Bachelier (normal) model with the flat cap vol of the
' loan maturity.
'
' DealData (Risk sheet) = loan notional | maturity in months | linear (1/0) | hedge ratio
'
' Worksheet functions:
'   IR_PV(product, k1, k2, MktData, DealData, shiftBp, volShiftBp, keyRateMonths)
'       mid value to the client in EUR; "PAR" = par swap rate, "ANNUITY" = sum N tau DF
'   IR_COLLAR_FLOOR(capStrike, marginBp, MktData, DealData, shiftBp, volShiftBp)
'       floor strike that makes the collar zero-cost (minus a sales margin)
'==============================================================================
Option Explicit

Private nP As Long
Private pW() As Double      ' hedged notional x accrual x DF(ESTR, pay date)
Private pF() As Double      ' Euribor 3M forward
Private pSd() As Double     ' normal vol x sqrt(time to fixing)

' Fills the period grid of the loan for the curves in cache slot s
Private Sub BuildGrid(deal As Variant, ByVal s As Long, ByVal volShiftBp As Double)
    Dim a As Variant, r0 As Long, c0 As Long, notional As Double, months As Long
    Dim linear As Boolean, ratio As Double, i As Long, v As Double, d0 As Double, d1 As Double
    Dim sigma As Double, nAmort As Double
    If IsObject(deal) Then a = deal.Value2 Else a = deal
    r0 = LBound(a, 1): c0 = LBound(a, 2)
    notional = CDbl(a(r0, c0))
    months = CLng(a(r0 + 1, c0))
    linear = (CDbl(a(r0 + 2, c0)) <> 0)
    ratio = CDbl(a(r0 + 3, c0))
    nP = months \ 3
    ReDim pW(1 To nP): ReDim pF(1 To nP): ReDim pSd(1 To nP)
    v = SlotVal(s)
    sigma = SlotCapVol(s, (AddMonths(v, 3 * nP) - v) / 365#) + volShiftBp / 10000#
    For i = 1 To nP
        d0 = AddMonths(v, 3 * (i - 1))
        d1 = AddMonths(v, 3 * i)
        If linear Then nAmort = notional * (1# - (i - 1) / nP) Else nAmort = notional
        pW(i) = ratio * nAmort * (d1 - d0) / 360# * SlotDF(s, 1, d1)
        pF(i) = SlotFwd(s, 2, d0, d1)
        If sigma > 0 Then pSd(i) = sigma * Sqr((d0 - v) / 365#) Else pSd(i) = 0
    Next i
End Sub

Public Function Bachelier(ByVal f As Double, ByVal k As Double, ByVal sd As Double, ByVal isCall As Boolean) As Double
    Dim d As Double
    If sd <= 0.000000000001 Then
        If isCall Then
            If f > k Then Bachelier = f - k
        Else
            If k > f Then Bachelier = k - f
        End If
        Exit Function
    End If
    d = (f - k) / sd
    If isCall Then
        Bachelier = (f - k) * NormCDF(d) + sd * NormPDF(d)
    Else
        Bachelier = (k - f) * NormCDF(-d) + sd * NormPDF(d)
    End If
End Function

Private Function OptStrip(ByVal k As Double, ByVal isCall As Boolean) As Double
    Dim i As Long, t As Double
    For i = 1 To nP
        t = t + pW(i) * Bachelier(pF(i), k, pSd(i), isCall)
    Next i
    OptStrip = t
End Function

Private Function GridValue(ByVal product As String, ByVal k1 As Double, ByVal k2 As Double) As Double
    Dim i As Long, t As Double, w As Double
    Select Case UCase(product)
        Case "SWAP", "FSWAP"
            For i = 1 To nP
                t = t + pW(i) * (pF(i) - k1)
            Next i
            If UCase(product) = "FSWAP" Then t = t + OptStrip(k2, False)
        Case "CAP"
            t = OptStrip(k1, True)
        Case "FLOOR"
            t = OptStrip(k1, False)
        Case "COLLAR"
            t = OptStrip(k1, True) - OptStrip(k2, False)
        Case "ANNUITY"
            For i = 1 To nP
                t = t + pW(i)
            Next i
        Case "PAR"
            For i = 1 To nP
                t = t + pW(i) * pF(i)
                w = w + pW(i)
            Next i
            If w <> 0 Then t = t / w
    End Select
    GridValue = t
End Function

Public Function IR_PV(ByVal product As String, ByVal k1 As Double, ByVal k2 As Double, mkt As Variant, _
                      deal As Variant, ByVal shiftBp As Double, ByVal volShiftBp As Double, _
                      ByVal keyRateMonths As Double) As Variant
    Dim s As Long
    On Error GoTo Fail
    s = GetCurves(mkt, shiftBp, CLng(keyRateMonths))
    If s = 0 Then IR_PV = "no market": Exit Function
    BuildGrid deal, s, volShiftBp
    IR_PV = GridValue(product, k1, k2)
    Exit Function
Fail:
    IR_PV = "error: " & Err.Description
End Function

Public Function IR_COLLAR_FLOOR(ByVal capStrike As Double, ByVal marginBp As Double, mkt As Variant, _
                                deal As Variant, ByVal shiftBp As Double, ByVal volShiftBp As Double) As Variant
    Dim s As Long, margin As Double, lo As Double, hi As Double, xm As Double
    Dim flo As Double, fmid As Double, it As Long
    On Error GoTo Fail
    s = GetCurves(mkt, shiftBp, 0)
    If s = 0 Then IR_COLLAR_FLOOR = "no market": Exit Function
    BuildGrid deal, s, volShiftBp
    margin = marginBp / 10000# * GridValue("ANNUITY", 0, 0)
    lo = -0.05: hi = capStrike - 0.000001
    flo = GridValue("COLLAR", capStrike, lo) + margin
    If flo * (GridValue("COLLAR", capStrike, hi) + margin) > 0 Then
        IR_COLLAR_FLOOR = "none"
        Exit Function
    End If
    For it = 1 To 80
        xm = 0.5 * (lo + hi)
        fmid = GridValue("COLLAR", capStrike, xm) + margin
        If fmid * flo > 0 Then
            lo = xm: flo = fmid
        Else
            hi = xm
        End If
        If hi - lo < 1E-13 Then Exit For
    Next it
    IR_COLLAR_FLOOR = 0.5 * (lo + hi)
    Exit Function
Fail:
    IR_COLLAR_FLOOR = "error: " & Err.Description
End Function

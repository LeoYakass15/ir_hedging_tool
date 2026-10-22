Attribute VB_Name = "modCurves"
'==============================================================================
' modCurves - EUR dual-curve bootstrap, exposed as worksheet functions
'
'   ESTR curve     discounting, bootstrapped from EUR STR OIS quotes
'                  (<= 1Y: single period; > 1Y: annual ACT/360 fixed leg)
'   EURIBOR curve  projection of Euribor 3M, from today's fixing and par swaps
'                  (annual 30/360 fixed vs quarterly Euribor, discounted on ESTR)
'
' Both curves are discount factors at pillar dates (ACT/365 times), interpolated
' log-linearly (piecewise-flat forwards), flat-forward extrapolation.
'
' The market table (named range MktData on the Market sheet) is passed to every
' function as an argument, so Excel recalculates them when a quote changes.
' Columns of MktData: instrument | tenor | months | quote. Rows:
'   Valuation date | OIS ... | Fixing Euribor 3M | IRS ... | Cap vol (bp) ...
' Rates are decimals (0.0190 = 1.90%), cap vols are normal vols in bp.
'
' Worksheet functions:
'   IR_DF(curve, date, MktData, shiftBp)            discount factor ("ESTR" or "EURIBOR")
'   IR_QUOTE(instrument, months, MktData, shiftBp)  quote repriced off the curves
'   IR_CAPVOL(years, MktData, volShiftBp)           flat cap normal vol in bp
' shiftBp moves every rate quote in parallel (scenario or DV01).
'==============================================================================
Option Explicit

Private Const MAXC As Long = 64      ' curves kept in memory (base + bumped)
Private Const MAXP As Long = 40      ' pillars per curve

Private cKey(1 To MAXC) As String
Private cVal(1 To MAXC) As Double
Private cN(1 To MAXC, 1 To 2) As Long
Private cT(1 To MAXC, 1 To 2, 0 To MAXP) As Double
Private cL(1 To MAXC, 1 To 2, 0 To MAXP) As Double
Private cNV(1 To MAXC) As Long
Private cTV(1 To MAXC, 1 To MAXP) As Double
Private cVV(1 To MAXC, 1 To MAXP) As Double
Private cCount As Long, cNext As Long

' parsed market (scratch)
Private qVal As Double
Private nO As Long, oM(1 To MAXP) As Long, oQ(1 To MAXP) As Double
Private nS As Long, sM(1 To MAXP) As Long, sQ(1 To MAXP) As Double
Private nV As Long, vM(1 To MAXP) As Long, vQ(1 To MAXP) As Double
Private qFix As Double

'------------------------------------------------------------------ dates
Public Function AddMonths(ByVal serial As Double, ByVal k As Long) As Double
    Dim d As Date, y As Long, m As Long, dd As Long, y2 As Long, m2 As Long, lastDay As Long
    d = CDate(serial)
    y = Year(d): m = Month(d) + k: dd = Day(d)
    y = y + (m - 1) \ 12
    m = (m - 1) Mod 12 + 1
    y2 = y: m2 = m + 1
    If m2 > 12 Then m2 = 1: y2 = y2 + 1
    lastDay = Day(DateSerial(y2, m2, 1) - 1)
    If dd > lastDay Then dd = lastDay
    AddMonths = CDbl(DateSerial(y, m, dd))
End Function

Public Function Yf30360(ByVal d1 As Double, ByVal d2 As Double) As Double
    Dim a As Date, b As Date, dd1 As Long, dd2 As Long
    a = CDate(d1): b = CDate(d2)
    dd1 = Day(a): If dd1 > 30 Then dd1 = 30
    dd2 = Day(b)
    If dd1 = 30 And dd2 = 31 Then dd2 = 30
    Yf30360 = ((Year(b) - Year(a)) * 360 + (Month(b) - Month(a)) * 30 + (dd2 - dd1)) / 360#
End Function

'------------------------------------------------------------------ market parsing
Private Function AsArray(x As Variant) As Variant
    If IsObject(x) Then AsArray = x.Value2 Else AsArray = x
End Function

Private Sub AddSorted(ByRef n As Long, mm() As Long, qq() As Double, ByVal m As Long, ByVal q As Double)
    Dim i As Long
    n = n + 1
    i = n
    Do While i > 1
        If mm(i - 1) <= m Then Exit Do
        mm(i) = mm(i - 1): qq(i) = qq(i - 1)
        i = i - 1
    Loop
    mm(i) = m: qq(i) = q
End Sub

' Reads MktData, returns the cache key
Private Function ParseMarket(mkt As Variant) As String
    Dim a As Variant, r As Long, c0 As Long, code As String, key As String
    Dim m As Long, q As Double
    a = AsArray(mkt)
    c0 = LBound(a, 2)
    nO = 0: nS = 0: nV = 0: qFix = 0: qVal = 0
    For r = LBound(a, 1) To UBound(a, 1)
        code = UCase(Trim(CStr(a(r, c0))))
        If Len(code) > 0 And IsNumeric(a(r, c0 + 3)) And Not IsEmpty(a(r, c0 + 3)) Then
            q = CDbl(a(r, c0 + 3))
            m = 0
            If IsNumeric(a(r, c0 + 2)) And Not IsEmpty(a(r, c0 + 2)) Then m = CLng(a(r, c0 + 2))
            key = key & Left(code, 3) & m & "=" & CStr(q) & ";"
            If Left(code, 3) = "VAL" Then
                qVal = q
            ElseIf Left(code, 3) = "OIS" Then
                If m > 0 Then AddSorted nO, oM, oQ, m, q
            ElseIf Left(code, 3) = "FIX" Then
                qFix = q
            ElseIf Left(code, 3) = "IRS" Then
                If m > 0 Then AddSorted nS, sM, sQ, m, q
            ElseIf Left(code, 3) = "CAP" Then
                If m > 0 Then AddSorted nV, vM, vQ, m, q
            End If
        End If
    Next r
    ParseMarket = key
End Function

'------------------------------------------------------------------ curve maths
' log discount factor of curve c (1 = ESTR, 2 = EURIBOR) in cache slot s at time t
Private Function LogDF(ByVal s As Long, ByVal c As Long, ByVal t As Double) As Double
    Dim n As Long, k As Long, r As Double
    n = cN(s, c)
    If t <= 0# Or n = 0 Then LogDF = 0#: Exit Function
    If t >= cT(s, c, n) Then
        r = (cL(s, c, n) - cL(s, c, n - 1)) / (cT(s, c, n) - cT(s, c, n - 1))
        LogDF = cL(s, c, n) + r * (t - cT(s, c, n))
        Exit Function
    End If
    k = 1
    Do While cT(s, c, k) < t
        k = k + 1
    Loop
    LogDF = cL(s, c, k - 1) + (cL(s, c, k) - cL(s, c, k - 1)) * (t - cT(s, c, k - 1)) / _
            (cT(s, c, k) - cT(s, c, k - 1))
End Function

Public Function SlotDF(ByVal s As Long, ByVal c As Long, ByVal d As Double) As Double
    SlotDF = Exp(LogDF(s, c, (d - cVal(s)) / 365#))
End Function

Public Function SlotVal(ByVal s As Long) As Double
    SlotVal = cVal(s)
End Function

' Simple ACT/360 forward of curve c between two dates
Public Function SlotFwd(ByVal s As Long, ByVal c As Long, ByVal d1 As Double, ByVal d2 As Double) As Double
    SlotFwd = (SlotDF(s, c, d1) / SlotDF(s, c, d2) - 1#) / ((d2 - d1) / 360#)
End Function

' Flat cap normal vol (decimal) for a cap maturity in years, linear in maturity, flat outside
Public Function SlotCapVol(ByVal s As Long, ByVal yrs As Double) As Double
    Dim n As Long, k As Long, t0 As Double, t1 As Double
    n = cNV(s)
    If n = 0 Then SlotCapVol = 0: Exit Function
    If yrs <= cTV(s, 1) Then SlotCapVol = cVV(s, 1): Exit Function
    If yrs >= cTV(s, n) Then SlotCapVol = cVV(s, n): Exit Function
    For k = 2 To n
        If yrs <= cTV(s, k) Then
            t0 = cTV(s, k - 1): t1 = cTV(s, k)
            SlotCapVol = cVV(s, k - 1) + (cVV(s, k) - cVV(s, k - 1)) * (yrs - t0) / (t1 - t0)
            Exit Function
        End If
    Next k
End Function

Private Function OisPar(ByVal s As Long, ByVal m As Long) As Double
    Dim v As Double, e As Double, k As Long, prev As Double, dk As Double, p As Double, ann As Double
    v = cVal(s)
    If m <= 12 Then
        e = AddMonths(v, m)
        OisPar = (1# / SlotDF(s, 1, e) - 1#) / ((e - v) / 360#)
    Else
        prev = v
        For k = 1 To m \ 12
            dk = AddMonths(v, 12 * k)
            p = SlotDF(s, 1, dk)
            ann = ann + (dk - prev) / 360# * p
            prev = dk
        Next k
        OisPar = (1# - p) / ann
    End If
End Function

Private Function IrsPar(ByVal s As Long, ByVal m As Long) As Double
    Dim v As Double, k As Long, prev As Double, dk As Double, ann As Double, fl As Double
    v = cVal(s)
    prev = v
    For k = 1 To m \ 12
        dk = AddMonths(v, 12 * k)
        ann = ann + Yf30360(prev, dk) * SlotDF(s, 1, dk)
        prev = dk
    Next k
    prev = v
    For k = 1 To m \ 3
        dk = AddMonths(v, 3 * k)
        fl = fl + (SlotDF(s, 2, prev) / SlotDF(s, 2, dk) - 1#) * SlotDF(s, 1, dk)
        prev = dk
    Next k
    IrsPar = fl / ann
End Function

' Secant search on the log discount factor of the last pillar of curve c
Private Sub SolvePillar(ByVal s As Long, ByVal c As Long, ByVal m As Long, ByVal target As Double)
    Dim n As Long, x0 As Double, x1 As Double, x2 As Double, f0 As Double, f1 As Double, it As Long
    n = cN(s, c)
    x0 = cL(s, c, n - 1) - target * (cT(s, c, n) - cT(s, c, n - 1))
    x1 = x0 - 0.0001
    cL(s, c, n) = x0
    If c = 1 Then f0 = OisPar(s, m) - target Else f0 = IrsPar(s, m) - target
    For it = 1 To 60
        cL(s, c, n) = x1
        If c = 1 Then f1 = OisPar(s, m) - target Else f1 = IrsPar(s, m) - target
        If Abs(f1) < 1E-15 Or f1 = f0 Then Exit For
        x2 = x1 - f1 * (x1 - x0) / (f1 - f0)
        x0 = x1: f0 = f1: x1 = x2
    Next it
    cL(s, c, n) = x1
End Sub

Private Sub Bootstrap(ByVal s As Long, ByVal shift As Double, ByVal krMonths As Long)
    Dim i As Long, e As Double, n As Long, q As Double
    cVal(s) = qVal
    ' ESTR
    cN(s, 1) = 0: cT(s, 1, 0) = 0: cL(s, 1, 0) = 0
    For i = 1 To nO
        e = AddMonths(qVal, oM(i))
        n = cN(s, 1) + 1
        cN(s, 1) = n
        cT(s, 1, n) = (e - qVal) / 365#
        q = oQ(i) + shift
        If oM(i) <= 12 Then
            cL(s, 1, n) = -Log(1# + q * (e - qVal) / 360#)
        Else
            SolvePillar s, 1, oM(i), q
        End If
    Next i
    ' EURIBOR 3M: fixing, then swaps
    cN(s, 2) = 1: cT(s, 2, 0) = 0: cL(s, 2, 0) = 0
    e = AddMonths(qVal, 3)
    cT(s, 2, 1) = (e - qVal) / 365#
    cL(s, 2, 1) = -Log(1# + (qFix + shift) * (e - qVal) / 360#)
    For i = 1 To nS
        If sM(i) > 3 Then
            e = AddMonths(qVal, sM(i))
            n = cN(s, 2) + 1
            cN(s, 2) = n
            cT(s, 2, n) = (e - qVal) / 365#
            q = sQ(i) + shift
            If sM(i) = krMonths Then q = q + 0.0001
            SolvePillar s, 2, sM(i), q
        End If
    Next i
    ' cap vols (decimal)
    cNV(s) = nV
    For i = 1 To nV
        cTV(s, i) = vM(i) / 12#
        cVV(s, i) = vQ(i) / 10000#
    Next i
End Sub

' Cache slot holding the curves for this market, shift (bp) and key-rate bump
Public Function GetCurves(mkt As Variant, ByVal shiftBp As Double, ByVal krMonths As Long) As Long
    Dim key As String, i As Long, s As Long
    key = ParseMarket(mkt) & "|" & CStr(shiftBp) & "|" & krMonths
    For i = 1 To cCount
        If cKey(i) = key Then GetCurves = i: Exit Function
    Next i
    If qVal <= 0 Or nO = 0 Or nS = 0 Then GetCurves = 0: Exit Function
    If cCount < MAXC Then
        cCount = cCount + 1
        s = cCount
    Else
        cNext = cNext Mod MAXC + 1
        s = cNext
    End If
    cKey(s) = ""
    Bootstrap s, shiftBp / 10000#, krMonths
    cKey(s) = key
    GetCurves = s
End Function

Private Function CurveIndex(ByVal curve As String) As Long
    If InStr(1, UCase(curve), "EURIBOR") > 0 Then CurveIndex = 2 Else CurveIndex = 1
End Function

'------------------------------------------------------------------ worksheet functions
Public Function IR_DF(ByVal curve As String, ByVal d As Double, mkt As Variant, ByVal shiftBp As Double) As Variant
    Dim s As Long
    s = GetCurves(mkt, shiftBp, 0)
    If s = 0 Then IR_DF = "no market": Exit Function
    IR_DF = SlotDF(s, CurveIndex(curve), d)
End Function

Public Function IR_QUOTE(ByVal instrument As String, ByVal months As Double, mkt As Variant, _
                         ByVal shiftBp As Double) As Variant
    Dim s As Long, code As String, v As Double
    s = GetCurves(mkt, shiftBp, 0)
    If s = 0 Then IR_QUOTE = "no market": Exit Function
    code = Left(UCase(Trim(instrument)), 3)
    v = cVal(s)
    If code = "OIS" Then
        IR_QUOTE = OisPar(s, CLng(months))
    ElseIf code = "IRS" Then
        IR_QUOTE = IrsPar(s, CLng(months))
    ElseIf code = "FIX" Then
        IR_QUOTE = SlotFwd(s, 2, v, AddMonths(v, 3))
    Else
        IR_QUOTE = ""
    End If
End Function

Public Function IR_CAPVOL(ByVal yrs As Double, mkt As Variant, ByVal volShiftBp As Double) As Variant
    Dim s As Long
    s = GetCurves(mkt, 0, 0)
    If s = 0 Then IR_CAPVOL = "no market": Exit Function
    IR_CAPVOL = SlotCapVol(s, yrs) * 10000# + volShiftBp
End Function

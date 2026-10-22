Attribute VB_Name = "modMath"
'==============================================================================
' modMath - numerical building blocks (no dependency on Excel worksheet functions)
'   * MRG32k3a random number generator (L'Ecuyer), reproducible across machines
'   * inverse normal (Acklam) and normal CDF (Hart, double precision)
'   * linear interpolation, sorting and percentiles (numpy "linear" convention)
'==============================================================================
Option Explicit

Private s10 As Double, s11 As Double, s12 As Double
Private s20 As Double, s21 As Double, s22 As Double

Private Const M1 As Double = 4294967087#
Private Const M2 As Double = 4294944443#
Private Const A12 As Double = 1403580#
Private Const A13N As Double = 810728#
Private Const A21 As Double = 527612#
Private Const A23N As Double = 1370589#
Private Const RNG_NORM As Double = 2.32830654929573E-10

Public Sub RngSeed(ByVal seed As Long)
    s10 = 12345# + Abs(seed): s11 = 12345#: s12 = 12345#
    s20 = 12345#: s21 = 12345# + Abs(seed): s22 = 12345#
End Sub

Public Function RngUniform() As Double
    Dim p1 As Double, p2 As Double, k As Double
    p1 = A12 * s11 - A13N * s10
    k = Int(p1 / M1)
    p1 = p1 - k * M1
    If p1 < 0 Then p1 = p1 + M1
    s10 = s11: s11 = s12: s12 = p1

    p2 = A21 * s22 - A23N * s20
    k = Int(p2 / M2)
    p2 = p2 - k * M2
    If p2 < 0 Then p2 = p2 + M2
    s20 = s21: s21 = s22: s22 = p2

    If p1 <= p2 Then
        RngUniform = (p1 - p2 + M1) * RNG_NORM
    Else
        RngUniform = (p1 - p2) * RNG_NORM
    End If
End Function

Public Function RngNormal() As Double
    RngNormal = NormInv(RngUniform())
End Function

' Inverse of the standard normal CDF (P.J. Acklam, relative error < 1.2e-9)
Public Function NormInv(ByVal p As Double) As Double
    Const a1 As Double = -39.6968302866538, a2 As Double = 220.946098424521
    Const a3 As Double = -275.928510446969, a4 As Double = 138.357751867269
    Const a5 As Double = -30.6647980661472, a6 As Double = 2.50662827745924
    Const b1 As Double = -54.4760987982241, b2 As Double = 161.585836858041
    Const b3 As Double = -155.698979859887, b4 As Double = 66.8013118877197
    Const b5 As Double = -13.2806815528857
    Const c1 As Double = -7.78489400243029E-03, c2 As Double = -0.322396458041136
    Const c3 As Double = -2.40075827716184, c4 As Double = -2.54973253934373
    Const c5 As Double = 4.37466414146497, c6 As Double = 2.93816398269878
    Const d1 As Double = 7.78469570904146E-03, d2 As Double = 0.32246712907004
    Const d3 As Double = 2.445134137143, d4 As Double = 3.75440866190742
    Const PLOW As Double = 0.02425
    Dim q As Double, r As Double
    If p <= 0# Then p = 1E-300
    If p >= 1# Then p = 1# - 1E-16
    If p < PLOW Then
        q = Sqr(-2# * Log(p))
        NormInv = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                  ((((d1 * q + d2) * q + d3) * q + d4) * q + 1#)
    ElseIf p <= 1# - PLOW Then
        q = p - 0.5
        r = q * q
        NormInv = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q / _
                  (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1#)
    Else
        q = Sqr(-2# * Log(1# - p))
        NormInv = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                   ((((d1 * q + d2) * q + d3) * q + d4) * q + 1#)
    End If
End Function

' Standard normal CDF (Hart 1968 as given by G. West 2005, double precision)
Public Function NormCDF(ByVal x As Double) As Double
    Dim xa As Double, e As Double, b As Double, c As Double
    xa = Abs(x)
    If xa > 37# Then
        c = 0#
    Else
        e = Exp(-xa * xa / 2#)
        If xa < 7.07106781186547 Then
            b = 3.52624965998911E-02 * xa + 0.700383064443688
            b = b * xa + 6.37396220353165
            b = b * xa + 33.912866078383
            b = b * xa + 112.079291497871
            b = b * xa + 221.213596169931
            b = b * xa + 220.206867912376
            c = e * b
            b = 8.83883476483184E-02 * xa + 1.75566716318264
            b = b * xa + 16.064177579207
            b = b * xa + 86.7807322029461
            b = b * xa + 296.564248779674
            b = b * xa + 637.333633378831
            b = b * xa + 793.826512519948
            b = b * xa + 440.413735824752
            c = c / b
        Else
            b = xa + 0.65
            b = xa + 4# / b
            b = xa + 3# / b
            b = xa + 2# / b
            b = xa + 1# / b
            c = e / b / 2.506628274631
        End If
    End If
    If x > 0 Then c = 1# - c
    NormCDF = c
End Function

Public Function NormPDF(ByVal x As Double) As Double
    NormPDF = Exp(-0.5 * x * x) / 2.506628274631
End Function

' Linear interpolation with flat extrapolation (same as numpy.interp)
Public Function Interp(ByVal x As Double, xs() As Double, ys() As Double) As Double
    Dim lo As Long, hi As Long, i As Long, w As Double
    lo = LBound(xs): hi = UBound(xs)
    If x <= xs(lo) Then Interp = ys(lo): Exit Function
    If x >= xs(hi) Then Interp = ys(hi): Exit Function
    For i = lo + 1 To hi
        If x <= xs(i) Then
            w = (x - xs(i - 1)) / (xs(i) - xs(i - 1))
            Interp = ys(i - 1) + w * (ys(i) - ys(i - 1))
            Exit Function
        End If
    Next i
    Interp = ys(hi)
End Function

' In-place ascending quicksort on v(lo..hi)
Public Sub QuickSort(v() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Double, tmp As Double
    Do While lo < hi
        i = lo: j = hi
        pivot = v((lo + hi) \ 2)
        Do While i <= j
            Do While v(i) < pivot: i = i + 1: Loop
            Do While v(j) > pivot: j = j - 1: Loop
            If i <= j Then
                tmp = v(i): v(i) = v(j): v(j) = tmp
                i = i + 1: j = j - 1
            End If
        Loop
        ' recurse on the smaller side to bound the stack depth
        If j - lo < hi - i Then
            If lo < j Then QuickSort v, lo, j
            lo = i
        Else
            If i < hi Then QuickSort v, i, hi
            hi = j
        End If
    Loop
End Sub

' q in [0,1]; v must be sorted ascending (1-based)
Public Function PercentileSorted(v() As Double, ByVal q As Double) As Double
    Dim n As Long, pos As Double, k As Long, f As Double
    n = UBound(v) - LBound(v) + 1
    pos = (n - 1) * q
    k = Int(pos)
    f = pos - k
    If k + 1 >= n Then
        PercentileSorted = v(LBound(v) + n - 1)
    Else
        PercentileSorted = v(LBound(v) + k) + f * (v(LBound(v) + k + 1) - v(LBound(v) + k))
    End If
End Function

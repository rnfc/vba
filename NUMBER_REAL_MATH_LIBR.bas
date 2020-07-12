Attribute VB_Name = "NUMBER_REAL_MATH_LIBR"
'********************************************************************************
'* Module: mMathSpecFun.bas                              v.1.0.0 June 2004      *
'*                                                       by Leonardo Volpi      *
'*                                                          & Foxes Team        *
'*                                                                              *
'*                                                                              *
'*                                                                              *
'*  Math Special Functions Library for clsMathParser.cls class v.4x             *
'*  for VB 6, VBA 97/2000/XP                                                    *
'********************************************************************************
Option Explicit
Option Private Module
Const PI_ As Double = 3.14159265358979
'
'*******************************************************************************
' CREDITS                                                                       '
' Many routines of this VB module was derived from the                          '
' LIBRARY FOR COMPUTATION of SPECIAL FUNCTIONS written in FORTRAN-77            '
' by Shanjie Zhang and Jianming Jin.                                            '
' All these programs and subroutines are copyrighted.                           '
' However, authors give kindly permission to incorporate any of these           '
' routines into other programs providing that the copyright is acknowledged.    '
' We have modified only minimal parts in order to adapt them to VB and VBA.     '
'*******************************************************************************

'-------------------------------------------------------------------------------
' error distribution function
'-------------------------------------------------------------------------------
Sub Herf(ByVal x As Double, ByRef Y As Double)
Const MaxLoop As Long = 400
Const tiny As Double = 0.000000000000001
Dim t As Double, p As Double, s As Double, i As Long
Dim A0 As Double, B0 As Double, A1 As Double, b1 As Double, A2 As Double, b2 As Double
Dim F1 As Double, F2 As Double, g As Double, D As Double
If x <= 2 Then
    t = 2 * x * x
    p = 1
    s = 1
    For i = 3 To MaxLoop Step 2
        p = p * t / i
        s = s + p
        If p < tiny Then Exit For
    Next
    Y = 2 * s * x * Exp(-x * x) / Sqr(PI_)
Else
    A0 = 1: B0 = 0
    A1 = 0: b1 = 1
    F1 = 0
    For i = 1 To MaxLoop
        g = 2 - (i Mod 2)
        A2 = g * x * A1 + i * A0
        b2 = g * x * b1 + i * B0
        F2 = A2 / b2
        D = Abs(F2 - F1)
        If D < tiny Then Exit For
        A0 = A1: B0 = b1
        A1 = A2: b1 = b2
        F1 = F2
    Next
    Y = 1 - 2 * Exp(-x * x) / (2 * x + F2) / Sqr(PI_)
End If
End Sub

'-------------------------------------------------------------------------------
' gamma function
'-------------------------------------------------------------------------------
Sub HGamma(ByVal x As Double, ByRef Y As Double)
'compute y = gamma(x)
Dim mantissa As Double, Expo As Double
    gamma_split x, mantissa, Expo
    Y = mantissa * 10 ^ Expo
End Sub

' gamma  - Lanczos approximation algorithm for gamma function
Sub gamma_split(ByVal x As Double, ByRef mantissa As Double, ByRef Expo As Double)
Dim z As Double, cf(14) As Double, W As Double, i As Long, s As Double, p As Double
Const DOUBLEPI As Double = 6.28318530717959
Const G_ As Double = 4.7421875  '607/128
    z = x - 1
    
    cf(0) = 0.999999999999997
    cf(1) = 57.1562356658629
    cf(2) = -59.5979603554755
    cf(3) = 14.1360979747417
    cf(4) = -0.49191381609762
    cf(5) = 3.39946499848119E-05
    cf(6) = 4.65236289270486E-05
    cf(7) = -9.83744753048796E-05
    cf(8) = 1.58088703224912E-04
    cf(9) = -2.10264441724105E-04
    cf(10) = 2.17439618115213E-04
    cf(11) = -1.64318106536764E-04
    cf(12) = 8.44182239838528E-05
    cf(13) = -2.61908384015814E-05
    cf(14) = 3.68991826595316E-06
    
    W = Exp(G_) / Sqr(DOUBLEPI)
    s = cf(0)
    For i = 1 To 14
        s = s + cf(i) / (z + i)
    Next
    s = s / W
    p = Log((z + G_ + 0.5) / Exp(1)) * (z + 0.5) / Log(10)
    'split in mantissa and exponent to avoid overflow
    Expo = Int(p)
    p = p - Int(p)
    mantissa = 10 ^ p * s
    'rescaling
    p = Int(Log(mantissa) / Log(10))
    mantissa = mantissa * 10 ^ -p
    Expo = Expo + p
End Sub

'-------------------------------------------------------------------------------
' logarithm gamma function
'-------------------------------------------------------------------------------
Private Function gammaln_(ByVal x)
Dim mantissa As Double, Expo As Double
    gamma_split x, mantissa, Expo
    gammaln_ = Log(mantissa) + Expo * Log(10)
End Function

'-------------------------------------------------------------------------------
' beta function
'---------------------------------------------------------------------------------
Sub HBeta(z, W, Y)
Y = Exp(gammaln_(z) + gammaln_(W) - gammaln_(z + W))
End Sub

'-------------------------------------------------------------------------------
' digamma function
'-------------------------------------------------------------------------------
Sub HDigamma(ByVal x As Double, ByRef Y As Double)
Dim b1(11) As Double, b2(11) As Double
Dim z As Double, s As Double, k As Long, tmp As Double
Const LIM_LOW As Long = 8
'Bernoulli's numbers
b1(0) = 1: b2(0) = 1
b1(1) = 1: b2(1) = 6
b1(2) = -1: b2(2) = 30
b1(3) = 1: b2(3) = 42
b1(4) = -1: b2(4) = 30
b1(5) = 5: b2(5) = 66
b1(6) = -691: b2(6) = 2730
b1(7) = 7: b2(7) = 6
b1(8) = -3617: b2(8) = 360
b1(9) = 43867: b2(9) = 798
b1(10) = -174611: b2(10) = 330
b1(11) = 854513: b2(11) = 138
If x <= LIM_LOW Then
    z = x - 1 + LIM_LOW
Else
    z = x - 1
End If
s = 0
For k = 1 To 11
    tmp = b1(k) / b2(k) / k / z ^ (2 * k)
    s = s + tmp
Next
Y = Log(z) + 0.5 * (1 / z - s)

If x <= LIM_LOW Then
    s = 0
    For k = 0 To LIM_LOW - 1
        s = s + 1 / (x + k)
    Next
    Y = Y - s
End If
End Sub

'-------------------------------------------------------------------------------
' Riemman's zeta function
'-------------------------------------------------------------------------------
Sub HZeta(ByVal x As Double, ByRef Y As Double)
Dim Cnk As Double, k As Long, n As Long, s As Double, s1 As Double, coeff As Double
Const N_MAX As Long = 1000
Const tiny As Double = 1E-16
n = 0: s = 0
Do
    s1 = 0: Cnk = 1
    For k = 0 To n
        If k > 0 Then Cnk = Cnk * (n - k + 1) / k
        s1 = s1 + (-1) ^ k * Cnk / (k + 1) ^ x
    Next k
    coeff = s1 / 2 ^ (1 + n)
    s = s + coeff
    n = n + 1
Loop Until Abs(coeff) < tiny Or n > N_MAX
Y = s / (1 - 2 ^ (1 - x))
End Sub

'-------------------------------------------------------------------------------
' exponential integral Ei(x) for x >0.
'-------------------------------------------------------------------------------
Sub Hexp_integr(ByVal x As Double, ByRef Y As Double)
Dim k As Long, fact As Double, prev As Double, Sum As Double, term As Double
Const eps As Double = 0.000000000000001
Const EULER As Double = 0.577215664901532
Const MAXIT As Long = 100, FPMIN As Double = 1E-30
If (x <= 0) Then Exit Sub      '
If (x < FPMIN) Then          'Special case: avoid failure of convergence test be-
      Y = Log(x) + EULER            'cause of under ow.
ElseIf (x <= -Log(eps)) Then 'Use power series.
     Sum = 0
     fact = 1
     For k = 1 To MAXIT
        fact = fact * x / k
        term = fact / k
        Sum = Sum + term
        If (term < eps * Sum) Then Exit For
    Next
    Y = Sum + Log(x) + EULER
Else 'Use asymptotic series.
    Sum = 0 'Start with second term.
    term = 1
    For k = 1 To MAXIT
      prev = term
      term = term * k / x
      If (term < eps) Then Exit For 'Since al sum is greater than one, term itself ap-
      If (term < prev) Then
         Sum = Sum + term 'Still converging: add new term.
      Else
         Sum = Sum - prev 'Diverging: subtract previous term and exit.
         Exit For
      End If
    Next
    Y = Exp(x) * (1 + Sum) / x
End If
End Sub


Sub Hexpn_integr(ByVal x As Double, ByVal n As Double, ByRef Y As Double)
'Evaluates the exponential integral En(x).
'Parameters: MAXIT is the maximum allowed number of iterations; EPS is the desired rel-
'ative error, not smaller than the machine precision; FPMIN is a number near the smallest
'representable foating-point number; EULER is Euler's constant .
Const MAXIT As Long = 100
Const eps As Double = 0.000000000000001
Const FPMIN As Double = 1E-30
Const EULER As Double = 0.577215664901532
Dim nm1 As Long, A As Double, B As Double, c As Double, D As Double, h As Double, i As Long, del As Double, fact As Double, Psi As Double, ii As Long
nm1 = n - 1
If (n < 0 Or x < 0 Or (x = 0 And (n = 0 Or n = 1))) Then
   Exit Sub
ElseIf (n = 0) Then 'Special case.
   Y = Exp(-x) / x
ElseIf (x = 0) Then 'Another special case.
   Y = 1 / nm1
ElseIf (x > 1) Then 'Lentz's algorithm .
   B = x + n
   c = 1 / FPMIN
   D = 1 / B
   h = D
   For i = 1 To MAXIT
      A = -i * (nm1 + i)
      B = B + 2
      D = 1 / (A * D + B)  'Denominators cannot be zero.
      c = B + A / c
      del = c * D
      h = h * del
      If (Abs(del - 1) < eps) Then
         Y = h * Exp(-x)
         Exit Sub
      End If
   Next
   Y = "?"
    Exit Sub      'continued fraction failed '
Else 'Evaluate series.
   If (nm1 <> 0) Then 'Set rst term.
      Y = 1 / nm1
   Else
      Y = -Log(x) - EULER
   End If
   fact = 1
   For i = 1 To MAXIT
      fact = -fact * x / i
      If (i <> nm1) Then
         del = -fact / (i - nm1)
      Else
         Psi = -EULER '.
         For ii = 1 To nm1
            Psi = Psi + 1 / ii
         Next
         del = fact * (-Log(x) + Psi)
      End If
      Y = Y + del
      If (Abs(del) < Abs(Y) * eps) Then Exit Sub
   Next
   Y = "?"
   Exit Sub      'series failed in'
End If

End Sub

 Sub JY01A(ByVal x As Double, ByRef BJ0 As Double, ByRef DJ0 As Double, ByRef BJ1 As Double, ByRef DJ1 As Double, ByRef BY0 As Double, ByRef DY0 As Double, ByRef BY1 As Double, ByRef DY1 As Double)
'=======================================================
' Purpose: Compute Bessel functions J0(x), J1(x), Y0(x),
'         Y1(x), and their derivatives
' Input :  x   --- Argument of Jn(x) & Yn(x) ( x ò 0 )
' Output:  BJ0 --- J0(x)
'          DJ0 --- J0'(x)
'          BJ1 --- J1(x)
'          DJ1 --- J1'(x)
'          BY0 --- Y0(x)
'          DY0 --- Y0'(x)
'          BY1 --- Y1(x)
'          DY1 --- Y1'(x)
'=======================================================
Dim rp2 As Double, x2 As Double, r As Double, k As Long, EC As Double, CS0 As Double, CS1 As Double, W0 As Double, W1 As Double, R0 As Double, R1 As Double, A0() As Double, B0() As Double, A1() As Double, b1() As Double
Dim K0 As Double, T1 As Double, P0 As Double, p1 As Double, q0 As Double, q1 As Double, i As Long, CU As Double, T2 As Double
rp2 = 0.63661977236758
x2 = x * x
If (x = 0) Then
   BJ0 = 1
   BJ1 = 0
   DJ0 = 0
   DJ1 = 0.5
   BY0 = -1E+300
   BY1 = -1E+300
   DY0 = 1E+300
   DY1 = 1E+300
   Return
End If
If (x <= 12) Then
   BJ0 = 1
   r = 1
   For k = 1 To 30
      r = -0.25 * r * x2 / (k * k)
      BJ0 = BJ0 + r
      If (Abs(r) < Abs(BJ0) * 0.000000000000001) Then Exit For
   Next
   BJ1 = 1
   r = 1
   For k = 1 To 30
      r = -0.25 * r * x2 / (k * (k + 1))
      BJ1 = BJ1 + r
      If (Abs(r) < Abs(BJ1) * 0.000000000000001) Then Exit For
   Next
   BJ1 = 0.5 * x * BJ1
   EC = Log(x / 2) + 0.577215664901533
   CS0 = 0
   W0 = 0
   R0 = 1
   For k = 1 To 30
      W0 = W0 + 1 / k
      R0 = -0.25 * R0 / (k * k) * x2
      r = R0 * W0
      CS0 = CS0 + r
      If (Abs(r) < Abs(CS0) * 0.000000000000001) Then Exit For
   Next
   BY0 = rp2 * (EC * BJ0 - CS0)
   CS1 = 1
   W1 = 0
   R1 = 1
   For k = 1 To 30
      W1 = W1 + 1 / k
      R1 = -0.25 * R1 / (k * (k + 1)) * x2
      r = R1 * (2 * W1 + 1 / (k + 1))
      CS1 = CS1 + r
      If (Abs(r) < Abs(CS1) * 0.000000000000001) Then Exit For
   Next
   BY1 = rp2 * (EC * BJ1 - 1 / x - 0.25 * x * CS1)
Else
    A0 = Array(-0.0703125, 0.112152099609375, _
         -0.572501420974731, 6.07404200127348, _
         -110.017140269247, 3038.09051092238, _
         -118838.426256783, 6252951.4934348, _
         -425939216.504767, 36468400807.0656, _
         -3833534661393.94, 485401468685290#)
   B0 = Array(0.0732421875, -0.227108001708984, _
          1.72772750258446, -24.3805296995561, _
          551.335896122021, -18257.7554742932, _
          832859.304016289, -50069589.5319889, _
          3836255180.23043, -364901081884.983, _
          42189715702841#, -5.82724463156691E+15)
   A1 = Array(0.1171875, -0.144195556640625, _
          0.676592588424683, -6.88391426810995, _
          121.597891876536, -3302.27229448085, _
          127641.272646175, -6656367.71881769, _
          450278600.305039, -38338575207.4279, _
          4011838599133.2, -506056850331473#)
   b1 = Array(-0.1025390625, 0.277576446533203, _
          -1.9935317337513, 27.2488273112685, _
          -603.84407670507, 19718.3759122366, _
          -890297.876707068, 53104110.1096852, _
          -4043620325.10775, 382701134659.86, _
          -44064814178522.8, 6.0650913512227E+15)
   K0 = 12
   If (x >= 35) Then K0 = 10
   If (x >= 50) Then K0 = 8
   T1 = x - 0.25 * PI_
   P0 = 1
   q0 = -0.125 / x
   For k = 1 To K0
    i = k - 1
    P0 = P0 + A0(i) * x ^ (-2 * k)
    q0 = q0 + B0(i) * x ^ (-2 * k - 1)
   Next
   CU = Sqr(rp2 / x)
   BJ0 = CU * (P0 * Cos(T1) - q0 * Sin(T1))
   BY0 = CU * (P0 * Sin(T1) + q0 * Cos(T1))
   T2 = x - 0.75 * PI_
   p1 = 1
   q1 = 0.375 / x
   For k = 1 To K0
      i = k - 1
      p1 = p1 + A1(i) * x ^ (-2 * k)
      q1 = q1 + b1(i) * x ^ (-2 * k - 1)
   Next
   CU = Sqr(rp2 / x)
   BJ1 = CU * (p1 * Cos(T2) - q1 * Sin(T2))
   BY1 = CU * (p1 * Sin(T2) + q1 * Cos(T2))
End If
DJ0 = -BJ1
DJ1 = BJ0 - BJ1 / x
DY0 = -BY1
DY1 = BY0 - BY1 / x
End Sub


 Sub JYNA(ByVal n As Double, ByVal x As Double, ByRef NM As Double, ByRef BJ() As Double, ByRef DJ() As Double, ByRef BY() As Double, ByRef DY() As Double)
'  ==========================================================
'       Purpose: Compute Bessel functions Jn(x) & Yn(x) and
'                their derivatives
'       Input :  x --- Argument of Jn(x) & Yn(x)  ( x > 0 )
'                n --- Order of Jn(x) & Yn(x)
'       Output:  BJ(n) --- Jn(x)
'                DJ(n) --- Jn'(x)
'                BY(n) --- Yn(x)
'                DY(n) --- Yn'(x)
'                NM --- Highest order computed
'       Routines called:
'            (1) JY01A to calculate J0(x), J1(x), Y0(x) & Y1(x)
'            (2) MSTA1 and MSTA2 to calculate the starting
'                point for backward recurrence
'  =========================================================
Dim k As Double, BJ0 As Double, DJ0 As Double, BJ1 As Double, DJ1 As Double, BY0 As Double, DY0 As Double, BY1 As Double, DY1 As Double, BJK As Double, m As Double, F2 As Double, F1 As Double, f As Double, F0 As Double, CS As Double
ReDim BJ(n) As Double, BY(n) As Double, DJ(n) As Double, DY(n) As Double

    NM = n
    If (x < 1E-100) Then
       For k = 0 To n
          BJ(k) = 0
          DJ(k) = 0
          BY(k) = -1E+300
          DY(k) = 1E+300
       Next
       BJ(0) = 1
       DJ(1) = 0.5
       Exit Sub
    End If
    Call JY01A(x, BJ0, DJ0, BJ1, DJ1, BY0, DY0, BY1, DY1)
    BJ(0) = BJ0
    BJ(1) = BJ1
    BY(0) = BY0
    BY(1) = BY1
    DJ(0) = DJ0
    DJ(1) = DJ1
    DY(0) = DY0
    DY(1) = DY1
    If (n <= 1) Then Exit Sub
    If (n < Int(0.9 * x)) Then
       For k = 2 To n
          BJK = 2 * (k - 1) / x * BJ1 - BJ0
          BJ(k) = BJK
          BJ0 = BJ1
          BJ1 = BJK
      Next
    Else
       m = MSTA1(x, 200)
       If (m < n) Then
          NM = m
       Else
          m = MSTA2(x, n, 15)
       End If
       F2 = 0
       F1 = 1E-100
       For k = m To 0 Step -1
          f = 2 * (k + 1) / x * F1 - F2
          If (k <= NM) Then BJ(k) = f
          F2 = F1
          F1 = f
       Next
        If (Abs(BJ0) > Abs(BJ1)) Then
           CS = BJ0 / f
        Else
           CS = BJ1 / F2
        End If
        For k = 0 To NM
            BJ(k) = CS * BJ(k)
        Next
    End If
    
    For k = 2 To NM
       DJ(k) = BJ(k - 1) - k / x * BJ(k)
    Next
    F0 = BY(0)
    F1 = BY(1)
    For k = 2 To NM
       f = 2 * (k - 1) / x * F1 - F0
       BY(k) = f
       F0 = F1
       F1 = f
    Next
    For k = 2 To NM
       DY(k) = BY(k - 1) - k * BY(k) / x
    Next
End Sub


Private Function MSTA1(ByVal x As Double, ByVal mp As Double) As Integer
'  ===================================================
'  Purpose: Determine the starting point for backward
'           recurrence such that the magnitude of
'           Jn(x) at that point is about 10^(-MP)
'  Input :  x     --- Argument of Jn(x)
'           MP    --- Value of magnitude
'  Output:  MSTA1 --- Starting point
' ===================================================
Dim A0 As Double, n0 As Double, F0 As Double, N1 As Double, F1 As Double, IT As Long, nn As Double, f As Double
A0 = Abs(x)
n0 = Int(1.1 * A0) + 1
F0 = ENVJ(n0, A0) - mp
N1 = n0 + 5
F1 = ENVJ(N1, A0) - mp
For IT = 1 To 20
   nn = N1 - (N1 - n0) / (1 - F0 / F1)
   f = ENVJ(nn, A0) - mp
   If (Abs(nn - N1) < 1) Then Exit For
   n0 = N1
   F0 = F1
   N1 = nn
   F1 = f
Next
MSTA1 = nn
End Function


Private Function MSTA2(ByVal x As Double, ByVal n As Double, ByVal mp As Double) As Integer
' ===================================================
' Purpose: Determine the starting point for backward
'         recurrence such that all Jn(x) has MP
'         significant digits
' Input :  x  --- Argument of Jn(x)
'          n  --- Order of Jn(x)
'          MP --- Significant digit
' Output:  MSTA2 --- Starting point
' ===================================================
Dim A0 As Double, HMP As Double, EJN As Double, OBJ As Double, n0 As Double, F0 As Double, N1 As Double, F1 As Double, IT As Long, nn As Double, f As Double
A0 = Abs(x)
HMP = 0.5 * mp
EJN = ENVJ(n, A0)
If (EJN <= HMP) Then
   OBJ = mp
   n0 = Int(1.1 * A0) + 1 'bug for x<0.1 - VL, 2-8.2002
Else
   OBJ = HMP + EJN
   n0 = n
End If
F0 = ENVJ(n0, A0) - OBJ
N1 = n0 + 5
F1 = ENVJ(N1, A0) - OBJ
For IT = 1 To 20
   nn = N1 - (N1 - n0) / (1 - F0 / F1)
   f = ENVJ(nn, A0) - OBJ
   If (Abs(nn - N1) < 1) Then Exit For
   n0 = N1
   F0 = F1
   N1 = nn
   F1 = f
Next
MSTA2 = nn + 10
End Function

Private Function ENVJ(ByVal n As Double, ByVal x As Double) As Double
ENVJ = 0.5 * Log10(6.28 * n) - n * Log10(1.36 * x / n)
End Function

Private Function Log10(ByVal x As Double) As Double
Log10 = Log(x) / Log(10)
End Function


 Sub IK01A(ByVal x As Double, ByRef BI0 As Double, ByRef DI0 As Double, ByRef BI1 As Double, ByRef DI1 As Double, ByRef BK0 As Double, ByRef DK0 As Double, ByRef BK1 As Double, ByRef DK1 As Double)
'=========================================================
'Purpose: Compute modified Bessel functions I0(x), I1(1),
'         K0(x) and K1(x), and their derivatives
'Input :  x   --- Argument ( x ò 0 )
'Output:  BI0 --- I0(x)
'         DI0 --- I0'(x)
'         BI1 --- I1(x)
'         DI1 --- I1'(x)
'         BK0 --- K0(x)
'         DK0 --- K0'(x)
'         BK1 --- K1(x)
'         DK1 --- K1'(x)
'=========================================================
 Const EL As Double = 0.577215664901533
 Dim x2 As Double, r As Double, i As Long, k As Long, A0() As Double, B0() As Double, K0 As Double, CA As Double, XR As Double, CT As Double, W0 As Double, WW As Double, A1() As Double, CB As Double, XR2 As Double
 x2 = x * x
 If (x = 0) Then
    BI0 = 1
    BI1 = 0
    BK0 = 1E+300
    BK1 = 1E+300
    DI0 = 0
    DI1 = 0.5
    DK0 = -1E+300
    DK1 = -1E+300
    Exit Sub
 ElseIf (x <= 18) Then
    BI0 = 1
    r = 1
    For k = 1 To 50
       r = 0.25 * r * x2 / (k * k)
       BI0 = BI0 + r
       If (Abs(r / BI0) < 0.000000000000001) Then Exit For
    Next
    BI1 = 1
    r = 1
    For k = 1 To 50
       r = 0.25 * r * x2 / (k * (k + 1))
       BI1 = BI1 + r
       If (Abs(r / BI1) < 0.000000000000001) Then Exit For
    Next
    BI1 = 0.5 * x * BI1
 Else
    A0 = Array(0.125, 0.0703125, _
          0.0732421875, 0.11215209960938, _
          0.22710800170898, 0.57250142097473, _
          1.7277275025845, 6.0740420012735, _
          24.380529699556, 110.01714026925, _
          551.33589612202, 3038.0905109224)
    B0 = Array(-0.375, -0.1171875, _
          -0.1025390625, -0.14419555664063, _
          -0.2775764465332, -0.67659258842468, _
          -1.9935317337513, -6.8839142681099, _
          -27.248827311269, -121.59789187654, _
          -603.84407670507, -3302.2722944809)
    K0 = 12
    If (x >= 35) Then K0 = 9
    If (x >= 50) Then K0 = 7
    CA = Exp(x) / Sqr(2 * PI_ * x)
    BI0 = 1
    XR = 1 / x
    For k = 1 To K0
        i = k - 1
       BI0 = BI0 + A0(i) * XR ^ k
    Next
    BI0 = CA * BI0
    BI1 = 1
    For k = 1 To K0
        i = k - 1
       BI1 = BI1 + B0(i) * XR ^ k
    Next
    BI1 = CA * BI1
 End If
 If (x <= 9) Then
    CT = -(Log(x / 2) + EL)
    BK0 = 0
    W0 = 0
    r = 1
    For k = 1 To 50
       W0 = W0 + 1 / k
       r = 0.25 * r / (k * k) * x2
       BK0 = BK0 + r * (W0 + CT)
       If (Abs((BK0 - WW) / BK0) < 0.000000000000001) Then Exit For
       WW = BK0
   Next
    BK0 = BK0 + CT
 Else
    A1 = Array(0.125, 0.2109375, _
           1.0986328125, 11.775970458984, _
           214.61706161499, 5951.1522710323, _
           233476.45606175, 12312234.987631)
    CB = 0.5 / x
    XR2 = 1 / x2
    BK0 = 1
    For k = 1 To 8
        i = k - 1
       BK0 = BK0 + A1(i) * XR2 ^ k
    Next
    BK0 = CB * BK0 / BI0
 End If
 BK1 = (1 / x - BI1 * BK0) / BI0
 DI0 = BI1
 DI1 = BI0 - BI1 / x
 DK0 = -BK1
 DK1 = -BK0 - BK1 / x

 End Sub
 
 Sub IKNA(ByVal n As Double, ByVal x As Double, ByRef NM As Double, ByRef BI() As Double, ByRef DI() As Double, ByRef BK() As Double, ByRef DK() As Double)
' ========================================================
' Purpose: Compute modified Bessel functions In(x) and
'          Kn(x), and their derivatives
' Input:   x --- Argument of In(x) and Kn(x) ( x ò 0 )
'          n --- Order of In(x) and Kn(x)
' Output:  BI(n) --- In(x)
'          DI(n) --- In'(x)
'          BK(n) --- Kn(x)
'          DK(n) --- Kn'(x)
'          NM --- Highest order computed
' Routines called:
'      (1) IK01A for computing I0(x),I1(x),K0(x) & K1(x)
'      (2) MSTA1 and MSTA2 for computing the starting
'          point for backward recurrence
' ========================================================
Dim k As Long, BI0 As Double, DI0 As Double, BI1 As Double, DI1 As Double, BK0 As Double, DK0 As Double, BK1 As Double, DK1 As Double
Dim H0 As Double, H1 As Double, h As Double, m As Double, F0 As Double, F1 As Double, f As Double, S0 As Double
Dim G0 As Double, G1 As Double, g As Double
ReDim BI(n) As Double, DI(n) As Double, BK(n) As Double, DK(n) As Double
NM = n
If (x <= 1E-100) Then
   For k = 0 To n
      BI(k) = 0
      DI(k) = 0
      BK(k) = 1E+300
      DK(k) = -1E+300
   Next
   BI(0) = 1
   DI(1) = 0.5
   Exit Sub
End If
Call IK01A(x, BI0, DI0, BI1, DI1, BK0, DK0, BK1, DK1)
BI(0) = BI0
BI(1) = BI1
BK(0) = BK0
BK(1) = BK1
DI(0) = DI0
DI(1) = DI1
DK(0) = DK0
DK(1) = DK1
If (n <= 1) Then Exit Sub
If (x > 40 And n < Int(0.25 * x)) Then
   H0 = BI0
   H1 = BI1
   For k = 2 To n
     h = -2 * (k - 1) / x * H1 + H0
     BI(k) = h
     H0 = H1
     H1 = h
   Next
Else
   m = MSTA1(x, 200)
   If (m < n) Then
      NM = m
   Else
      m = MSTA2(x, n, 15)
   End If
   F0 = 0
   F1 = 1E-100
   For k = m To 0 Step -1
      f = 2 * (k + 1) * F1 / x + F0
      If (k <= NM) Then BI(k) = f
      F0 = F1
      F1 = f
   Next
   S0 = BI0 / f
   For k = 0 To NM
      BI(k) = S0 * BI(k)
   Next
End If
G0 = BK0
G1 = BK1
For k = 2 To NM
   g = 2 * (k - 1) / x * G1 + G0
   BK(k) = g
   G0 = G1
   G1 = g
Next
For k = 2 To NM
   DI(k) = BI(k - 1) - k / x * BI(k)
   DK(k) = -BK(k - 1) - k / x * BK(k)
Next
End Sub

 Sub CISIA(ByVal x As Double, ByRef CI As Double, ByRef SI As Double)
'=============================================
' Purpose: Compute cosine and sine integrals
'          Si(x) and Ci(x)  ( x ò 0 )
' Input :  x  --- Argument of Ci(x) and Si(x)
' Output:  CI --- Ci(x)
'          SI --- Si(x)
'=============================================
Dim BJ(101) As Double, p2 As Double, EL As Double, eps As Double, x2 As Double, XR As Double, k As Long, m As Double
Dim XA0 As Double, XA1 As Double, Xa As Double, XS As Double, XG1 As Double, XG2 As Double
Dim XCS As Double, XSS As Double, XF As Double, XG As Double

p2 = PI_ / 2
EL = 0.577215664901533
eps = 0.000000000000001
x2 = x * x
If (x = 0) Then
   CI = -1E+300
   SI = 0
ElseIf (x <= 16) Then
   XR = -0.25 * x2
   CI = EL + Log(x) + XR
   For k = 2 To 40
      XR = -0.5 * XR * (k - 1) / (k * k * (2 * k - 1)) * x2
      CI = CI + XR
      If (Abs(XR) < Abs(CI) * eps) Then Exit For
   Next
   XR = x
   SI = x
   For k = 1 To 40
      XR = -0.5 * XR * (2 * k - 1) / k / (4 * k * k + 4 * k + 1) * x2
      SI = SI + XR
      If (Abs(XR) < Abs(SI) * eps) Then Exit For
   Next
ElseIf (x <= 32) Then
   m = Int(47.2 + 0.82 * x)
   XA1 = 0
   XA0 = 1E-100
   For k = m To 1 Step -1
      Xa = 4 * k * XA0 / x - XA1
      BJ(k) = Xa
      XA1 = XA0
      XA0 = Xa
   Next
   XS = BJ(1)
   For k = 3 To m Step 2
      XS = XS + 2 * BJ(k)
   Next
   BJ(1) = BJ(1) / XS
   For k = 2 To m
      BJ(k) = BJ(k) / XS
   Next
   XR = 1
   XG1 = BJ(1)
   For k = 2 To m
      XR = 0.25 * XR * (2 * k - 3) ^ 2 / ((k - 1) * (2 * k - 1) ^ 2) * x
      XG1 = XG1 + BJ(k) * XR
   Next
   XR = 1
   XG2 = BJ(1)
   For k = 2 To m
      XR = 0.25 * XR * (2 * k - 5) ^ 2 / ((k - 1) * (2 * k - 3) ^ 2) * x
      XG2 = XG2 + BJ(k) * XR
   Next
   XCS = Cos(x / 2)
   XSS = Sin(x / 2)
   CI = EL + Log(x) - x * XSS * XG1 + 2 * XCS * XG2 - 2 * XCS * XCS
   SI = x * XCS * XG1 + 2 * XSS * XG2 - Sin(x)
Else
   XR = 1
   XF = 1
   For k = 1 To 9
      XR = -2 * XR * k * (2 * k - 1) / x2
      XF = XF + XR
   Next
   XR = 1 / x
   XG = XR
   For k = 1 To 8
      XR = -2 * XR * (2 * k + 1) * k / x2
      XG = XG + XR
   Next
   CI = XF * Sin(x) / x - XG * Cos(x) / x
   SI = p2 - XF * Cos(x) / x - XG * Sin(x) / x
End If
End Sub

 Sub FCS(ByVal x As Double, ByRef c As Double, ByRef s As Double)
' =================================================
'  Purpose: Compute Fresnel integrals C(x) and S(x)
'  Input :  x --- Argument of C(x) and S(x)
'  Output:  C --- C(x)
'           S --- S(x)
' =================================================
   Const eps As Double = 0.000000000000001
   Dim Xa As Double, PX As Double, t As Double, T0 As Double, T1 As Double, T2 As Double, r As Double, k As Long, m As Double, SU As Double, f As Double, F0 As Double, F1 As Double, q As Double, g As Double
   
   Xa = Abs(x)
   PX = PI_ * Xa
   t = 0.5 * PX * Xa
   T2 = t * t
   If (Xa = 0) Then
      c = 0
      s = 0
   ElseIf (Xa < 2.5) Then
      r = Xa
      c = r
      For k = 1 To 50
         r = -0.5 * r * (4 * k - 3) / k / (2 * k - 1) / (4 * k + 1) * T2
         c = c + r
         If (Abs(r) < Abs(c) * eps) Then Exit For
      Next
      s = Xa * t / 3
      r = s
      For k = 1 To 50
         r = -0.5 * r * (4 * k - 1) / k / (2 * k + 1) / (4 * k + 3) * T2
         s = s + r
         If (Abs(r) < Abs(s) * eps) Then GoTo Label40
      Next
   ElseIf (Xa < 4.5) Then
      m = Int(42 + 1.75 * t)
      SU = 0
      c = 0
      s = 0
      F1 = 0
      F0 = 1E-100
      For k = m To 0 Step -1
         f = (2 * k + 3) * F0 / t - F1
         If (k = Int(k / 2) * 2) Then
            c = c + f
         Else
            s = s + f
         End If
         SU = SU + (2 * k + 1) * f * f
         F1 = F0
         F0 = f
      Next
      q = Sqr(SU)
      c = c * Xa / q
      s = s * Xa / q
   Else
      r = 1
      f = 1
      For k = 1 To 20
         r = -0.25 * r * (4 * k - 1) * (4 * k - 3) / T2
         f = f + r
      Next
      r = 1 / (PX * Xa)
      g = r
      For k = 1 To 12
         r = -0.25 * r * (4 * k + 1) * (4 * k - 1) / T2
         g = g + r
      Next
      T0 = t - Int(t / (2 * PI_)) * 2 * PI_
      c = 0.5 + (f * Sin(T0) - g * Cos(T0)) / PX
      s = 0.5 - (f * Cos(T0) + g * Sin(T0)) / PX
   End If
Exit Sub
Label40:
If (x < 0) Then
   c = -c
   s = -s
End If

End Sub

Sub HYGFX(ByVal A As Double, ByVal B As Double, ByVal c As Double, ByVal x As Double, ByRef hf As Double, ByRef ErrorMsg As String)
' ====================================================
'       Purpose: Compute hypergeometric function F(a,b,c,x)
'       Input :  a --- Parameter
'                b --- Parameter
'                c --- Parameter, c <> 0,-1,-2,...
'                x --- Argument   ( x < 1 )
'       Output:  HF --- F(a,b,c,x)
'====================================================
Dim L0 As Boolean, L1 As Boolean, L2 As Boolean, L3 As Boolean, L4 As Boolean, L5 As Boolean
Dim EL As Double, eps As Double, GC As Double, GCAB As Double, GCA As Double, GCB As Double, G0 As Double, G1 As Double, G2 As Double
Dim G3 As Double, NM As Double, r As Double, j As Long, k As Long, AA As Double, BB As Double, x1 As Double, GM As Double, m As Double, GA As Double, GB As Double
Dim GAM As Double, GBM As Double, PA As Double, PB As Double, RM As Double, F0 As Double, R0 As Double, R1 As Double, SP0 As Double, SP As Double, C0 As Double
Dim C1 As Double, F1 As Double, SM As Double, RP As Double, HW As Double, GABC As Double, A0 As Double

    EL = 0.577215664901533
    eps = 0.000000000000001
    L0 = (c = Int(c)) And (c < 0)
    L1 = ((1 - x) < eps) And ((c - A - B) <= 0)
    L2 = (A = Int(A)) And (A < 0)
    L3 = (B = Int(B)) And (B < 0)
    L4 = (c - A = Int(c - A)) And (c - A <= 0)
    L5 = (c - B = Int(c - B)) And (c - B <= 0)
    If (L0 Or L1) Then
       ErrorMsg = "The hypergeometric series is divergent"
       Exit Sub
    End If
    If (x > 0.95) Then eps = 0.00000001
    If (x = 0 Or A = 0 Or B = 0) Then
       hf = 1
       Exit Sub
    ElseIf ((1 - x = eps) And (c - A - B) > 0) Then
       Call HGamma(c, GC)
       Call HGamma(c - A \ -B, GCAB)
       Call HGamma(c - A, GCA)
       Call HGamma(c - B, GCB)
       hf = GC * GCAB / (GCA * GCB)
       Exit Sub
    ElseIf ((1 + x <= eps) And (Abs(c - A + B - 1) <= eps)) Then
       G0 = Sqr(PI_) * 2 ^ (-A)
       Call HGamma(c, G1)
       Call HGamma(1 + A / 2 - B, G2)
       Call HGamma(0.5 + 0.5 * A, G3)
       hf = G0 * G1 / (G2 * G3)
       Exit Sub
    ElseIf (L2 Or L3) Then
       If (L2) Then NM = Int(Abs(A))
       If (L3) Then NM = Int(Abs(B))
       hf = 1
       r = 1
       For k = 1 To NM
          r = r * (A + k - 1) * (B + k - 1) / (k * (c + k - 1)) * x
          hf = hf + r
       Next k
       Exit Sub
    ElseIf (L4 Or L5) Then
       If (L4) Then NM = Int(Abs(c - A))
       If (L5) Then NM = Int(Abs(c - B))
       hf = 1
       r = 1
       For k = 1 To NM
          r = r * (c - A + k - 1) * (c - B + k - 1) / (k * (c + k - 1)) * x
          hf = hf + r
       Next k
       hf = (1 - x) ^ (c - A - B) * hf
       Exit Sub
    End If
    AA = A
    BB = B
    x1 = x
    If (x < 0) Then
       x = x / (x - 1)
       If (c > A And B < A And B > 0) Then
          A = BB
          B = AA
       End If
       B = c - B
    End If
    If (x >= 0.75) Then
       GM = 0
       If (Abs(c - A - B - Int(c - A - B)) < 0.000000000000001) Then
          m = Int(c - A - B)
          Call HGamma(A, GA)
          Call HGamma(B, GB)
          Call HGamma(c, GC)
          Call HGamma(A + m, GAM)
          Call HGamma(B + m, GBM)
          Call HDigamma(A, PA)
          Call HDigamma(B, PB)
          If (m <> 0) Then GM = 1
          For j = 1 To Abs(m) - 1
             GM = GM * j
          Next j
          RM = 1
          For j = 1 To Abs(m)
             RM = RM * j
          Next j
          F0 = 1
          R0 = 1
          R1 = 1
          SP0 = 0
          SP = 0
          If (m >= 0) Then
             C0 = GM * GC / (GAM * GBM)
             C1 = -GC * (x - 1) ^ m / (GA * GB * RM)
             For k = 1 To m - 1
                R0 = R0 * (A + k - 1) * (B + k - 1) / (k * (k - m)) * (1 - x)
                F0 = F0 + R0
             Next k
             For k = 1 To m
                SP0 = SP0 + 1 / (A + k - 1) + 1 / (B + k - 1) - 1 / k
             Next k
             F1 = PA + PB + SP0 + 2 * EL + Log(1 - x)
             For k = 1 To 250
                SP = SP + (1 - A) / (k * (A + k - 1)) + (1 - B) / (k * (B + k - 1))
                SM = 0
                For j = 1 To m
                   SM = SM + (1 - A) / ((j + k) * (A + j + k - 1)) + 1 / (B + j + k - 1)
                Next j
                RP = PA + PB + 2 * EL + SP + SM + Log(1 - x)
                R1 = R1 * (A + m + k - 1) * (B + m + k - 1) / (k * (m + k)) * (1 - x)
                F1 = F1 + R1 * RP
                If (Abs(F1 - HW) < Abs(F1) * eps) Then GoTo 60
                HW = F1
             Next k
60:              hf = F0 * C0 + F1 * C1
          ElseIf (m < 0) Then
             m = -m
             C0 = GM * GC / (GA * GB * (1 - x) ^ m)
             C1 = -(-1) ^ m * GC / (GAM * GBM * RM)
             For k = 1 To m - 1
                R0 = R0 * (A - m + k - 1) * (B - m + k - 1) / (k * (k - m)) * (1 - x)
                F0 = F0 + R0
             Next k
             For k = 1 To m
                SP0 = SP0 + 1 / k
             Next k
             F1 = PA + PB - SP0 + 2 * EL + Log(1 - x)
             For k = 1 To 250
                SP = SP + (1 - A) / (k * (A + k - 1)) + (1 - B) / (k * (B + k - 1))
                SM = 0
                For j = 1 To m
                   SM = SM + 1 / (j + k)
                Next j
                RP = PA + PB + 2 * EL + SP - SM + Log(1 - x)
                R1 = R1 * (A + k - 1) * (B + k - 1) / (k * (m + k)) * (1 - x)
                F1 = F1 + R1 * RP
                If (Abs(F1 - HW) < (Abs(F1) * eps)) Then GoTo 85
                HW = F1
             Next k
85:              hf = F0 * C0 + F1 * C1
          End If
       Else
            Call HGamma(A, GA)
            Call HGamma(B, GB)
            Call HGamma(c, GC)
            Call HGamma(c - A, GCA)
            Call HGamma(c - B, GCB)
            Call HGamma(c - A - B, GCAB)
            Call HGamma(A + B - c, GABC)
          C0 = GC * GCAB / (GCA * GCB)
          C1 = GC * GABC / (GA * GB) * (1 - x) ^ (c - A - B)
          hf = 0
          R0 = C0
          R1 = C1
          For k = 1 To 250
             R0 = R0 * (A + k - 1) * (B + k - 1) / (k * (A + B - c + k)) * (1 - x)
             R1 = R1 * (c - A + k - 1) * (c - B + k - 1) / (k * (c - A - B + k)) * (1 - x)
             hf = hf + R0 + R1
             If (Abs(hf - HW) < (Abs(hf) * eps)) Then GoTo 95
             HW = hf
          Next k
95:           hf = hf + C0 + C1
       End If
    Else
       A0 = 1
       If ((c > A) And (c < (2 * A)) And (c > B) And (c < 2 * B)) Then
          A0 = (1 - x) ^ (c - A - B)
          A = c - A
          B = c - B
       End If
       hf = 1
       r = 1
       For k = 1 To 250
          r = r * (A + k - 1) * (B + k - 1) / (k * (c + k - 1)) * x
          hf = hf + r
          If (Abs(hf - HW) <= (Abs(hf) * eps)) Then GoTo 105
          HW = hf
       Next k
105:       hf = A0 * hf
    End If
    If (x1 < 0) Then
       x = x1
       C0 = 1 / (1 - x) ^ AA
       hf = C0 * hf
    End If
    A = AA
    B = BB
    If (k > 120) Then
       ErrorMsg = "Warning! You should check the accuracy"
       Exit Sub
    End If
End Sub

 Sub INCOG(ByVal A As Double, ByVal x As Double, ByRef GIN As Double, ByRef GIM As Double, ByRef GIP As Double, ByRef MSG As String)
' ===================================================
'       Purpose: Compute the incomplete gamma function
'        c R(a, x), â(a, x) And P(a, x)
'       Input :  a   --- Parameter ( a < 170 )
'                x   - --Argument
'       Output:        GIN ---R(a, x)
'                      GIM - --â(a, x)
'                      GIP - --P(a, x)
'       Routine called: GAMMA for computing â(x)
'===================================================
Dim k As Long, XAM As Double, GA As Double, s As Double, r As Double, T0 As Double
        XAM = -x + A * Log(x)
        If (XAM > 700 Or A > 170) Then
           MSG = "a and/or x too large"
           Exit Sub
        End If
        If (x = 0) Then
           GIN = 0
           Call HGamma(A, GA)
           GIM = GA
           GIP = 0
        ElseIf (x <= 1 + A) Then
           s = 1 / A
           r = s
            For k = 1 To 60
              r = r * x / (A + k)
              s = s + r
              If (Abs(r / s) < 10 ^ -15) Then Exit For
            Next k
           GIN = Exp(XAM) * s
           Call HGamma(A, GA)
           GIP = GIN / GA
           GIM = GA - GIN
        ElseIf (x > 1 + A) Then
           T0 = 0
           For k = 60 To 1 Step -1
              T0 = (k - A) / (1 + k / (x + T0))
           Next k
           GIM = Exp(XAM) / (x + T0)
           Call HGamma(A, GA)
           GIN = GA - GIM
           GIP = 1 - GIM / GA
        End If
End Sub

 Sub INCOB(ByVal A As Double, ByVal B As Double, ByVal x As Double, ByRef BIX As Double)
' ========================================================
'      Purpose: Compute the incomplete beta function Ix(a,b)
'       Input :  a --- Parameter
'                b - --Parameter
'                x --- Argument ( 0 ó x ó 1 )
'       Output:        BIX ---Ix(a, b)
'       Routine called: BETA for computing beta function B(p,q)
' ========================================================
Dim DK(51) As Double, FK(51) As Double, k As Long, S0 As Double, T1 As Double, T2 As Double, TA As Double, TB As Double, BT As Double
    S0 = (A + 1) / (A + B + 2)
    Call HBeta(A, B, BT)
    If (x <= S0) Then
       For k = 1 To 20
          DK(2 * k) = k * (B - k) * x / (A + 2 * k - 1) / (A + 2 * k)
       Next k
       For k = 0 To 20
          DK(2 * k + 1) = -(A + k) * (A + B + k) * x / (A + 2 * k) / (A + 2 * k + 1)
       Next k
       T1 = 0
       For k = 20 To 1 Step -1
          T1 = DK(k) / (1 + T1)
       Next k
       TA = 1 / (1 + T1)
       BIX = x ^ A * (1 - x) ^ B / (A * BT) * TA
    Else
       For k = 1 To 20
          FK(2 * k) = k * (A - k) * (1 - x) / (B + 2 * k - 1) / (B + 2 * k)
       Next k
       For k = 0 To 20
          FK(2 * k + 1) = -(B + k) * (A + B + k) * (1 - x) / (B + 2 * k) / (B + 2 * k + 1)
       Next k
       T2 = 0
       For k = 20 To 1 Step -1
          T2 = FK(k) / (1 + T2)
       Next k
       TB = 1 / (1 + T2)
       BIX = 1 - x ^ A * (1 - x) ^ B / (B * BT) * TB
    End If
End Sub

Sub AIRYB(ByVal x As Double, ByRef AI As Double, ByRef BI As Double, ByRef AD As Double, ByRef BD As Double)
'=======================================================
'       Purpose: Compute Airy functions and their derivatives
'       Input:   x  --- Argument of Airy function
'       Output:  AI --- Ai(x)
'                BI --- Bi(x)
'                AD --- Ai'(x)
'                BD --- Bi'(x)
'=======================================================
   Dim CK(41) As Double, DK(41) As Double
   Dim eps As Double, C1 As Double, C2 As Double, SR3 As Double, Xa As Double, XQ As Double, XM As Double, FX As Double, r As Double, GX As Double, df As Double, DG As Double
   Dim XE As Double, XR1 As Double, XAR As Double, XF As Double, RP As Double, KM As Double
   Dim SAI As Double, SAD As Double, SBI As Double, SBD As Double, XP1 As Double, XCS As Double, XSS As Double, SSA As Double, SDA As Double, XR2 As Double, SSB As Double, SDB As Double
   Dim k As Long
   
   eps = 0.000000000000001
   C1 = 0.355028053887817
   C2 = 0.258819403792807
   SR3 = 1.73205080756888
   Xa = Abs(x)
   XQ = Sqr(Xa)
   If (x > 0) Then XM = 5
   If (x <= 0) Then XM = 8
   If (x = 0) Then
      AI = C1
      BI = SR3 * C1
      AD = -C2
      BD = SR3 * C2
      Exit Sub
   End If
   If (Xa <= XM) Then
      FX = 1
      r = 1
      For k = 1 To 40
         r = r * x / (3 * k) * x / (3 * k - 1) * x
         FX = FX + r
         If (Abs(r) < Abs(FX) * eps) Then Exit For
      Next k
      GX = x
      r = x
      For k = 1 To 40
         r = r * x / (3 * k) * x / (3 * k + 1) * x
         GX = GX + r
         If (Abs(r) < Abs(GX) * eps) Then Exit For
      Next k
      AI = C1 * FX - C2 * GX
      BI = SR3 * (C1 * FX + C2 * GX)
      df = 0.5 * x * x
      r = df
      For k = 1 To 40
         r = r * x / (3 * k) * x / (3 * k + 2) * x
         df = df + r
         If (Abs(r) < Abs(df) * eps) Then Exit For
      Next k
      DG = 1
      r = 1
      For k = 1 To 40
         r = r * x / (3 * k) * x / (3 * k - 2) * x
         DG = DG + r
         If (Abs(r) < Abs(DG) * eps) Then Exit For
      Next k
      AD = C1 * df - C2 * DG
      BD = SR3 * (C1 * df + C2 * DG)
   Else
      XE = Xa * XQ / 1.5
      XR1 = 1 / XE
      XAR = 1 / XQ
      XF = Sqr(XAR)
      RP = 0.564189583547756
      r = 1
      For k = 1 To 40
         r = r * (6 * k - 1) / 216 * (6 * k - 3) / k * (6 * k - 5) / (2 * k - 1)
         CK(k) = r
         DK(k) = -(6 * k + 1) / (6 * k - 1) * CK(k)
      Next k
      KM = Int(24.5 - Xa)
      If (Xa < 6) Then KM = 14
      If (Xa > 15) Then KM = 10
      If (x > 0) Then
         SAI = 1
         SAD = 1
         r = 1
         For k = 1 To KM
            r = -r * XR1
            SAI = SAI + CK(k) * r
            SAD = SAD + DK(k) * r
         Next k
         SBI = 1
         SBD = 1
         r = 1
         For k = 1 To KM
            r = r * XR1
            SBI = SBI + CK(k) * r
            SBD = SBD + DK(k) * r
         Next k
         XP1 = Exp(-XE)
         AI = 0.5 * RP * XF * XP1 * SAI
         BI = RP * XF / XP1 * SBI
         AD = -0.5 * RP / XF * XP1 * SAD
         BD = RP / XF / XP1 * SBD
      Else
         XCS = Cos(XE + PI_ / 4)
         XSS = Sin(XE + PI_ / 4)
         SSA = 1
         SDA = 1
         r = 1
         XR2 = 1 / (XE * XE)
         For k = 1 To KM
            r = -r * XR2
            SSA = SSA + CK(2 * k) * r
            SDA = SDA + DK(2 * k) * r
         Next k
         SSB = CK(1) * XR1
         SDB = DK(1) * XR1
         r = XR1
         For k = 1 To KM
            r = -r * XR2
            SSB = SSB + CK(2 * k + 1) * r
            SDB = SDB + DK(2 * k + 1) * r
         Next k
         AI = RP * XF * (XSS * SSA - XCS * SSB)
         BI = RP * XF * (XCS * SSA + XSS * SSB)
         AD = -RP / XF * (XCS * SDA + XSS * SDB)
         BD = RP / XF * (XSS * SDA - XCS * SDB)
      End If
   End If
        
End Sub


Sub ELIT(ByVal HK As Double, ByVal phi As Double, ByRef FE As Double, ByRef EE As Double)
' ==================================================
'       Purpose: Compute complete and incomplete elliptic
'                integrals F(k,phi) and E(k,phi)
'       Input  : HK  --- Modulus k ( 0 ó k ó 1 )
'                Phi --- Argument ( in degrees )
'       Output : FE  --- F(k,phi)
'                EE  --- E(k,phi)
' ==================================================
Dim G1 As Double, A0 As Double, B0 As Double, A1 As Double, b1 As Double, C1 As Double, D0 As Double, D1 As Double, r As Double, FAC As Double, CK As Double, CE As Double
Dim n As Long

    G1 = 0
    A0 = 1
    B0 = Sqr(1 - HK * HK)
    D0 = (PI_ / 180) * phi
    r = HK * HK
    If (HK = 1 And phi = 90) Then
       FE = 1E+300
       EE = 1
    ElseIf (HK = 1) Then
       FE = Log((1 + Sin(D0)) / Cos(D0))
       EE = Sin(D0)
    Else
       FAC = 1
       For n = 1 To 40
          A1 = (A0 + B0) / 2
          b1 = Sqr(A0 * B0)
          C1 = (A0 - B0) / 2
          FAC = 2 * FAC
          r = r + FAC * C1 * C1
          If (phi <> 90) Then
             D1 = D0 + Atn((B0 / A0) * Tan(D0))
             G1 = G1 + C1 * Sin(D1)
             D0 = D1 + PI_ * Int(D1 / PI_ + 0.5)
          End If
          A0 = A1
          B0 = b1
          If (C1 < 0.0000001) Then Exit For
       Next n
       CK = PI_ / (2 * A1)
       CE = PI_ * (2 - r) / (4 * A1)
       If (phi = 90) Then
          FE = CK
          EE = CE
       Else
          FE = D1 / (FAC * A1)
          EE = FE * CE / CK + G1
       End If
    End If
End Sub


'-------------------------------------------------------------------------------
' Legendre's polynomials
'-------------------------------------------------------------------------------
Sub PLegendre(ByVal x As Double, ByVal n As Double, ByRef Y As Double)
Dim i As Long, P0 As Double, p1 As Double, p2 As Double
P0 = 0: p1 = 1: p2 = p1
For i = 1 To n
    p2 = (2 * i - 1) / i * x * p1 - (i - 1) / i * P0
    P0 = p1
    p1 = p2
Next i
Y = p2
End Sub

'-------------------------------------------------------------------------------
' Hermite's polynomials
'-------------------------------------------------------------------------------
Sub PHermite(ByVal x As Double, ByVal n As Double, ByRef Y As Double)
Dim i As Long, P0, p1, p2
P0 = 0: p1 = 1: p2 = p1
For i = 1 To n
    p2 = 2 * x * p1 - 2 * (i - 1) * P0
    P0 = p1
    p1 = p2
Next i
Y = p2
End Sub

'-------------------------------------------------------------------------------
' Laguerre's polynomials
'-------------------------------------------------------------------------------
Sub PLaguerre(ByVal x As Double, ByVal n As Double, ByRef Y As Double)
Dim i As Long, P0 As Double, p1 As Double, p2 As Double
P0 = 0: p1 = 1: p2 = p1
For i = 1 To n
    p2 = (2 * i - 1 - x) * p1 - (i - 1) ^ 2 * P0
    P0 = p1
    p1 = p2
Next i
Y = p2
End Sub

'-------------------------------------------------------------------------------
' Chebycev's polynomials
'-------------------------------------------------------------------------------
Sub PChebycev(ByVal x As Double, ByVal n As Double, ByRef Y As Double)
Dim i As Long, P0 As Double, p1 As Double, p2 As Double
If n = 0 Then Y = 1: Exit Sub
If n = 1 Then Y = x: Exit Sub
P0 = 1: p1 = x
For i = 1 To n - 1
    p2 = 2 * x * p1 - P0
    P0 = p1
    p1 = p2
Next i
Y = p2
End Sub

'------------------------------------------------------------------------------------
' special periodic functions
'-----------------------------------------------------------------------------------
Private Function MopUp(ByVal x As Double) As Double
If Abs(x) < 0.00000000000005 Then x = 0
MopUp = x
End Function
'triangular wave
Function WAVE_TRI(ByVal t As Double, ByVal p As Double) As Double
WAVE_TRI = MopUp(4 * Abs(Int(t / p + 1 / 2) - t / p) - 1)
End Function
'square wave
Function WAVE_SQR(ByVal t As Double, ByVal p As Double) As Double
WAVE_SQR = MopUp(-2 * Int(t / p + 1 / 2) + 2 * Int(t / p) + 1)
End Function
'rectangular wave
Function WAVE_RECT(ByVal t As Double, ByVal p As Double, ByVal duty_cicle As Double) As Double
WAVE_RECT = MopUp(-2 * Int(t / p - duty_cicle) + 2 * Int(t / p) - 1)
End Function
'trapez. wave
Function WAVE_TRAPEZ(ByVal t As Double, ByVal p As Double, ByVal duty_cicle As Double) As Double
Dim Y As Double
Y = 1 / duty_cicle * (2 * Abs(Int(t / p + 1 / 2) - t / p) - Abs(2 * Int(-duty_cicle / (2 * p) + t / p + 1 / 2) + (duty_cicle - 2 * t) / p))
WAVE_TRAPEZ = MopUp(Y)
End Function
'Saw wave
Function WAVE_SAW(ByVal t As Double, ByVal p As Double) As Double
WAVE_SAW = MopUp(2 * t / p - 2 * Int(t / p + 1 / 2))
End Function
'Rampa wave
Function WAVE_RAISE(ByVal t As Double, ByVal p As Double) As Double
WAVE_RAISE = MopUp(t / p - Int(t / p))
End Function
'Linear wave
Function WAVE_LIN(ByVal t As Double, ByVal p As Double, ByVal duty_cicle As Double) As Double
Dim Y As Double
Y = (p * Int(t / p - duty_cicle) ^ 2 + (2 * duty_cicle * p + p - 2 * t) * Int(t / p - duty_cicle) - p * Int(-duty_cicle) ^ 2 - p * (2 * duty_cicle + 1) _
    * Int(-duty_cicle) - p * Int(t / p) ^ 2 + (2 * t - p) * Int(t / p) + duty_cicle * (duty_cicle * p - p - 2 * t)) / (duty_cicle * p * (1 - duty_cicle))
WAVE_LIN = MopUp(Y)
End Function
'rectangular pulse wave
Function WAVE_PULSE(ByVal t As Double, ByVal p As Double, ByVal duty_cicle As Double) As Double
WAVE_PULSE = MopUp(-Int(t / p - duty_cicle) + Int(t / p))
End Function
'steps wave
Function WAVE_STEPS(ByVal t As Double, ByVal p As Double, ByVal n As Double) As Double
WAVE_STEPS = MopUp(1 / (n - 1) * (Int(n * t / p) - n * Int(t / p)))
End Function
'exponential pulse wave
Function WAVE_EXP(ByVal t As Double, ByVal p As Double, ByVal A As Double) As Double
WAVE_EXP = MopUp(Exp(-A * t / p + A * Int(t / p)))
End Function
'exponential bipolar pulse wave
Function WAVE_EXPB(ByVal t As Double, ByVal p As Double, ByVal A As Double) As Double
WAVE_EXPB = MopUp(Exp(-A * t / p + A * Int(t / p)) - Exp(-A * (t / p + 1 / 2) + A * Int(t / p + 1 / 2)))
End Function
'filtered pulse wave
Function WAVE_PULSEF(ByVal t As Double, ByVal p As Double, ByVal A As Double) As Double
WAVE_PULSEF = (-Int(t / p + 1 / 2) + Int(t / p) + 1 - (Exp(-A * t / p + A * Int(t / p)) - Exp(-A * (t / p + 1 / 2) + A * Int(t / p + 1 / 2))))
End Function
'ringing wave
Function WAVE_RING(ByVal t As Double, ByVal p As Double, ByVal A As Double, ByVal omega As Double) As Double
WAVE_RING = (-Exp(A * Int(t / p) - A * t / p) * Sin(2 * PI_ * omega * Int(t / p) - 2 * PI_ * omega * t / p))
End Function
'parabolic pulse wave
Function WAVE_PARAB(ByVal t As Double, ByVal p As Double) As Double
WAVE_PARAB = MopUp((2 * Abs(Int(t / p + 1 / 2) - t / p)) ^ 2)
End Function
'ripple wave
Function WAVE_RIPPLE(ByVal t As Double, ByVal p As Double, ByVal A As Double) As Double
Dim x As Double, Y As Double, r As Double
Y = Abs(Cos(PI_ / p * t))
x = Exp(A * Int(t / p) - A * t / p)
If x > Y Then r = x Else r = Y
WAVE_RIPPLE = r
End Function
'rectifire wave
Function WAVE_SINREC(ByVal t As Double, ByVal p As Double) As Double
WAVE_SINREC = Abs(Sin(PI_ * t / p))
End Function
'Amplitude modulation
Function WAVE_AM(ByVal t As Double, ByVal fo As Double, ByVal fm As Double, ByVal m As Double) As Double
WAVE_AM = (1 + m * Sin(2 * PI_ * fm * t)) * Sin(2 * PI_ * fo * t)
End Function
'frequecy modulation
Function WAVE_FM(ByVal t As Double, ByVal fo As Double, ByVal fm As Double, ByVal m As Double) As Double
WAVE_FM = Sin(2 * PI_ * fo * (1 + m * Sin(2 * PI_ * fm * t)) * t)
End Function


'***********  End of Library for computation of Special Functions ******************



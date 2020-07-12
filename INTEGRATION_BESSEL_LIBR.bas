Attribute VB_Name = "INTEGRATION_BESSEL_LIBR"
Private Const MachineEpsilon = 5E-16
Private Const MaxRealNumber = 1E+300
Private Const MinRealNumber = 1E-300

Private Const BigNumber As Double = 1E+70
Private Const SmallNumber As Double = 1E-70
Private Const PiNumber As Double = 3.14159265358979
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Cephes Math Library Release 2.8:  June, 2000
'Copyright by Stephen L. Moshier
'
'Contributors:
'    * Sergey Bochkanov (ALGLIB project). Translation from C to
'      pseudocode.
'
'See subroutines comments for additional copyrights.
'
'Redistribution and use in source and binary forms, with or without
'modification, are permitted provided that the following conditions are
'met:
'
'- Redistributions of source code must retain the above copyright
'  notice, this list of conditions and the following disclaimer.
'
'- Redistributions in binary form must reproduce the above copyright
'  notice, this list of conditions and the following disclaimer listed
'  in this license in the documentation and/or other materials
'  provided with the distribution.
'
'- Neither the name of the copyright holders nor the names of its
'  contributors may be used to endorse or promote products derived from
'  this software without specific prior written permission.
'
'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
'OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
'SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
'LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Routines
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Bessel function of order zero
'
'Returns Bessel function of order zero of the argument.
'
'The domain is divided into the intervals [0, 5] and
'(5, infinity). In the first interval the following rational
'approximation is used:
'
'
'       2         2
'(w - r  ) (w - r  ) P (w) / Q (w)
'      1         2    3       8
'
'           2
'where w = x  and the two r's are zeros of the function.
'
'In the second interval, the Hankel asymptotic expansion
'is employed with two rational functions of degree 6/6
'and 7/7.
'
'ACCURACY:
'
'                     Absolute error:
'arithmetic   domain     # trials      peak         rms
'   IEEE      0, 30       60000       4.2e-16     1.1e-16
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 1989, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BESSEL_J0_FUNC(ByVal x As Double) As Double
    Dim Result As Double
    Dim XSq As Double
    Dim nn As Double
    Dim PZero As Double
    Dim QZero As Double
    Dim p1 As Double
    Dim q1 As Double

    If x < 0# Then
        x = -x
    End If
    If x > 8# Then
        Call BESSEL_ASYMPT0_FUNC(x, PZero, QZero)
        nn = x - PiNumber / 4#
        Result = Sqr(2# / PiNumber / x) * (PZero * Cos(nn) - QZero * Sin(nn))
        BESSEL_J0_FUNC = Result
        Exit Function
    End If
    XSq = Square(x)
    p1 = 26857.8685698002
    p1 = -40504123.7183313 + XSq * p1
    p1 = 25071582855.3688 + XSq * p1
    p1 = -8085222034853.79 + XSq * p1
    p1 = 1.43435493914034E+15 + XSq * p1
    p1 = -1.36762035308817E+17 + XSq * p1
    p1 = 6.38205934107236E+18 + XSq * p1
    p1 = -1.17915762910761E+20 + XSq * p1
    p1 = 4.93378725179413E+20 + XSq * p1
    q1 = 1#
    q1 = 1363.06365232897 + XSq * q1
    q1 = 1114636.09846299 + XSq * q1
    q1 = 669998767.298224 + XSq * q1
    q1 = 312304311494.121 + XSq * q1
    q1 = 112775673967980# + XSq * q1
    q1 = 3.02463561670946E+16 + XSq * q1
    q1 = 5.42891838409228E+18 + XSq * q1
    q1 = 4.93378725179413E+20 + XSq * q1
    Result = p1 / q1

    BESSEL_J0_FUNC = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Bessel function of order one
'
'Returns Bessel function of order one of the argument.
'
'The domain is divided into the intervals [0, 8] and
'(8, infinity). In the first interval a 24 term Chebyshev
'expansion is used. In the second, the asymptotic
'trigonometric representation is employed using two
'rational functions of degree 5/5.
'
'ACCURACY:
'
'                     Absolute error:
'arithmetic   domain      # trials      peak         rms
'   IEEE      0, 30       30000       2.6e-16     1.1e-16
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 1989, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BESSEL_J1_FUNC(ByVal x As Double) As Double
    Dim Result As Double
    Dim s As Double
    Dim XSq As Double
    Dim nn As Double
    Dim PZero As Double
    Dim QZero As Double
    Dim p1 As Double
    Dim q1 As Double

    s = Sgn(x)
    If x < 0# Then
        x = -x
    End If
    If x > 8# Then
        Call BESSEL_ASYMPT1_FUNC(x, PZero, QZero)
        nn = x - 3# * PiNumber / 4#
        Result = Sqr(2# / PiNumber / x) * (PZero * Cos(nn) - QZero * Sin(nn))
        If s < 0# Then
            Result = -Result
        End If
        BESSEL_J1_FUNC = Result
        Exit Function
    End If
    XSq = Square(x)
    p1 = 2701.12271089232
    p1 = -4695753.530643 + XSq * p1
    p1 = 3413234182.3017 + XSq * p1
    p1 = -1322983480332.13 + XSq * p1
    p1 = 290879526383478# + XSq * p1
    p1 = -3.58881756991011E+16 + XSq * p1
    p1 = 2.316433580634E+18 + XSq * p1
    p1 = -6.67210656892492E+19 + XSq * p1
    p1 = 5.81199354001606E+20 + XSq * p1
    q1 = 1#
    q1 = 1606.93157348149 + XSq * q1
    q1 = 1501793.59499859 + XSq * q1
    q1 = 1013863514.35867 + XSq * q1
    q1 = 524371026216.765 + XSq * q1
    q1 = 208166122130761# + XSq * q1
    q1 = 6.09206139891752E+16 + XSq * q1
    q1 = 1.18577071219032E+19 + XSq * q1
    q1 = 1.16239870800321E+21 + XSq * q1
    Result = s * x * p1 / q1

    BESSEL_J1_FUNC = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Bessel function of integer order
'
'Returns Bessel function of order n, where n is a
'(possibly negative) integer.
'
'The ratio of jn(x) to j0(x) is computed by backward
'recurrence.  First the ratio jn/jn-1 is found by a
'continued fraction expansion.  Then the recurrence
'relating successive orders is applied until j0 or j1 is
'reached.
'
'If n = 0 or 1 the routine for j0 or j1 is called
'directly.
'
'ACCURACY:
'
'                     Absolute error:
'arithmetic   range      # trials      peak         rms
'   IEEE      0, 30        5000       4.4e-16     7.9e-17
'
'
'Not suitable for large n or x. Use jv() (fractional order) instead.
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BESSEL_JN_FUNC(ByVal n As Long, ByVal x As Double) As Double
    Dim Result As Double
    Dim pkm2 As Double
    Dim pkm1 As Double
    Dim pk As Double
    Dim xk As Double
    Dim r As Double
    Dim ans As Double
    Dim k As Long
    Dim sg As Long

    If n < 0# Then
        n = -n
        If n Mod 2# = 0# Then
            sg = 1#
        Else
            sg = -1#
        End If
    Else
        sg = 1#
    End If
    If x < 0# Then
        If n Mod 2# <> 0# Then
            sg = -sg
        End If
        x = -x
    End If
    If n = 0# Then
        Result = sg * BESSEL_J0_FUNC(x)
        BESSEL_JN_FUNC = Result
        Exit Function
    End If
    If n = 1# Then
        Result = sg * BESSEL_J1_FUNC(x)
        BESSEL_JN_FUNC = Result
        Exit Function
    End If
    If n = 2# Then
        If x = 0# Then
            Result = 0#
        Else
            Result = sg * (2# * BESSEL_J1_FUNC(x) / x - BESSEL_J0_FUNC(x))
        End If
        BESSEL_JN_FUNC = Result
        Exit Function
    End If
    If x < MachineEpsilon Then
        Result = 0#
        BESSEL_JN_FUNC = Result
        Exit Function
    End If
    k = 53#
    pk = 2# * (n + k)
    ans = pk
    xk = x * x
    Do
        pk = pk - 2#
        ans = pk - xk / ans
        k = k - 1#
    Loop Until k = 0#
    ans = x / ans
    pk = 1#
    pkm1 = 1# / ans
    k = n - 1#
    r = 2# * k
    Do
        pkm2 = (pkm1 * r - pk * x) / x
        pk = pkm1
        pkm1 = pkm2
        r = r - 2#
        k = k - 1#
    Loop Until k = 0#
    If Abs(pk) > Abs(pkm1) Then
        ans = BESSEL_J1_FUNC(x) / pk
    Else
        ans = BESSEL_J0_FUNC(x) / pkm1
    End If
    Result = sg * ans

    BESSEL_JN_FUNC = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Bessel function of the second kind, order zero
'
'Returns Bessel function of the second kind, of order
'zero, of the argument.
'
'The domain is divided into the intervals [0, 5] and
'(5, infinity). In the first interval a rational approximation
'R(x) is employed to compute
'  y0(x)  = R(x)  +   2 * log(x) * j0(x) / PI.
'Thus a call to j0() is required.
'
'In the second interval, the Hankel asymptotic expansion
'is employed with two rational functions of degree 6/6
'and 7/7.
'
'
'
'ACCURACY:
'
' Absolute error, when y0(x) < 1; else relative error:
'
'arithmetic   domain     # trials      peak         rms
'   IEEE      0, 30       30000       1.3e-15     1.6e-16
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 1989, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BESSEL_Y0_FUNC(ByVal x As Double) As Double
    Dim Result As Double
    Dim nn As Double
    Dim XSq As Double
    Dim PZero As Double
    Dim QZero As Double
    Dim P4 As Double
    Dim Q4 As Double

    If x > 8# Then
        Call BESSEL_ASYMPT0_FUNC(x, PZero, QZero)
        nn = x - PiNumber / 4#
        Result = Sqr(2# / PiNumber / x) * (PZero * Sin(nn) + QZero * Cos(nn))
        BESSEL_Y0_FUNC = Result
        Exit Function
    End If
    XSq = Square(x)
    P4 = -41370.3549793315
    P4 = 59152134.6568689 + XSq * P4
    P4 = -34363712229.7904 + XSq * P4
    P4 = 10255208596863.9 + XSq * P4
    P4 = -1.64860581718573E+15 + XSq * P4
    P4 = 1.37562431639934E+17 + XSq * P4
    P4 = -5.24706558111277E+18 + XSq * P4
    P4 = 6.58747327571955E+19 + XSq * P4
    P4 = -2.75028667862911E+19 + XSq * P4
    Q4 = 1#
    Q4 = 1282.45277247899 + XSq * Q4
    Q4 = 1001702.64128891 + XSq * Q4
    Q4 = 579512264.070073 + XSq * Q4
    Q4 = 261306575504.108 + XSq * Q4
    Q4 = 91620380340751.9 + XSq * Q4
    Q4 = 2.39288304349978E+16 + XSq * Q4
    Q4 = 4.19241704341084E+18 + XSq * Q4
    Q4 = 3.72645883898617E+20 + XSq * Q4
    Result = P4 / Q4 + 2# / PiNumber * BESSEL_J0_FUNC(x) * Log(x)

    BESSEL_Y0_FUNC = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Bessel function of second kind of order one
'
'Returns Bessel function of the second kind of order one
'of the argument.
'
'The domain is divided into the intervals [0, 8] and
'(8, infinity). In the first interval a 25 term Chebyshev
'expansion is used, and a call to j1() is required.
'In the second, the asymptotic trigonometric representation
'is employed using two rational functions of degree 5/5.
'
'ACCURACY:
'
'                     Absolute error:
'arithmetic   domain      # trials      peak         rms
'   IEEE      0, 30       30000       1.0e-15     1.3e-16
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 1989, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BESSEL_Y1_FUNC(ByVal x As Double) As Double
    Dim Result As Double
    Dim nn As Double
    Dim XSq As Double
    Dim PZero As Double
    Dim QZero As Double
    Dim P4 As Double
    Dim Q4 As Double

    If x > 8# Then
        Call BESSEL_ASYMPT1_FUNC(x, PZero, QZero)
        nn = x - 3# * PiNumber / 4#
        Result = Sqr(2# / PiNumber / x) * (PZero * Sin(nn) + QZero * Cos(nn))
        BESSEL_Y1_FUNC = Result
        Exit Function
    End If
    XSq = Square(x)
    P4 = -2108847.54013312
    P4 = 3639488548.124 + XSq * P4
    P4 = -2580681702194.45 + XSq * P4
    P4 = 956993023992168# + XSq * P4
    P4 = -1.96588746272214E+17 + XSq * P4
    P4 = 2.1931073399178E+19 + XSq * P4
    P4 = -1.21229755541451E+21 + XSq * P4
    P4 = 2.65547383143485E+22 + XSq * P4
    P4 = -9.96375342430692E+22 + XSq * P4
    Q4 = 1#
    Q4 = 1612.361029677 + XSq * Q4
    Q4 = 1563282.75489958 + XSq * Q4
    Q4 = 1128686837.16944 + XSq * Q4
    Q4 = 646534088126.528 + XSq * Q4
    Q4 = 297663212564728# + XSq * Q4
    Q4 = 1.08225825940882E+17 + XSq * Q4
    Q4 = 2.95498793589715E+19 + XSq * Q4
    Q4 = 5.43531037718885E+21 + XSq * Q4
    Q4 = 5.08206736694124E+23 + XSq * Q4
    Result = x * P4 / Q4 + 2# / PiNumber * (BESSEL_J1_FUNC(x) * Log(x) - 1# / x)

    BESSEL_Y1_FUNC = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Bessel function of second kind of integer order
'
'Returns Bessel function of order n, where n is a
'(possibly negative) integer.
'
'The function is evaluated by forward recurrence on
'n, starting with values computed by the routines
'y0() and y1().
'
'If n = 0 or 1 the routine for y0 or y1 is called
'directly.
'
'ACCURACY:
'                     Absolute error, except relative
'                     when y > 1:
'arithmetic   domain     # trials      peak         rms
'   IEEE      0, 30       30000       3.4e-15     4.3e-16
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BESSEL_YN_FUNC(ByVal n As Long, ByVal x As Double) As Double
    Dim Result As Double
    Dim i As Long
    Dim A As Double
    Dim B As Double
    Dim tmp As Double
    Dim s As Double

    s = 1#
    If n < 0# Then
        n = -n
        If n Mod 2# <> 0# Then
            s = -1#
        End If
    End If
    If n = 0# Then
        Result = BESSEL_Y0_FUNC(x)
        BESSEL_YN_FUNC = Result
        Exit Function
    End If
    If n = 1# Then
        Result = s * BESSEL_Y1_FUNC(x)
        BESSEL_YN_FUNC = Result
        Exit Function
    End If
    A = BESSEL_Y0_FUNC(x)
    B = BESSEL_Y1_FUNC(x)
    For i = 1# To n - 1# Step 1
        tmp = B
        B = 2# * i / x * B - A
        A = tmp
    Next i
    Result = s * B

    BESSEL_YN_FUNC = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modified Bessel function of order zero
'
'Returns modified Bessel function of order zero of the
'argument.
'
'The function is defined as i0(x) = j0( ix ).
'
'The range is partitioned into the two intervals [0,8] and
'(8, infinity).  Chebyshev polynomial expansions are employed
'in each interval.
'
'ACCURACY:
'
'                     Relative error:
'arithmetic   domain     # trials      peak         rms
'   IEEE      0,30        30000       5.8e-16     1.4e-16
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BESSEL_I0_FUNC(ByVal x As Double) As Double
    Dim Result As Double
    Dim Y As Double
    Dim V As Double
    Dim z As Double
    Dim B0 As Double
    Dim b1 As Double
    Dim b2 As Double

    If x < 0# Then
        x = -x
    End If
    If x <= 8# Then
        Y = x / 2# - 2#
        Call BESSEL_M_FIRST_CHEB_FUNC(-4.41534164647934E-18, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 3.33079451882224E-17, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -2.43127984654795E-16, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 1.71539128555513E-15, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -1.16853328779935E-14, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 7.67618549860494E-14, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -4.85644678311193E-13, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 2.95505266312964E-12, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -1.72682629144156E-11, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 9.67580903537324E-11, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -5.18979560163526E-10, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 2.65982372468239E-09, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -1.30002500998625E-08, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 6.04699502254192E-08, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -2.67079385394061E-07, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 1.1173875391201E-06, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -4.41673835845875E-06, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 1.64484480707289E-05, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -5.7541950100821E-05, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 1.88502885095842E-04, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -5.76375574538582E-04, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 1.63947561694134E-03, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -4.32430999505058E-03, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 0.010546460394595, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -2.37374148058995E-02, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 4.93052842396707E-02, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -9.49010970480476E-02, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 0.171620901522209, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -0.304682672343198, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 0.676795274409476, B0, b1, b2)
        V = 0.5 * (B0 - b2)
        Result = Exp(x) * V
        BESSEL_I0_FUNC = Result
        Exit Function
    End If
    z = 32# / x - 2#
    Call BESSEL_M_FIRST_CHEB_FUNC(-7.23318048787475E-18, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, -4.83050448594418E-18, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 4.46562142029676E-17, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 3.46122286769746E-17, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, -2.82762398051658E-16, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, -3.42548561967722E-16, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 1.77256013305653E-15, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 3.81168066935262E-15, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, -9.55484669882831E-15, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, -4.15056934728722E-14, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 1.54008621752141E-14, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 3.85277838274214E-13, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 7.18012445138367E-13, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, -1.79417853150681E-12, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, -1.32158118404477E-11, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, -3.14991652796324E-11, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 1.18891471078464E-11, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 4.94060238822497E-10, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 3.39623202570839E-09, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 2.26666899049818E-08, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 2.04891858946906E-07, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 2.89137052083476E-06, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 6.88975834691682E-05, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 3.36911647825569E-03, B0, b1, b2)
    Call BESSEL_M_NEXT_CHEB_FUNC(z, 0.804490411014109, B0, b1, b2)
    V = 0.5 * (B0 - b2)
    Result = Exp(x) * V / Sqr(x)

    BESSEL_I0_FUNC = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modified Bessel function of order one
'
'Returns modified Bessel function of order one of the
'argument.
'
'The function is defined as i1(x) = -i j1( ix ).
'
'The range is partitioned into the two intervals [0,8] and
'(8, infinity).  Chebyshev polynomial expansions are employed
'in each interval.
'
'ACCURACY:
'
'                     Relative error:
'arithmetic   domain     # trials      peak         rms
'   IEEE      0, 30       30000       1.9e-15     2.1e-16
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1985, 1987, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BESSEL_I1_FUNC(ByVal x As Double) As Double
    Dim Result As Double
    Dim Y As Double
    Dim z As Double
    Dim V As Double
    Dim B0 As Double
    Dim b1 As Double
    Dim b2 As Double

    z = Abs(x)
    If z <= 8# Then
        Y = z / 2# - 2#
        Call BESSEL_M1_FIRST_CHEB_FUNC(2.77791411276105E-18, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -2.11142121435817E-17, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 1.5536319577362E-16, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.10559694773539E-15, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 7.60068429473541E-15, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -5.04218550472791E-14, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 3.22379336594557E-13, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.98397439776494E-12, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 1.17361862988909E-11, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -6.66348972350203E-11, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 3.62559028155212E-10, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.88724975172283E-09, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 9.38153738649577E-09, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -4.44505912879633E-08, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 2.00329475355214E-07, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -8.56872026469545E-07, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 3.47025130813768E-06, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.32731636560394E-05, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 4.78156510755005E-05, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.61760815825897E-04, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 5.12285956168576E-04, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.51357245063125E-03, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 4.15642294431289E-03, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.05640848946262E-02, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 2.47264490306265E-02, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -0.052945981208095, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 0.102643658689847, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -0.176416518357834, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 0.252587186443634, B0, b1, b2)
        V = 0.5 * (B0 - b2)
        z = V * z * Exp(z)
    Else
        Y = 32# / z - 2#
        Call BESSEL_M1_FIRST_CHEB_FUNC(7.51729631084211E-18, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 4.41434832307171E-18, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -4.65030536848936E-17, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -3.20952592199342E-17, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 2.96262899764595E-16, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 3.30820231092093E-16, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.88035477551078E-15, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -3.81440307243701E-15, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 1.04202769841288E-14, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 4.27244001671195E-14, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -2.10154184277266E-14, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -4.0835511110922E-13, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -7.19855177624591E-13, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 2.03562854414709E-12, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 1.41258074366138E-11, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 3.25260358301549E-11, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.89749581235054E-11, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -5.58974346219658E-10, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -3.83538038596424E-09, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -2.63146884688952E-08, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -2.51223623787021E-07, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -3.88256480887769E-06, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.10588938762624E-04, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -9.76109749136147E-03, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 0.77857623501828, B0, b1, b2)
        V = 0.5 * (B0 - b2)
        z = V * Exp(z) / Sqr(z)
    End If
    If x < 0# Then
        z = -z
    End If
    Result = z

    BESSEL_I1_FUNC = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modified Bessel function, second kind, order zero
'
'Returns modified Bessel function of the second kind
'of order zero of the argument.
'
'The range is partitioned into the two intervals [0,8] and
'(8, infinity).  Chebyshev polynomial expansions are employed
'in each interval.
'
'ACCURACY:
'
'Tested at 2000 random points between 0 and 8.  Peak absolute
'error (relative when K0 > 1) was 1.46e-14; rms, 4.26e-15.
'                     Relative error:
'arithmetic   domain     # trials      peak         rms
'   IEEE      0, 30       30000       1.2e-15     1.6e-16
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BESSEL_K0_FUNC(ByVal x As Double) As Double
    Dim Result As Double
    Dim Y As Double
    Dim z As Double
    Dim V As Double
    Dim B0 As Double
    Dim b1 As Double
    Dim b2 As Double

    If x <= 2# Then
        Y = x * x - 2#
        Call BESSEL_M_FIRST_CHEB_FUNC(1.37446543561352E-16, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 4.25981614279661E-14, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 1.03496952576338E-11, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 1.90451637722021E-09, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 2.53479107902615E-07, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 2.28621210311945E-05, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 1.26461541144693E-03, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 3.59799365153615E-02, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, 0.344289899924629, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(Y, -0.535327393233903, B0, b1, b2)
        V = 0.5 * (B0 - b2)
        V = V - Log(0.5 * x) * BESSEL_I0_FUNC(x)
    Else
        z = 8# / x - 2#
        Call BESSEL_M_FIRST_CHEB_FUNC(5.30043377268626E-18, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -1.64758043015242E-17, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 5.21039150503903E-17, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -1.67823109680541E-16, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 5.51205597852432E-16, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -1.84859337734378E-15, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 6.34007647740507E-15, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -2.22751332699167E-14, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 8.03289077536358E-14, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -2.98009692317273E-13, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 1.14034058820848E-12, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -4.51459788337394E-12, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 1.85594911495472E-11, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -7.95748924447711E-11, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 3.5773972814003E-10, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -1.69753450938906E-09, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 8.57403401741423E-09, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -4.66048989768795E-08, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 2.76681363944501E-07, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -1.83175552271912E-06, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 1.39498137188765E-05, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -1.28495495816278E-04, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 1.56988388573005E-03, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, -3.14481013119645E-02, B0, b1, b2)
        Call BESSEL_M_NEXT_CHEB_FUNC(z, 2.44030308206596, B0, b1, b2)
        V = 0.5 * (B0 - b2)
        V = V * Exp(-x) / Sqr(x)
    End If
    Result = V

    BESSEL_K0_FUNC = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modified Bessel function, second kind, order one
'
'Computes the modified Bessel function of the second kind
'of order one of the argument.
'
'The range is partitioned into the two intervals [0,2] and
'(2, infinity).  Chebyshev polynomial expansions are employed
'in each interval.
'
'ACCURACY:
'
'                     Relative error:
'arithmetic   domain     # trials      peak         rms
'   IEEE      0, 30       30000       1.2e-15     1.6e-16
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BESSEL_K1_FUNC(ByVal x As Double) As Double
    Dim Result As Double
    Dim Y As Double
    Dim z As Double
    Dim V As Double
    Dim B0 As Double
    Dim b1 As Double
    Dim b2 As Double

    z = 0.5 * x
    If x <= 2# Then
        Y = x * x - 2#
        Call BESSEL_M1_FIRST_CHEB_FUNC(-7.02386347938629E-18, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -2.42744985051937E-15, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -6.66690169419933E-13, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.41148839263353E-10, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -2.21338763073473E-08, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -2.43340614156597E-06, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.73028895751305E-04, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -6.97572385963986E-03, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -0.122611180822657, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -0.353155960776545, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 1.52530022733895, B0, b1, b2)
        V = 0.5 * (B0 - b2)
        Result = Log(z) * BESSEL_I1_FUNC(x) + V / x
    Else
        Y = 8# / x - 2#
        Call BESSEL_M1_FIRST_CHEB_FUNC(-5.75674448366502E-18, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 1.79405087314756E-17, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -5.68946255844286E-17, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 1.83809354436664E-16, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -6.05704724837332E-16, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 2.03870316562433E-15, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -7.01983709041831E-15, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 2.4771544244813E-14, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -8.97670518232499E-14, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 3.34841966607843E-13, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.28917396095103E-12, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 5.13963967348173E-12, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -2.12996783842757E-11, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 9.21831518760501E-11, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -4.1903547593419E-10, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 2.01504975519703E-09, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.03457624656781E-08, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 5.74108412545005E-08, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -3.50196060308781E-07, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 2.40648494783722E-06, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -1.93619797416608E-05, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 1.95215518471352E-04, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, -2.85781685962278E-03, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 0.103923736576817, B0, b1, b2)
        Call BESSEL_M1_NEXT_CHEB_FUNC(Y, 2.72062619048444, B0, b1, b2)
        V = 0.5 * (B0 - b2)
        Result = Exp(-x) * V / Sqr(x)
    End If

    BESSEL_K1_FUNC = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modified Bessel function, second kind, integer order
'
'Returns modified Bessel function of the second kind
'of order n of the argument.
'
'The range is partitioned into the two intervals [0,9.55] and
'(9.55, infinity).  An ascending power series is used in the
'low range, and an asymptotic expansion in the high range.
'
'ACCURACY:
'
'                     Relative error:
'arithmetic   domain     # trials      peak         rms
'   IEEE      0,30        90000       1.8e-8      3.0e-10
'
'Error is high only near the crossover point x = 9.55
'between the two expansions used.
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 1988, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BESSEL_KN_FUNC(ByVal nn As Long, ByVal x As Double) As Double
    Dim Result As Double
    Dim k As Double
    Dim kf As Double
    Dim nk1f As Double
    Dim nkf As Double
    Dim zn As Double
    Dim t As Double
    Dim s As Double
    Dim z0 As Double
    Dim z As Double
    Dim ans As Double
    Dim fn As Double
    Dim pn As Double
    Dim pk As Double
    Dim zmn As Double
    Dim tlg As Double
    Dim tox As Double
    Dim i As Long
    Dim n As Long
    Dim EUL As Double

    EUL = 0.577215664901533
    If nn < 0# Then
        n = -nn
    Else
        n = nn
    End If
    If x <= 9.55 Then
        ans = 0#
        z0 = 0.25 * x * x
        fn = 1#
        pn = 0#
        zmn = 1#
        tox = 2# / x
        If n > 0# Then
            pn = -EUL
            k = 1#
            For i = 1# To n - 1# Step 1
                pn = pn + 1# / k
                k = k + 1#
                fn = fn * k
            Next i
            zmn = tox
            If n = 1# Then
                ans = 1# / x
            Else
                nk1f = fn / n
                kf = 1#
                s = nk1f
                z = -z0
                zn = 1#
                For i = 1# To n - 1# Step 1
                    nk1f = nk1f / (n - i)
                    kf = kf * i
                    zn = zn * z
                    t = nk1f * zn / kf
                    s = s + t
                    zmn = zmn * tox
                Next i
                s = s * 0.5
                t = Abs(s)
                ans = s * zmn
            End If
        End If
        tlg = 2# * Log(0.5 * x)
        pk = -EUL
        If n = 0# Then
            pn = pk
            t = 1#
        Else
            pn = pn + 1# / n
            t = 1# / fn
        End If
        s = (pk + pn - tlg) * t
        k = 1#
        Do
            t = t * (z0 / (k * (k + n)))
            pk = pk + 1# / k
            pn = pn + 1# / (k + n)
            s = s + (pk + pn - tlg) * t
            k = k + 1#
        Loop Until Abs(t / s) <= MachineEpsilon
        s = 0.5 * s / zmn
        If n Mod 2# <> 0# Then
            s = -s
        End If
        ans = ans + s
        Result = ans
        BESSEL_KN_FUNC = Result
        Exit Function
    End If
    If x > Log(MaxRealNumber) Then
        Result = 0#
        BESSEL_KN_FUNC = Result
        Exit Function
    End If
    k = n
    pn = 4# * k * k
    pk = 1#
    z0 = 8# * x
    fn = 1#
    t = 1#
    s = t
    nkf = MaxRealNumber
    i = 0#
    Do
        z = pn - pk * pk
        t = t * z / (fn * z0)
        nk1f = Abs(t)
        If i >= n And nk1f > nkf Then
            Exit Do
        End If
        nkf = nk1f
        s = s + t
        fn = fn + 1#
        pk = pk + 2#
        i = i + 1#
    Loop Until Abs(t / s) <= MachineEpsilon
    Result = Exp(-x) * Sqr(PiNumber / (2# * x)) * s

    BESSEL_KN_FUNC = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal subroutine
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BESSEL_M_FIRST_CHEB_FUNC(ByVal c As Double, _
         ByRef B0 As Double, _
         ByRef b1 As Double, _
         ByRef b2 As Double)

    B0 = c
    b1 = 0#
    b2 = 0#
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal subroutine
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BESSEL_M_NEXT_CHEB_FUNC(ByVal x As Double, _
         ByVal c As Double, _
         ByRef B0 As Double, _
         ByRef b1 As Double, _
         ByRef b2 As Double)

    b2 = b1
    b1 = B0
    B0 = x * b1 - b2 + c
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal subroutine
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BESSEL_M1_FIRST_CHEB_FUNC(ByVal c As Double, _
         ByRef B0 As Double, _
         ByRef b1 As Double, _
         ByRef b2 As Double)

    B0 = c
    b1 = 0#
    b2 = 0#
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal subroutine
'
'Cephes Math Library Release 2.8:  June, 2000
'Copyright 1984, 1987, 2000 by Stephen L. Moshier
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BESSEL_M1_NEXT_CHEB_FUNC(ByVal x As Double, _
         ByVal c As Double, _
         ByRef B0 As Double, _
         ByRef b1 As Double, _
         ByRef b2 As Double)

    b2 = b1
    b1 = B0
    B0 = x * b1 - b2 + c
End Sub


Private Sub BESSEL_ASYMPT0_FUNC(ByVal x As Double, _
         ByRef PZero As Double, _
         ByRef QZero As Double)
    Dim XSq As Double
    Dim p2 As Double
    Dim Q2 As Double
    Dim P3 As Double
    Dim Q3 As Double

    XSq = 64# / (x * x)
    p2 = 0#
    p2 = 2485.2719289574 + XSq * p2
    p2 = 153982.653262391 + XSq * p2
    p2 = 2016135.28304998 + XSq * p2
    p2 = 8413041.45655044 + XSq * p2
    p2 = 12332384.7681764 + XSq * p2
    p2 = 5393485.08386944 + XSq * p2
    Q2 = 1#
    Q2 = 2615.70073692084 + XSq * Q2
    Q2 = 156001.727694003 + XSq * Q2
    Q2 = 2025066.80157013 + XSq * Q2
    Q2 = 8426449.0506298 + XSq * Q2
    Q2 = 12338310.2278633 + XSq * Q2
    Q2 = 5393485.08386944 + XSq * Q2
    P3 = -0#
    P3 = -4.88719939584126 + XSq * P3
    P3 = -226.26306419337 + XSq * P3
    P3 = -2365.95617077911 + XSq * P3
    P3 = -8239.06631348561 + XSq * P3
    P3 = -10381.4169874846 + XSq * P3
    P3 = -3984.61735759522 + XSq * P3
    Q3 = 1#
    Q3 = 408.77146739835 + XSq * Q3
    Q3 = 15704.891915154 + XSq * Q3
    Q3 = 156021.320667929 + XSq * Q3
    Q3 = 533291.36342169 + XSq * Q3
    Q3 = 666745.423931983 + XSq * Q3
    Q3 = 255015.510886094 + XSq * Q3
    PZero = p2 / Q2
    QZero = 8# * P3 / Q3 / x
End Sub


Private Sub BESSEL_ASYMPT1_FUNC(ByVal x As Double, _
         ByRef PZero As Double, _
         ByRef QZero As Double)
    Dim XSq As Double
    Dim p2 As Double
    Dim Q2 As Double
    Dim P3 As Double
    Dim Q3 As Double

    XSq = 64# / (x * x)
    p2 = -1611.61664432461
    p2 = -109824.055434593 + XSq * p2
    p2 = -1523529.35118114 + XSq * p2
    p2 = -6603373.24836494 + XSq * p2
    p2 = -9942246.50507764 + XSq * p2
    p2 = -4435757.81679413 + XSq * p2
    Q2 = 1#
    Q2 = -1455.0094401905 + XSq * Q2
    Q2 = -107263.859911038 + XSq * Q2
    Q2 = -1511809.50663416 + XSq * Q2
    Q2 = -6585339.47972309 + XSq * Q2
    Q2 = -9934124.38993459 + XSq * Q2
    Q2 = -4435757.81679413 + XSq * Q2
    P3 = 35.265133846636
    P3 = 1706.37542902077 + XSq * P3
    P3 = 18494.2628732239 + XSq * P3
    P3 = 66178.8365812708 + XSq * P3
    P3 = 85145.1606753357 + XSq * P3
    P3 = 33220.9134098572 + XSq * P3
    Q3 = 1#
    Q3 = 863.836776960499 + XSq * Q3
    Q3 = 37890.2297457722 + XSq * Q3
    Q3 = 400294.43582267 + XSq * Q3
    Q3 = 1419460.66960372 + XSq * Q3
    Q3 = 1819458.042244 + XSq * Q3
    Q3 = 708712.819410287 + XSq * Q3
    PZero = p2 / Q2
    QZero = 8# * P3 / Q3 / x
End Sub

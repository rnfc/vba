Attribute VB_Name = "STAT_DIST_SMITH_LIBR"
'// Copyright © Ian Smith 2002-2007
'// Version 3.3.3
' requires >= Excel 2000 (for long lines [for Array assignment])
'// Thanks to Jerry W. Lewis for lots of help with testing and
'improvements to the code.
Option Explicit

Const NonIntegralValuesAllowed_df = False
' Are non-integral degrees of freedom for t, chi_square and f distributions allowed?
Const NonIntegralValuesAllowed_NB = False
' Is "successes required" parameter for negative binomial allowed to be non-integral?

Const NonIntegralValuesAllowed_Others = False
' Can Int function be applied to parameters like sample_size or is it a fault if
' the parameter is non-integral?

Const nc_limit = 1000000#
' Upper Limit for non-centrality parameters - as far as I know it's ok but slower
'and slower up to 1e12. Above that I don't know.
Const lstpi = 0.918938533204673
' 0.9189385332046727417803297364 = ln(sqrt(2*Pi))
Const sumAcc = 5E-16
Const cfSmall = 0.00000000000001
Const cfVSmall = 0.000000000000001
Const minLog1Value = -0.79149064
Const OneOverSqrTwoPi = 0.398942280401433
' 0.39894228040143267793994605993438
Const scalefactor = 1.15792089237316E+77
' 1.1579208923731619542357098500869e+77 = 2^256  ' used for rescaling
'calcs w/o impacting accuracy, to avoid over/underflow
Const scalefactor2 = 8.63616855509444E-78
' 8.6361685550944446253863518628004e-78 = 2^-256
Const max_discrete = 9.00719925474099E+15
' 2^53 required for exact addition of 1 in hypergeometric routines
Const max_crit = 4.5035996273705E+15
' 2^52 to make sure plenty of room for exact addition of 1 in crit routines
Const nearly_zero = 9.99999983659714E-317
Const cSmall = 5.562684646268E-309
' (smallest number before we start losing precision)/4
Const t_nc_limit = 1.34078079299426E+154
' just under 1/Sqr(cSmall)
Const Log1p5 = 0.405465108108164
' 0.40546510810816438197801311546435 = Log(1.5)
Const logfbit0p5 = 5.48141210519177E-02
' 0.054814121051917653896138702348386 = logfbit(0.5)

'For logfbit functions
' Stirling's series for ln(Gamma(x)), A046968/A046969
Const lfbc1 = 1# / 12#
Const lfbc2 = 1# / 30#
' lfbc2 on are Sloane's ratio times 12
Const lfbc3 = 1# / 105#
Const lfbc4 = 1# / 140#
Const lfbc5 = 1# / 99#
Const lfbc6 = 691# / 30030#
Const lfbc7 = 1# / 13#
Const lfbc8 = 0.350686068964593
' Chosen to make logfbit(6) & logfbit(7) correct
Const lfbc9 = 1.67699982016711
' Chosen to make logfbit(6) & logfbit(7) correct

'For invcnormal                             ' http://lib.stat.cmu.edu/apstat/241
Const A0 = 3.38713287279637                 ' 3.3871328727963666080
Const A1 = 133.141667891784                 ' 133.14166789178437745
Const A2 = 1971.59095030655                 ' 1971.5909503065514427
Const A3 = 13731.6937655095                 ' 13731.693765509461125
Const a4 = 45921.9539315499                 ' 45921.953931549871457
Const a5 = 67265.7709270087                 ' 67265.770927008700853
Const a6 = 33430.5755835881                 ' 33430.575583588128105
Const a7 = 2509.08092873012                 ' 2509.0809287301226727
Const b1 = 42.3133307016009                 ' 42.313330701600911252
Const b2 = 687.187007492058                 ' 687.18700749205790830
Const b3 = 5394.19602142475                 ' 5394.1960214247511077
Const b4 = 21213.7943015866                 ' 21213.794301586595867
Const b5 = 39307.8958000927                 ' 39307.895800092710610
Const b6 = 28729.0857357219                 ' 28729.085735721942674
Const b7 = 5226.49527885285                 ' 5226.4952788528545610
'//Coefficients for P not close to 0, 0.5 or 1.
Const C0 = 1.42343711074968                 ' 1.42343711074968357734
Const C1 = 4.63033784615655                 ' 4.63033784615654529590
Const C2 = 5.76949722146069                 ' 5.76949722146069140550
Const c3 = 3.6478483247632                  ' 3.64784832476320460504
Const c4 = 1.27045825245237                 ' 1.27045825245236838258
Const c5 = 0.241780725177451                ' 0.241780725177450611770
Const c6 = 2.27238449892692E-02             ' 2.27238449892691845833E-02
Const c7 = 7.74545014278341E-04             ' 7.74545014278341407640E-04
Const D1 = 2.05319162663776                 ' 2.05319162663775882187
Const d2 = 1.6763848301838                  ' 1.67638483018380384940
Const d3 = 0.6897673349851                  ' 0.689767334985100004550
Const d4 = 0.14810397642748                 ' 0.148103976427480074590
Const d5 = 1.51986665636165E-02             ' 1.51986665636164571966E-02
Const d6 = 5.47593808499535E-04             ' 5.47593808499534494600E-04
Const d7 = 1.05075007164442E-09             ' 1.05075007164441684324E-09
'//Coefficients for P near 0 or 1.
Const e0 = 6.6579046435011                  ' 6.65790464350110377720
Const e1 = 5.46378491116411                 ' 5.46378491116411436990
Const e2 = 1.78482653991729                 ' 1.78482653991729133580
Const e3 = 0.296560571828505                ' 0.296560571828504891230
Const e4 = 2.65321895265761E-02             ' 2.65321895265761230930E-02
Const e5 = 1.24266094738808E-03             ' 1.24266094738807843860E-03
Const e6 = 2.71155556874349E-05             ' 2.71155556874348757815E-05
Const e7 = 2.01033439929229E-07             ' 2.01033439929228813265E-07
Const F1 = 0.599832206555888                ' 0.599832206555887937690
Const F2 = 0.136929880922736                ' 0.136929880922735805310
Const f3 = 1.48753612908506E-02             ' 1.48753612908506148525E-02
Const f4 = 7.86869131145613E-04             ' 7.86869131145613259100E-04
Const f5 = 1.84631831751005E-05             ' 1.84631831751005468180E-05
Const f6 = 1.42151175831645E-07             ' 1.42151175831644588870E-07
Const f7 = 2.04426310338994E-15             ' 2.04426310338993978564E-15

'For poissapprox                            ' Stirling's series for Gamma(x), A001163/A001164
Const coef15 = 1# / 12#
Const coef25 = 1# / 288#
Const coef35 = -139# / 51840#
Const coef45 = -571# / 2488320#
Const coef55 = 163879# / 209018880#
Const coef65 = 5246819# / 75246796800#
Const coef75 = -534703531# / 902961561600#
Const coef1 = 2# / 3#                        ' Ramanujan's series for Gamma(x+1,x)-Gamma(x+1)/2, A065973
Const coef2 = -4# / 135#                     ' cf. http://www.whim.org/nebula/math/gammaratio.html
Const coef3 = 8# / 2835#
Const coef4 = 16# / 8505#
Const coef5 = -8992# / 12629925#
Const coef6 = -334144# / 492567075#
Const coef7 = 698752# / 1477701225#
Const coef8 = 23349012224# / 39565450299375#

Const twoThirds = 2# / 3#
Const twoFifths = 2# / 5#
Const twoSevenths = 2# / 7#
Const twoNinths = 2# / 9#
Const twoElevenths = 2# / 11#
Const twoThirteenths = 2# / 13#

'For binapprox
Const oneThird = 1# / 3#
Const twoTo27 = 134217728#                   ' 2^27

'For lngammaexpansion
Const eulers_const = 0.577215664901533      ' 0.5772156649015328606065120901

Private Function Min(ByVal x As Double, ByVal Y As Double) As Double
   If x < Y Then
      Min = x
   Else
      Min = Y
   End If
End Function
Private Function max(ByVal x As Double, ByVal Y As Double) As Double
   If x > Y Then
      max = x
   Else
      max = Y
   End If
End Function

Private Function expm1(ByVal x As Double) As Double
'// Accurate calculation of exp(x)-1, particularly for small x.
'// Uses a variation of the standard continued fraction for tanh(x) see A&S 4.5.70.
  If (Abs(x) < 2) Then
     Dim A1 As Double, A2 As Double, b1 As Double, b2 As Double, C1 As Double, x2 As Double
     A1 = 24#
     b1 = 2# * (12# - x * (6# - x))
     x2 = x * x * 0.25
     A2 = 8# * (15# + x2)
     b2 = 120# - x * (60# - x * (12# - x))
     C1 = 7#

     While ((Abs(A2 * b1 - A1 * b2) > Abs(cfSmall * b1 * A2)))

       A1 = C1 * A2 + x2 * A1
       b1 = C1 * b2 + x2 * b1
       C1 = C1 + 2#

       A2 = C1 * A1 + x2 * A2
       b2 = C1 * b1 + x2 * b2
       C1 = C1 + 2#
       If (b2 > scalefactor) Then
         A1 = A1 * scalefactor2
         b1 = b1 * scalefactor2
         A2 = A2 * scalefactor2
         b2 = b2 * scalefactor2
       End If
     Wend

     expm1 = x * A2 / b2
  Else
     expm1 = Exp(x) - 1#
  End If

End Function

Private Function logcf(ByVal x As Double, ByVal i As Double, ByVal D As Double) As Double
'// Continued fraction for calculation of 1/i + x/(i+d) + x*x/(i+2*d) + x*x*x/(i+3d) + ...
Dim A1 As Double, A2 As Double, b1 As Double, b2 As Double, C1 As Double, C2 As Double, c3 As Double, c4 As Double
     C1 = 2# * D
     C2 = i + D
     c4 = C2 + D
     A1 = C2
     b1 = i * (C2 - i * x)
     b2 = D * D * x
     A2 = c4 * C2 - b2
     b2 = c4 * b1 - i * b2

     While ((Abs(A2 * b1 - A1 * b2) > Abs(cfVSmall * b1 * A2)))

       c3 = C2 * C2 * x
       C2 = C2 + D
       c4 = c4 + D
       A1 = c4 * A2 - c3 * A1
       b1 = c4 * b2 - c3 * b1

       c3 = C1 * C1 * x
       C1 = C1 + D
       c4 = c4 + D
       A2 = c4 * A1 - c3 * A2
       b2 = c4 * b1 - c3 * b2
       If (b2 > scalefactor) Then
         A1 = A1 * scalefactor2
         b1 = b1 * scalefactor2
         A2 = A2 * scalefactor2
         b2 = b2 * scalefactor2
       ElseIf (b2 < scalefactor2) Then
         A1 = A1 * scalefactor
         b1 = b1 * scalefactor
         A2 = A2 * scalefactor
         b2 = b2 * scalefactor
       End If
     Wend
     logcf = A2 / b2
End Function

Private Function log0(ByVal x As Double) As Double
'//Accurate calculation of log(1+x), particularly for small x.
   Dim term As Double
   If (Abs(x) > 0.5) Then
      log0 = Log(1# + x)
   Else
     term = x / (2# + x)
     log0 = 2# * term * logcf(term * term, 1#, 2#)
   End If
End Function

Private Function log1(ByVal x As Double) As Double
'//Accurate calculation of log(1+x)-x, particularly for small x.
   Dim term As Double, Y  As Double
   If (Abs(x) < 0.01) Then
      term = x / (2# + x)
      Y = term * term
      log1 = term * ((((2# / 9# * Y + 2# / 7#) * Y + 0.4) * Y + 2# / 3#) * Y - x)
   ElseIf (x < minLog1Value Or x > 1#) Then
      log1 = Log(1# + x) - x
   Else
      term = x / (2# + x)
      Y = term * term
      log1 = term * (2# * Y * logcf(Y, 3#, 2#) - x)
   End If
End Function

Private Function logfbitdif(ByVal x As Double) As Double
'//Calculation of logfbit(x)-logfbit(1+x). x must be > -1.
  Dim Y As Double, y2 As Double
  Y = 1# / (2# * x + 3#)
  y2 = Y * Y
  logfbitdif = y2 * logcf(y2, 3#, 2#)
End Function

Private Function logfbit(ByVal x As Double) As Double
'//Error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
'//Are we ever concerned about the relative error involved in this function? I don't think so.
  Dim x1 As Double, x2 As Double, x3 As Double
  If (x >= 10000000000#) Then
     logfbit = lfbc1 / (x + 1#)
  ElseIf (x >= 6#) Then                      ' Abramowitz & Stegun's series 6.1.41
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = x2 * (lfbc6 - x2 * (lfbc7 - x2 * (lfbc8 - x2 * lfbc9)))
     x3 = x2 * (lfbc4 - x2 * (lfbc5 - x3))
     x3 = x2 * (lfbc2 - x2 * (lfbc3 - x3))
     logfbit = lfbc1 * (1# - x3) / x1
  ElseIf (x = 5#) Then
     logfbit = 1.38761288230707E-02                         ' 1.3876128823070747998745727023763E-02  ' calculated to give exact factorials
  ElseIf (x = 4#) Then
     logfbit = 1.66446911898212E-02                         ' 1.6644691189821192163194865373593E-02
  ElseIf (x = 3#) Then
     logfbit = 2.07906721037651E-02                         ' 2.0790672103765093111522771767849E-02
  ElseIf (x = 2#) Then
     logfbit = 2.76779256849983E-02                         ' 2.7677925684998339148789292746245E-02
  ElseIf (x = 1#) Then
     logfbit = 4.13406959554093E-02                         ' 4.1340695955409294093822081407118E-02
  ElseIf (x = 0#) Then
     logfbit = 8.10614667953273E-02                         ' 8.1061466795327258219670263594382E-02
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 6#)
        x2 = x2 + logfbitdif(x1)
        x1 = x1 + 1#
     Wend
     logfbit = x2 + logfbit(x1)
  Else
     logfbit = 1E+308
  End If
End Function

Private Function logdif(ByVal pr As Double, ByVal prob As Double) As Double
   Dim temp As Double
   temp = (pr - prob) / prob
   If Abs(temp) >= 0.5 Then
      logdif = Log(pr / prob)
   Else
      logdif = log0(temp)
   End If
End Function

Private Function CNormal(ByVal x As Double) As Double
'//Probability that a normal variate <= x
  Dim acc As Double, x2 As Double, D As Double, term As Double, A1 As Double, A2 As Double, b1 As Double, b2 As Double, C1 As Double, C2 As Double, c3 As Double

  If (Abs(x) < 1.5) Then
     acc = 0#
     x2 = x * x
     term = 1#
     D = 3#

     While (term > sumAcc * acc)

        D = D + 2#
        term = term * x2 / D
        acc = acc + term

     Wend

     acc = 1# + x2 / 3# * (1# + acc)
     CNormal = 0.5 + Exp(-x * x * 0.5 - lstpi) * x * acc
  ElseIf (Abs(x) > 40#) Then
     If (x > 0#) Then
        CNormal = 1#
     Else
        CNormal = 0#
     End If
  Else
     x2 = x * x
     A1 = 2#
     b1 = x2 + 5#
     C2 = x2 + 9#
     A2 = A1 * C2
     b2 = b1 * C2 - 12#
     C1 = 5#
     C2 = C2 + 4#

     While ((Abs(A2 * b1 - A1 * b2) > Abs(cfVSmall * b1 * A2)))

       c3 = C1 * (C1 + 1#)
       A1 = C2 * A2 - c3 * A1
       b1 = C2 * b2 - c3 * b1
       C1 = C1 + 2#
       C2 = C2 + 4#
       c3 = C1 * (C1 + 1#)
       A2 = C2 * A1 - c3 * A2
       b2 = C2 * b1 - c3 * b2
       C1 = C1 + 2#
       C2 = C2 + 4#
       If (b2 > scalefactor) Then
         A1 = A1 * scalefactor2
         b1 = b1 * scalefactor2
         A2 = A2 * scalefactor2
         b2 = b2 * scalefactor2
       End If

     Wend

     If (x > 0#) Then
        CNormal = 1# - Exp(-x * x * 0.5 - lstpi) * x / (x2 + 1# - A2 / b2)
     Else
        CNormal = -Exp(-x * x * 0.5 - lstpi) * x / (x2 + 1# - A2 / b2)
     End If

  End If
End Function

Private Function invcnormal(ByVal p As Double) As Double
'//Inverse of cnormal from AS241.
'//Require p to be strictly in the range 0..1

   Dim PPND16 As Double, q As Double, r As Double
   q = p - 0.5
   If (Abs(q) <= 0.425) Then
      r = 0.180625 - q * q
      PPND16 = q * (((((((a7 * r + a6) * r + a5) * r + a4) * r + A3) * r + A2) * r + A1) * r + A0) / (((((((b7 * r + b6) * r + b5) * r + b4) * r + b3) * r + b2) * r + b1) * r + 1#)
   Else
      If (q < 0#) Then
         r = p
      Else
         r = 1# - p
      End If
      r = Sqr(-Log(r))
      If (r <= 5#) Then
        r = r - 1.6
        PPND16 = (((((((c7 * r + c6) * r + c5) * r + c4) * r + c3) * r + C2) * r + C1) * r + C0) / (((((((d7 * r + d6) * r + d5) * r + d4) * r + d3) * r + d2) * r + D1) * r + 1#)
      Else
        r = r - 5#
        PPND16 = (((((((e7 * r + e6) * r + e5) * r + e4) * r + e3) * r + e2) * r + e1) * r + e0) / (((((((f7 * r + f6) * r + f5) * r + f4) * r + f3) * r + F2) * r + F1) * r + 1#)
      End If
      If (q < 0#) Then
         PPND16 = -PPND16
      End If
   End If
   invcnormal = PPND16
End Function

Private Function even(ByVal x As Double) As Boolean
   even = (Int(x / 2#) * 2# = x)
End Function

Private Function tdistexp(ByVal p As Double, ByVal q As Double, ByVal logqk2 As Double, ByVal k As Double, ByRef tdistDensity As Double) As Double
'//Special transformation of t-distribution useful for BinApprox.
'//Note approxtdistDens only used by binApprox if k > 100 or so.
   Dim Sum As Double, aki As Double, AI As Double, term As Double, q1 As Double, q8 As Double
   Dim C1 As Double, C2 As Double, A1 As Double, A2 As Double, b1 As Double, b2 As Double, cadd As Double
   Dim Result As Double, approxtdistDens As Double

   If (even(k)) Then
      approxtdistDens = Exp(logqk2 + logfbit(k - 1#) - 2# * logfbit(k * 0.5 - 1#) - lstpi)
   Else
      approxtdistDens = Exp(logqk2 + k * log1(1# / k) + 2# * logfbit((k - 1#) * 0.5) - logfbit(k - 1#) - lstpi)
   End If

   If (k * p < 4# * q) Then
     Sum = 0#
     aki = k + 1#
     AI = 3#
     term = 1#

     While (term > sumAcc * Sum)

        AI = AI + 2#
        aki = aki + 2#
        term = term * aki * p / AI
        Sum = Sum + term

     Wend

     Sum = 1# + (k + 1#) * p * (1# + Sum) / 3#
     Result = 0.5 - approxtdistDens * Sum * Sqr(k * p)
   ElseIf approxtdistDens = 0# Then
     Result = 0#
   Else
     q1 = 2# * (1# + q)
     q8 = 8# * q
     A1 = 1#
     b1 = (k - 3#) * p + 7#
     C1 = -20# * q
     A2 = (k - 5#) * p + 11#
     b2 = A2 * b1 + C1
     cadd = -30# * q
     C1 = -42# * q
     C2 = (k - 7#) * p + 15#

     While ((Abs(A2 * b1 - A1 * b2) > Abs(cfVSmall * b1 * A2)))

       A1 = C2 * A2 + C1 * A1
       b1 = C2 * b2 + C1 * b1
       C1 = C1 + cadd
       cadd = cadd - q8
       C2 = C2 + q1
       A2 = C2 * A1 + C1 * A2
       b2 = C2 * b1 + C1 * b2
       C1 = C1 + cadd
       cadd = cadd - q8
       C2 = C2 + q1
       If (Abs(b2) > scalefactor) Then
         A1 = A1 * scalefactor2
         b1 = b1 * scalefactor2
         A2 = A2 * scalefactor2
         b2 = b2 * scalefactor2
       ElseIf (Abs(b2) < scalefactor2) Then
         A1 = A1 * scalefactor
         b1 = b1 * scalefactor
         A2 = A2 * scalefactor
         b2 = b2 * scalefactor
       End If
     Wend

     Result = approxtdistDens * (1# - q / ((k - 1#) * p + 3# - 6# * q * A2 / b2)) / Sqr(k * p)
   End If
   tdistDensity = approxtdistDens * Sqr(q)
   tdistexp = Result
End Function

Private Function tdist(ByVal x As Double, ByVal k As Double, tdistDensity As Double) As Double
'//Probability that variate from t-distribution with k degress of freedom <= x
   Dim x2 As Double, k2 As Double, logterm As Double, A As Double, r As Double

   If Abs(x) >= Min(1#, k) Then
      k2 = k / x
      x2 = x + k2
      k2 = k2 / x2
      x2 = x / x2
   Else
      x2 = x * x
      k2 = k + x2
      x2 = x2 / k2
      k2 = k / k2
   End If
   If (k > 1E+30) Then
      tdist = CNormal(x)
      tdistDensity = Exp(-x * x / 2#)
   Else
      If (k2 < cSmall) Then
        logterm = k * 0.5 * (Log(k) - 2# * Log(Abs(x)))
      ElseIf (Abs(x2) < 0.5) Then
        logterm = k * 0.5 * log0(-x2)
      Else
        logterm = k * 0.5 * Log(k2)
      End If
      If (k >= 1#) Then
         If (x < 0#) Then
           tdist = tdistexp(x2, k2, logterm, k, tdistDensity)
         Else
           tdist = 1# - tdistexp(x2, k2, logterm, k, tdistDensity)
         End If
         Exit Function
      End If
      A = k / 2#
      tdistDensity = Exp(0.5 + (1# + 1# / k) * logterm + A * log0(-0.5 / (A + 1#)) + logfbit(A - 0.5) - logfbit(A)) * Sqr(A / ((1# + A))) * OneOverSqrTwoPi
      If (k2 < cSmall) Then
        r = (A + 1#) * log1(A / 1.5) + (logfbit(A + 0.5) - logfbit0p5) - lngammaexpansion(A)
        r = r + A * ((A - 0.5) / 1.5 + Log1p5 + (Log(k) - 2# * Log(Abs(x))))
        r = Exp(r) * (0.25 / (A + 0.5))
        If x < 0# Then
           tdist = r
        Else
           tdist = 1# - r
        End If
      ElseIf (x < 0#) Then
        If x2 < k2 Then
          tdist = 0.5 * compbeta(x2, 0.5, A)
        Else
          tdist = 0.5 * beta(k2, A, 0.5)
        End If
      Else
        If x2 < k2 Then
          tdist = 0.5 * (1# + beta(x2, 0.5, A))
        Else
          tdist = 0.5 * (1# + compbeta(k2, A, 0.5))
        End If
      End If
   End If
End Function

Private Function BetterThanTailApprox(ByVal prob As Double, ByVal df As Double) As Boolean
If df <= 2 Then
   BetterThanTailApprox = prob > 0.25 * Exp((1# - df) * 1.78514841051368)
ElseIf df <= 5 Then
   BetterThanTailApprox = prob > 0.045 * Exp((2# - df) * 1.30400766847605)
ElseIf df <= 20 Then
   BetterThanTailApprox = prob > 0.0009 * Exp((5# - df) * 0.921034037197618)
Else
   BetterThanTailApprox = prob > 0.0000000009 * Exp((20# - df) * 0.690775527898214)
End If
End Function

Private Function invtdist(ByVal prob As Double, ByVal df As Double) As Double
'//Inverse of tdist
'//Require prob to be in the range 0..1 df should be positive
  Dim xn As Double, xn2 As Double, tp As Double, tpDif As Double, tprob As Double, A As Double, pr As Double, lpr As Double, small As Double, smalllpr As Double, tdistDensity As Double
  If prob > 0.5 Then
     pr = 1# - prob
  Else
     pr = prob
  End If
  lpr = -Log(pr)
  small = 0.00000000000001
  smalllpr = small * lpr * pr
  If pr >= 0.5 Or df >= 1# And BetterThanTailApprox(pr, df) Then
'// Will divide by 0 if tp so small that tdistDensity underflows. Not a problem if prob > cSmall
     xn = invcnormal(pr)
     xn2 = xn * xn
'//Initial approximation is given in http://digital.library.adelaide.edu.au/coll/special//fisher/281.pdf. The modified NR correction then gets it right.
     tp = (((((27# * xn2 + 339#) * xn2 + 930#) * xn2 - 1782#) * xn2 - 765#) * xn2 + 17955#) / (368640# * df)
     tp = (tp + ((((79# * xn2 + 776#) * xn2 + 1482#) * xn2 - 1920#) * xn2 - 945#) / 92160#) / df
     tp = (tp + (((3# * xn2 + 19#) * xn2 + 17#) * xn2 - 15#) / 384#) / df
     tp = (tp + ((5# * xn2 + 16) * xn2 + 3#) / 96#) / df
     tp = (tp + (xn2 + 1#) / 4#) / df
     tp = xn * (1# + tp)
     tprob = 0#
     tpDif = 1# + Abs(tp)
  ElseIf df < 1# Then
     A = df / 2#
     tp = (A + 1#) * log1(A / 1.5) + (logfbit(A + 0.5) - logfbit0p5) - lngammaexpansion(A)
     tp = ((A - 0.5) / 1.5 + Log1p5 + Log(df)) / 2# + (tp - Log(4# * pr * (A + 0.5))) / df
     tp = -Exp(tp)
     tprob = tdist(tp, df, tdistDensity)
     If tdistDensity < nearly_zero Then
        tpDif = 0#
     Else
        tpDif = (tprob / tdistDensity) * log0((tprob - pr) / pr)
        tp = tp - tpDif
     End If
  Else
     tp = tdist(0, df, tdistDensity) 'Marginally quicker to get tdistDensity for integral df
     tp = Exp(-Log(Sqr(df) * pr / tdistDensity) / df)
     If df >= 2 Then
        tp = -Sqr(df * (tp * tp - 1#))
     Else
        tp = -Sqr(df) * Sqr(tp - 1#) * Sqr(tp + 1#)
     End If
     tpDif = tp / df
     tpDif = -log0((0.5 - 1# / (df + 2)) / (1# + tpDif * tp)) * (tpDif + 1# / tp)
     tp = tp - tpDif
     tprob = tdist(tp, df, tdistDensity)
     If tdistDensity < nearly_zero Then
        tpDif = 0#
     Else
        tpDif = (tprob / tdistDensity) * log0((tprob - pr) / pr)
        tp = tp - tpDif
     End If
  End If
  While (Abs(tprob - pr) > smalllpr And Abs(tpDif) > small * (1# + Abs(tp)))
     tprob = tdist(tp, df, tdistDensity)
     tpDif = (tprob / tdistDensity) * log0((tprob - pr) / pr)
     tp = tp - tpDif
  Wend
  invtdist = tp
  If prob > 0.5 Then invtdist = -invtdist
End Function

Private Function poissonTerm(ByVal i As Double, ByVal n As Double, ByVal diffFromMean As Double, ByVal logAdd As Double) As Double
'//Probability that poisson variate with mean n has value i (diffFromMean = n-i)
   Dim C2 As Double, c3 As Double
   Dim logpoissonTerm As Double, C1 As Double

   If ((i <= -1#) Or (n < 0#)) Then
      If (i = 0#) Then
         poissonTerm = Exp(logAdd)
      Else
         poissonTerm = 0#
      End If
   ElseIf ((i < 0#) And (n = 0#)) Then
      poissonTerm = [#VALUE!]
   Else
     c3 = i
     C2 = c3 + 1#
     C1 = (diffFromMean - 1#) / C2

     If (C1 < minLog1Value) Then
        If (i = 0#) Then
          logpoissonTerm = -n
          poissonTerm = Exp(logpoissonTerm + logAdd)
        ElseIf (n = 0#) Then
          poissonTerm = 0#
        Else
          logpoissonTerm = (c3 * Log(n / C2) - (diffFromMean - 1#)) - logfbit(c3)
          poissonTerm = Exp(logpoissonTerm + logAdd) / Sqr(C2) * OneOverSqrTwoPi
        End If
     Else
       logpoissonTerm = c3 * log1(C1) - C1 - logfbit(c3)
       poissonTerm = Exp(logpoissonTerm + logAdd) / Sqr(C2) * OneOverSqrTwoPi
     End If
   End If
End Function

Private Function poisson1(ByVal i As Double, ByVal n As Double, ByVal diffFromMean As Double) As Double
'//Probability that poisson variate with mean n has value <= i (diffFromMean = n-i)
'//For negative values of i (used for calculating the cumlative gamma distribution) there's a really nasty interpretation!
'//1-gamma(n,i) is calculated as poisson1(-i,n,0) since we need an accurate version of i rather than i-1.
'//Uses a simplified version of Legendre's continued fraction.
   Dim prob As Double, exact As Boolean
   If ((i >= 0#) And (n <= 0#)) Then
      exact = True
      prob = 1#
   ElseIf ((i > -1#) And (n <= 0#)) Then
      exact = True
      prob = 0#
   ElseIf ((i > -1#) And (i < 0#)) Then
      i = -i
      exact = False
      prob = poissonTerm(i, n, n - i, 0#) * i / n
      i = i - 1#
      diffFromMean = n - i
   Else
      exact = ((i <= -1#) Or (n < 0#))
      prob = poissonTerm(i, n, diffFromMean, 0#)
   End If
   If (exact Or prob = 0#) Then
      poisson1 = prob
      Exit Function
   End If

   Dim A1 As Double, A2 As Double, b1 As Double, b2 As Double, C1 As Double, C2 As Double, c3 As Double, c4 As Double, cfValue As Double
   Dim njj As Long, Numb As Long
   Dim sumAlways As Long, sumFactor As Long
   sumAlways = 0
   sumFactor = 6
   A1 = 0#
   If (i > sumAlways) Then
      Numb = Int(sumFactor * Exp(Log(n) / 3))
      Numb = max(0, Int(Numb - diffFromMean))
      If (Numb > i) Then
         Numb = Int(i)
      End If
   Else
      Numb = max(0, Int(i))
   End If

   b1 = 1#
   A2 = i - Numb
   b2 = diffFromMean + (Numb + 1#)
   C1 = 0#
   C2 = A2
   c4 = b2
   If C2 < 0# Then
      cfValue = cfVSmall
   Else
      cfValue = cfSmall
   End If
   While ((Abs(A2 * b1 - A1 * b2) > Abs(cfValue * b1 * A2)))

       C1 = C1 + 1#
       C2 = C2 - 1#
       c3 = C1 * C2
       c4 = c4 + 2#
       A1 = c4 * A2 + c3 * A1
       b1 = c4 * b2 + c3 * b1
       C1 = C1 + 1#
       C2 = C2 - 1#
       c3 = C1 * C2
       c4 = c4 + 2#
       A2 = c4 * A1 + c3 * A2
       b2 = c4 * b1 + c3 * b2
       If (b2 > scalefactor) Then
         A1 = A1 * scalefactor2
         b1 = b1 * scalefactor2
         A2 = A2 * scalefactor2
         b2 = b2 * scalefactor2
       End If
       If C2 < 0# And cfValue > cfVSmall Then
          cfValue = cfVSmall
       End If
   Wend

   A1 = A2 / b2

   C1 = i - Numb + 1#
   For njj = 1 To Numb
     A1 = (1# + A1) * (C1 / n)
     C1 = C1 + 1#
   Next njj

   poisson1 = (1# + A1) * prob
End Function

Private Function poisson2(ByVal i As Double, ByVal n As Double, ByVal diffFromMean As Double) As Double
'//Probability that poisson variate with mean n has value >= i (diffFromMean = n-i)
   Dim prob As Double, exact As Boolean
   If ((i <= 0#) And (n <= 0#)) Then
      exact = True
      prob = 1#
   Else
      exact = False
      prob = poissonTerm(i, n, diffFromMean, 0#)
   End If
   If (exact Or prob = 0#) Then
      poisson2 = prob
      Exit Function
   End If

   Dim A1 As Double, A2 As Double, b1 As Double, b2 As Double, C1 As Double, C2 As Double
   Dim njj As Long, Numb As Long
   Const sumFactor = 6
   Numb = Int(sumFactor * Exp(Log(n) / 3))
   Numb = max(0, Int(diffFromMean + Numb))

   A1 = 0#
   b1 = 1#
   A2 = n
   b2 = (Numb + 1#) - diffFromMean
   C1 = 0#
   C2 = b2

   While ((Abs(A2 * b1 - A1 * b2) > Abs(cfSmall * b1 * A2)))

      C1 = C1 + n
      C2 = C2 + 1#
      A1 = C2 * A2 + C1 * A1
      b1 = C2 * b2 + C1 * b1
      C1 = C1 + n
      C2 = C2 + 1#
      A2 = C2 * A1 + C1 * A2
      b2 = C2 * b1 + C1 * b2
      If (b2 > scalefactor) Then
        A1 = A1 * scalefactor2
        b1 = b1 * scalefactor2
        A2 = A2 * scalefactor2
        b2 = b2 * scalefactor2
      End If
   Wend

   A1 = A2 / b2

   C1 = i + Numb
   For njj = 1 To Numb
     A1 = (1# + A1) * (n / C1)
     C1 = C1 - 1#
   Next

   poisson2 = (1# + A1) * prob

End Function

Private Function poissonApprox(ByVal j As Double, ByVal diffFromMean As Double, ByVal comp As Boolean) As Double
'//Asymptotic expansion to calculate the probability that poisson variate has value <= j (diffFromMean = mean-j). If comp then calulate 1-probability.
'//cf. http://members.aol.com/iandjmsmith/PoissonApprox.htm
Dim pt As Double, s2pt As Double, res1 As Double, res2 As Double, elfb As Double, term As Double
Dim ig2 As Double, ig3 As Double, ig4 As Double, ig5 As Double, ig6 As Double, ig7 As Double, ig8 As Double
Dim ig05 As Double, ig25 As Double, ig35 As Double, ig45 As Double, ig55 As Double, ig65 As Double, ig75 As Double

pt = -log1(diffFromMean / j)
s2pt = Sqr(2# * j * pt)

ig2 = 1# / j + pt
term = pt * pt * 0.5
ig3 = ig2 / j + term
term = term * pt / 3#
ig4 = ig3 / j + term
term = term * pt / 4#
ig5 = ig4 / j + term
term = term * pt / 5#
ig6 = ig5 / j + term
term = term * pt / 6#
ig7 = ig6 / j + term
term = term * pt / 7#
ig8 = ig7 / j + term

ig05 = CNormal(-s2pt)
term = pt * twoThirds
ig25 = 1# / j + term
term = term * pt * twoFifths
ig35 = ig25 / j + term
term = term * pt * twoSevenths
ig45 = ig35 / j + term
term = term * pt * twoNinths
ig55 = ig45 / j + term
term = term * pt * twoElevenths
ig65 = ig55 / j + term
term = term * pt * twoThirteenths
ig75 = ig65 / j + term

elfb = ((((((coef75 / j + coef65) / j + coef55) / j + coef45) / j + coef35) / j + coef25) / j + coef15) + j
res1 = (((((((ig8 * coef8 + ig7 * coef7) + ig6 * coef6) + ig5 * coef5) + ig4 * coef4) + ig3 * coef3) + ig2 * coef2) + coef1) * Sqr(j)
res2 = ((((((ig75 * coef75 + ig65 * coef65) + ig55 * coef55) + ig45 * coef45) + ig35 * coef35) + ig25 * coef25) + coef15) * s2pt

If (comp) Then
   If (diffFromMean < 0#) Then
      poissonApprox = ig05 - (res1 - res2) * Exp(-j * pt - lstpi) / elfb
   Else
      poissonApprox = (1# - ig05) - (res1 + res2) * Exp(-j * pt - lstpi) / elfb
   End If
ElseIf (diffFromMean < 0#) Then
   poissonApprox = (1# - ig05) + (res1 - res2) * Exp(-j * pt - lstpi) / elfb
Else
   poissonApprox = ig05 + (res1 + res2) * Exp(-j * pt - lstpi) / elfb
End If
End Function

Private Function CPoisson(ByVal k As Double, ByVal LAMBDA As Double, ByVal dfm As Double) As Double
'//Probability that poisson variate with mean lambda has value <= k (diffFromMean = lambda-k) calculated by various methods.
   If ((k >= 21#) And (Abs(dfm) < (0.3 * k))) Then
      CPoisson = poissonApprox(k, dfm, False)
   ElseIf ((LAMBDA > k) And (LAMBDA >= 1#)) Then
      CPoisson = poisson1(k, LAMBDA, dfm)
   Else
      CPoisson = 1# - poisson2(k + 1#, LAMBDA, dfm - 1#)
   End If
End Function

Private Function comppoisson(ByVal k As Double, ByVal LAMBDA As Double, ByVal dfm As Double) As Double
'//Probability that poisson variate with mean lambda has value > k (diffFromMean = lambda-k) calculated by various methods.
   If ((k >= 21#) And (Abs(dfm) < (0.3 * k))) Then
      comppoisson = poissonApprox(k, dfm, True)
   ElseIf ((LAMBDA > k) And (LAMBDA >= 1#)) Then
      comppoisson = 1# - poisson1(k, LAMBDA, dfm)
   Else
      comppoisson = poisson2(k + 1#, LAMBDA, dfm - 1#)
   End If
End Function

Private Function invpoisson(ByVal k As Double, ByVal prob As Double) As Double
'//Inverse of poisson. Calculates mean such that poisson(k,mean,mean-k)=prob.
'//Require prob to be in the range 0..1, k should be -1/2 or non-negative
   If (k = 0#) Then
      invpoisson = -Log(prob + 9.99988867182683E-321)
   ElseIf (prob > 0.5) Then
      invpoisson = invcomppoisson(k, 1# - prob)
   Else '/*if (k > 0#)*/ then
      Dim temp2 As Double, xp As Double, dfm As Double, q As Double, qdif As Double, lpr As Double, small As Double, smalllpr As Double
      lpr = -Log(prob)
      small = 0.00000000000001
      smalllpr = small * lpr * prob
      xp = invcnormal(prob)
      dfm = 0.5 * xp * (xp - Sqr(4# * k + xp * xp))
      q = -1#
      qdif = -dfm
      If Abs(qdif) < 1# Then
         qdif = 1#
      ElseIf (k > 1E+50) Then
         invpoisson = k
         Exit Function
      End If
      While ((Abs(q - prob) > smalllpr) And (Abs(qdif) > (1# + Abs(dfm)) * small))
         q = CPoisson(k, k + dfm, dfm)
         If (q = 0#) Then
             qdif = qdif / 2#
             dfm = dfm + qdif
             q = -1#
         Else
            temp2 = poissonTerm(k, k + dfm, dfm, 0#)
            If (temp2 = 0#) Then
               qdif = qdif / 2#
               dfm = dfm + qdif
               q = -1#
            Else
               qdif = -2# * q * logdif(q, prob) / (1# + Sqr(Log(prob) / Log(q))) / temp2
               If (qdif > k + dfm) Then
                  qdif = dfm / 2#
                  dfm = dfm - qdif
                  q = -1#
               Else
                  dfm = dfm - qdif
               End If
            End If
         End If
      Wend
      invpoisson = k + dfm
   End If
End Function

Private Function invcomppoisson(ByVal k As Double, ByVal prob As Double) As Double
'//Inverse of comppoisson. Calculates mean such that comppoisson(k,mean,mean-k)=prob.
'//Require prob to be in the range 0..1, k should be -1/2 or non-negative
   If (prob > 0.5) Then
      invcomppoisson = invpoisson(k, 1# - prob)
   ElseIf (k = 0#) Then
      invcomppoisson = -log0(-prob)
   Else '/*if (k > 0#)*/ then
      Dim temp2 As Double, xp As Double, dfm As Double, q As Double, qdif As Double, LAMBDA As Double, qdifset As Boolean, lpr As Double, small As Double, smalllpr As Double
      lpr = -Log(prob)
      small = 0.00000000000001
      smalllpr = small * lpr * prob
      xp = invcnormal(prob)
      dfm = 0.5 * xp * (xp + Sqr(4# * k + xp * xp))
      LAMBDA = k + dfm
      If ((LAMBDA < 1#) And (k < 40#)) Then
         LAMBDA = Exp(Log(prob / poissonTerm(k + 1#, 1#, -k, 0#)) / (k + 1#))
         dfm = LAMBDA - k
      ElseIf (k > 1E+50) Then
         invcomppoisson = LAMBDA
         Exit Function
      End If
      q = -1#
      qdif = LAMBDA
      qdifset = False
      While ((Abs(q - prob) > smalllpr) And (Abs(qdif) > Min(LAMBDA, Abs(dfm)) * small))
         q = comppoisson(k, LAMBDA, dfm)
         If (q = 0#) Then
            If qdifset Then
               qdif = qdif / 2#
               dfm = dfm + qdif
               LAMBDA = LAMBDA + qdif
            Else
               LAMBDA = 2# * LAMBDA
               qdif = qdif * 2#
               dfm = LAMBDA - k
            End If
            q = -1#
         Else
            temp2 = poissonTerm(k, LAMBDA, dfm, 0#)
            If (temp2 = 0#) Then
               If qdifset Then
                  qdif = qdif / 2#
                  dfm = dfm + qdif
                  LAMBDA = LAMBDA + qdif
               Else
                  LAMBDA = 2# * LAMBDA
                  qdif = qdif * 2#
                  dfm = LAMBDA - k
               End If
               q = -1#
            Else
               qdif = 2# * q * logdif(q, prob) / (1# + Sqr(Log(prob) / Log(q))) / temp2
               If (qdif > LAMBDA) Then
                  LAMBDA = LAMBDA / 10#
                  qdif = dfm
                  dfm = LAMBDA - k
                  qdif = qdif - dfm
                  q = -1#
               Else
                  LAMBDA = LAMBDA - qdif
                  dfm = dfm - qdif
               End If
               qdifset = True
            End If
         End If
         If (Abs(dfm) > LAMBDA) Then
            dfm = LAMBDA - k
         Else
            LAMBDA = k + dfm
         End If
      Wend
      invcomppoisson = LAMBDA
   End If
End Function

Private Function binomialTerm(ByVal i As Double, ByVal j As Double, ByVal p As Double, ByVal q As Double, ByVal diffFromMean As Double, ByVal logAdd As Double) As Double
'//Probability that binomial variate with sample size i+j and event prob p (=1-q) has value i (diffFromMean = (i+j)*p-i)
   Dim C1 As Double, C2 As Double, c3 As Double
   Dim c4 As Double, c5 As Double, c6 As Double, ps As Double, logbinomialTerm As Double, dfm As Double
   If ((i = 0#) And (j <= 0#)) Then
      binomialTerm = Exp(logAdd)
   ElseIf ((i <= -1#) Or (j < 0#)) Then
      binomialTerm = 0#
   Else
      C1 = (i + 1#) + j
      If (p < q) Then
         C2 = i
         c3 = j
         ps = p
         dfm = diffFromMean
      Else
         c3 = i
         C2 = j
         ps = q
         dfm = -diffFromMean
      End If

      c5 = (dfm - (1# - ps)) / (C2 + 1#)
      c6 = -(dfm + ps) / (c3 + 1#)

      If (c5 < minLog1Value) Then
         If (C2 = 0#) Then
            logbinomialTerm = c3 * log0(-ps)
            binomialTerm = Exp(logbinomialTerm + logAdd)
         ElseIf ((ps = 0#) And (C2 > 0#)) Then
            binomialTerm = 0#
         Else
            c4 = logfbit(i + j) - logfbit(i) - logfbit(j)
            logbinomialTerm = c4 + C2 * (Log((ps * C1) / (C2 + 1#)) - c5) - c5 + c3 * log1(c6) - c6
            binomialTerm = Exp(logbinomialTerm + logAdd) * Sqr(C1 / ((C2 + 1#) * (c3 + 1#))) * OneOverSqrTwoPi
         End If
      Else
         c4 = logfbit(i + j) - logfbit(i) - logfbit(j)
         logbinomialTerm = c4 + (C2 * log1(c5) - c5) + (c3 * log1(c6) - c6)
         binomialTerm = Exp(logbinomialTerm + logAdd) * Sqr((C1 / (c3 + 1#)) / ((C2 + 1#))) * OneOverSqrTwoPi
      End If
   End If
End Function

Private Function binomialcf(ByVal ii As Double, ByVal jj As Double, ByVal pp As Double, ByVal qq As Double, ByVal diffFromMean As Double, ByVal comp As Boolean) As Double
'//Probability that binomial variate with sample size ii+jj and event prob pp (=1-qq) has value <=i (diffFromMean = (ii+jj)*pp-ii). If comp the returns 1 - probability.
Dim prob As Double, p As Double, q As Double, A1 As Double, A2 As Double, b1 As Double, b2 As Double
Dim C1 As Double, C2 As Double, c3 As Double, c4 As Double, N1 As Double, q1 As Double, dfm As Double
Dim i As Double, j As Double, ni As Double, nj As Double, Numb As Double, ip1 As Double, cfValue As Double
Dim swapped As Boolean, exact As Boolean

  If ((ii > -1#) And (ii < 0#)) Then
     ip1 = -ii
     ii = ip1 - 1#
  Else
     ip1 = ii + 1#
  End If
  N1 = (ii + 3#) + jj
  If ii < 0# Then
     cfValue = cfVSmall
     swapped = False
  ElseIf pp > qq Then
     cfValue = cfSmall
     swapped = N1 * qq >= jj + 1#
  Else
     cfValue = cfSmall
     swapped = N1 * pp <= ii + 2#
  End If
  If Not swapped Then
    i = ii
    j = jj
    p = pp
    q = qq
    dfm = diffFromMean
  Else
    j = ip1
    ip1 = jj
    i = jj - 1#
    p = qq
    q = pp
    dfm = 1# - diffFromMean
  End If
  If ((i > -1#) And ((j <= 0#) Or (p = 0#))) Then
     exact = True
     prob = 1#
  ElseIf ((i > -1#) And (i < 0#) Or (i = -1#) And (ip1 > 0#)) Then
     exact = False
     prob = binomialTerm(ip1, j, p, q, (ip1 + j) * p - ip1, 0#) * ip1 / ((ip1 + j) * p)
     dfm = (i + j) * p - i
  Else
     exact = ((i = 0#) And (j <= 0#)) Or ((i <= -1#) Or (j < 0#))
     prob = binomialTerm(i, j, p, q, dfm, 0#)
  End If
  If (exact) Or (prob = 0#) Then
     If (swapped = comp) Then
        binomialcf = prob
     Else
        binomialcf = 1# - prob
     End If
     Exit Function
  End If

  Dim sumAlways As Long, sumFactor As Long
  sumAlways = 0
  sumFactor = 6
  A1 = 0#
  If (i > sumAlways) Then
     Numb = Int(sumFactor * Sqr(p + 0.5) * Exp(Log(N1 * p * q) / 3))
     Numb = Int(Numb - dfm)
     If (Numb > i) Then
        Numb = Int(i)
     End If
  Else
     Numb = Int(i)
  End If
  If (Numb < 0#) Then
     Numb = 0#
  End If

  b1 = 1#
  q1 = q + 1#
  A2 = (i - Numb) * q
  b2 = dfm + Numb + 1#
  C1 = 0#
  C2 = A2
  c4 = b2
  While ((Abs(A2 * b1 - A1 * b2) > Abs(cfValue * b1 * A2)))

    C1 = C1 + 1#
    C2 = C2 - q
    c3 = C1 * C2
    c4 = c4 + q1
    A1 = c4 * A2 + c3 * A1
    b1 = c4 * b2 + c3 * b1
    C1 = C1 + 1#
    C2 = C2 - q
    c3 = C1 * C2
    c4 = c4 + q1
    A2 = c4 * A1 + c3 * A2
    b2 = c4 * b1 + c3 * b2
    If (Abs(b2) > scalefactor) Then
      A1 = A1 * scalefactor2
      b1 = b1 * scalefactor2
      A2 = A2 * scalefactor2
      b2 = b2 * scalefactor2
    ElseIf (Abs(b2) < scalefactor2) Then
      A1 = A1 * scalefactor
      b1 = b1 * scalefactor
      A2 = A2 * scalefactor
      b2 = b2 * scalefactor
    End If
    If C2 < 0# And cfValue > cfVSmall Then
       cfValue = cfVSmall
    End If
  Wend
  A1 = A2 / b2

  ni = (i - Numb + 1#) * q
  nj = (j + Numb) * p
  While (Numb > 0#)
     A1 = (1# + A1) * (ni / nj)
     ni = ni + q
     nj = nj - p
     Numb = Numb - 1#
  Wend

  A1 = (1# + A1) * prob
  If (swapped = comp) Then
     binomialcf = A1
  Else
     binomialcf = 1# - A1
  End If

End Function

Private Function binApprox(ByVal A As Double, ByVal B As Double, ByVal diffFromMean As Double, ByVal comp As Boolean) As Double
'//Asymptotic expansion to calculate the probability that binomial variate has value <= a (diffFromMean = (a+b)*p-a). If comp then calulate 1-probability.
'//cf. http://members.aol.com/iandjmsmith/BinomialApprox.htm
Dim n As Double, N1 As Double
Dim pq1 As Double, mfac As Double, res As Double, tp As Double, lval As Double, lvv As Double, temp As Double
Dim ib05 As Double, ib15 As Double, ib25 As Double, ib35 As Double, ib45 As Double, ib55 As Double, ib65 As Double
Dim ib2 As Double, ib3 As Double, ib4 As Double, ib5 As Double, ib6 As Double, ib7 As Double
Dim elfb As Double, coef15 As Double, coef25 As Double, coef35 As Double, coef45 As Double, coef55 As Double, coef65 As Double
Dim coef2 As Double, coef3 As Double, coef4 As Double, coef5 As Double, coef6 As Double, coef7 As Double
Dim tdistDensity As Double, approxtdistDens As Double

n = A + B
N1 = n + 1#
lvv = (B + diffFromMean) / N1 - diffFromMean
lval = (A * log1(lvv / A) + B * log1(-lvv / B)) / n
tp = -expm1(lval)

pq1 = (A / n) * (B / n)

coef15 = (-17# * pq1 + 2#) / 24#
coef25 = ((-503# * pq1 + 76#) * pq1 + 4#) / 1152#
coef35 = (((-315733# * pq1 + 53310#) * pq1 + 8196#) * pq1 - 1112#) / 414720#
coef45 = (4059192# + pq1 * (15386296# - 85262251# * pq1))
coef45 = (-9136# + pq1 * (-697376 + pq1 * coef45)) / 39813120#
coef55 = (3904584040# + pq1 * (10438368262# - 55253161559# * pq1))
coef55 = (5244128# + pq1 * (-43679536# + pq1 * (-703410640# + pq1 * coef55))) / 6688604160#
coef65 = (-3242780782432# + pq1 * (18320560326516# + pq1 * (38020748623980# - 194479285104469# * pq1)))
coef65 = (335796416# + pq1 * (61701376704# + pq1 * (-433635420336# + pq1 * coef65))) / 4815794995200#
elfb = (((((coef65 / ((n + 6.5) * pq1) + coef55) / ((n + 5.5) * pq1) + coef45) / ((n + 4.5) * pq1) + coef35) / ((n + 3.5) * pq1) + coef25) / ((n + 2.5) * pq1) + coef15) / ((n + 1.5) * pq1) + 1#

coef2 = (-pq1 - 2#) / 135#
coef3 = ((-44# * pq1 - 86#) * pq1 + 4#) / 2835#
coef4 = (((-404# * pq1 - 786#) * pq1 + 48#) * pq1 + 8#) / 8505#
coef5 = (((((-2421272# * pq1 - 4721524#) * pq1 + 302244#) * pq1) + 118160#) * pq1 - 4496#) / 12629925#
coef6 = ((((((-473759128# * pq1 - 928767700#) * pq1 + 57300188#) * pq1) + 38704888#) * pq1 - 1870064#) * pq1 - 167072#) / 492567075#
coef7 = (((((((-8530742848# * pq1 - 16836643200#) * pq1 + 954602040#) * pq1) + 990295352#) * pq1 - 44963088#) * pq1 - 11596512#) * pq1 + 349376#) / 1477701225#

ib05 = tdistexp(tp, 1# - tp, N1 * lval, 2# * N1, tdistDensity)
mfac = N1 * tp
ib15 = Sqr(2# * mfac)

If (mfac > 1E+50) Then
   ib2 = (1# + mfac) / (n + 2#)
   mfac = mfac * tp / 2#
   ib3 = (ib2 + mfac) / (n + 3#)
   mfac = mfac * tp / 3#
   ib4 = (ib3 + mfac) / (n + 4#)
   mfac = mfac * tp / 4#
   ib5 = (ib4 + mfac) / (n + 5#)
   mfac = mfac * tp / 5#
   ib6 = (ib5 + mfac) / (n + 6#)
   mfac = mfac * tp / 6#
   ib7 = (ib6 + mfac) / (n + 7#)
   res = (ib2 * coef2 + (ib3 * coef3 + (ib4 * coef4 + (ib5 * coef5 + (ib6 * coef6 + ib7 * coef7 / pq1) / pq1) / pq1) / pq1) / pq1) / pq1

   mfac = (n + 1.5) * tp * twoThirds
   ib25 = (1# + mfac) / (n + 2.5)
   mfac = mfac * tp * twoFifths
   ib35 = (ib25 + mfac) / (n + 3.5)
   mfac = mfac * tp * twoSevenths
   ib45 = (ib35 + mfac) / (n + 4.5)
   mfac = mfac * tp * twoNinths
   ib55 = (ib45 + mfac) / (n + 5.5)
   mfac = mfac * tp * twoElevenths
   ib65 = (ib55 + mfac) / (n + 6.5)
   temp = (((((coef65 * ib65 / pq1 + coef55 * ib55) / pq1 + coef45 * ib45) / pq1 + coef35 * ib35) / pq1 + coef25 * ib25) / pq1 + coef15)
Else
   ib2 = 1# + mfac
   mfac = mfac * (n + 2#) * tp / 2#
   ib3 = ib2 + mfac
   mfac = mfac * (n + 3#) * tp / 3#
   ib4 = ib3 + mfac
   mfac = mfac * (n + 4#) * tp / 4#
   ib5 = ib4 + mfac
   mfac = mfac * (n + 5#) * tp / 5#
   ib6 = ib5 + mfac
   mfac = mfac * (n + 6#) * tp / 6#
   ib7 = ib6 + mfac
   res = (ib2 * coef2 + (ib3 * coef3 + (ib4 * coef4 + (ib5 * coef5 + (ib6 * coef6 + ib7 * coef7 / ((n + 7#) * pq1)) / ((n + 6#) * pq1)) / ((n + 5#) * pq1)) / ((n + 4#) * pq1)) / ((n + 3#) * pq1)) / ((n + 2#) * pq1)

   mfac = (n + 1.5) * tp * twoThirds
   ib25 = 1# + mfac
   mfac = mfac * (n + 2.5) * tp * twoFifths
   ib35 = ib25 + mfac
   mfac = mfac * (n + 3.5) * tp * twoSevenths
   ib45 = ib35 + mfac
   mfac = mfac * (n + 4.5) * tp * twoNinths
   ib55 = ib45 + mfac
   mfac = mfac * (n + 5.5) * tp * twoElevenths
   ib65 = ib55 + mfac
   temp = (((((coef65 * ib65 / ((n + 6.5) * pq1) + coef55 * ib55) / ((n + 5.5) * pq1) + coef45 * ib45) / ((n + 4.5) * pq1) + coef35 * ib35) / ((n + 3.5) * pq1) + coef25 * ib25) / ((n + 2.5) * pq1) + coef15)
End If

approxtdistDens = tdistDensity / Sqr(1# - tp)
temp = ib15 * temp / ((n + 1.5) * pq1)
res = (oneThird + res) * 2# * (A - B) / (n * Sqr(N1 * pq1))
If (comp) Then
   If (lvv > 0#) Then
      binApprox = ib05 - (res - temp) * approxtdistDens / elfb
   Else
      binApprox = (1# - ib05) - (res + temp) * approxtdistDens / elfb
   End If
ElseIf (lvv > 0#) Then
   binApprox = (1# - ib05) + (res - temp) * approxtdistDens / elfb
Else
   binApprox = ib05 + (res + temp) * approxtdistDens / elfb
End If
End Function

Private Function binomial(ByVal ii As Double, ByVal jj As Double, ByVal pp As Double, ByVal qq As Double, ByVal diffFromMean As Double) As Double
'//Probability that binomial variate with sample size ii+jj and event prob pp (=1-qq) has value <=i (diffFromMean = (ii+jj)*pp-ii).
   Dim mij As Double
   mij = Min(ii, jj)
   If ((mij > 50#) And (Abs(diffFromMean) < (0.1 * mij))) Then
      binomial = binApprox(jj - 1#, ii, diffFromMean, False)
   Else
      binomial = binomialcf(ii, jj, pp, qq, diffFromMean, False)
   End If
End Function

Private Function compbinomial(ByVal ii As Double, ByVal jj As Double, ByVal pp As Double, ByVal qq As Double, ByVal diffFromMean As Double) As Double
'//Probability that binomial variate with sample size ii+jj and event prob pp (=1-qq) has value >i (diffFromMean = (ii+jj)*pp-ii).
   Dim mij As Double
   mij = Min(ii, jj)
   If ((mij > 50#) And (Abs(diffFromMean) < (0.1 * mij))) Then
       compbinomial = binApprox(jj - 1#, ii, diffFromMean, True)
   Else
       compbinomial = binomialcf(ii, jj, pp, qq, diffFromMean, True)
   End If
End Function

Private Function invbinom(ByVal k As Double, ByVal m As Double, ByVal prob As Double, ByRef oneMinusP As Double) As Double
'//Inverse of binomial. Delivers event probability p (q held in oneMinusP in case required) so that binomial(k,m,p,oneMinusp,dfm) = prob.
'//Note that dfm is calculated accurately but never made available outside of this routine.
'//Require prob to be in the range 0..1, m should be positive and k should be >= 0
   Dim temp1 As Double, temp2 As Double
   If (prob > 0.5) Then
      temp2 = invcompbinom(k, m, 1# - prob, oneMinusP)
   Else
      temp1 = invcompbinom(m - 1#, k + 1#, prob, oneMinusP)
      temp2 = oneMinusP
      oneMinusP = temp1
   End If
   invbinom = temp2
End Function

Private Function invcompbinom(ByVal k As Double, ByVal m As Double, ByVal prob As Double, ByRef oneMinusP As Double) As Double
'//Inverse of compbinomial. Delivers event probability p (q held in oneMinusP in case required) so that compbinomial(k,m,p,oneMinusp,dfm) = prob.
'//Note that dfm is calculated accurately but never made available outside of this routine.
'//Require prob to be in the range 0..1, m should be positive and k should be >= -0.5
Dim xp As Double, xp2 As Double, dfm As Double, n As Double, p As Double, q As Double, pr As Double, dif As Double, temp As Double, temp2 As Double, Result As Double, lpr As Double, small As Double, smalllpr As Double, nminpq As Double
   Result = -1#
   n = k + m
   If (prob > 0.5) Then
      Result = invbinom(k, m, 1# - prob, oneMinusP)
   ElseIf (k = 0#) Then
      Result = log0(-prob) / n
      If (Abs(Result) < 1#) Then
        Result = -expm1(Result)
        oneMinusP = 1# - Result
      Else
        oneMinusP = Exp(Result)
        Result = 1# - oneMinusP
      End If
   ElseIf (m = 1#) Then
      Result = Log(prob) / n
      If (Abs(Result) < 1#) Then
        oneMinusP = -expm1(Result)
        Result = 1# - oneMinusP
      Else
        Result = Exp(Result)
        oneMinusP = 1# - Result
      End If
   Else
      pr = -1#
      xp = invcnormal(prob)
      xp2 = xp * xp
      temp = 2# * xp * Sqr(k * (m / n) + xp2 / 4#)
      xp2 = xp2 / n
      dfm = (xp2 * (m - k) + temp) / (2# * (1# + xp2))
      If (k + dfm < 0#) Then
         dfm = -k
      End If
      q = (m - dfm) / n
      p = (k + dfm) / n
      dif = -dfm / n
      If (dif = 0#) Then
         dif = 1#
      ElseIf Min(k, m) > 1E+50 Then
         oneMinusP = q
         invcompbinom = p
         Exit Function
      End If
      lpr = -Log(prob)
      small = 0.00000000000004
      smalllpr = small * lpr * prob
      nminpq = n * Min(p, q)
      While ((Abs(pr - prob) > smalllpr) And (n * Abs(dif) > Min(Abs(dfm), nminpq) * small))
         pr = compbinomial(k, m, p, q, dfm)
         If (pr < nearly_zero) Then '/*Should not be happenning often */
            dif = dif / 2#
            dfm = dfm + n * dif
            p = p + dif
            q = q - dif
            pr = -1#
         Else
            temp2 = binomialTerm(k, m, p, q, dfm, 0#) * m / q
            If (temp2 < nearly_zero) Then '/*Should not be happenning often */
               dif = dif / 2#
               dfm = dfm + n * dif
               p = p + dif
               q = q - dif
               pr = -1#
            Else
               dif = 2# * pr * logdif(pr, prob) / (1# + Sqr(Log(prob) / Log(pr))) / temp2
               If (q + dif <= 0#) Then '/*not v. good */
                  dif = -0.9999 * q
                  dfm = dfm - n * dif
                  p = p - dif
                  q = q + dif
                  pr = -1#
               ElseIf (p - dif <= 0#) Then '/*v. good */
                  temp = Exp(Log(prob / pr) / (k + 1#))
                  dif = p
                  p = temp * p
                  dif = p - dif
                  dfm = n * p - k
                  q = 1# - p
                  pr = -1#
               Else
                  dfm = dfm - n * dif
                  p = p - dif
                  q = q + dif
               End If
            End If
         End If
      Wend
      Result = p
      oneMinusP = q
   End If
   invcompbinom = Result
End Function

Private Function abMinuscd(ByVal A As Double, ByVal B As Double, ByVal c As Double, ByVal D As Double) As Double
   Dim A1 As Double, b1 As Double, C1 As Double, D1 As Double, A2 As Double, b2 As Double, C2 As Double, d2 As Double, R1 As Double, r2 As Double, r3 As Double
   A2 = Int(A / twoTo27) * twoTo27
   A1 = A - A2
   b2 = Int(B / twoTo27) * twoTo27
   b1 = B - b2
   C2 = Int(c / twoTo27) * twoTo27
   C1 = c - C2
   d2 = Int(D / twoTo27) * twoTo27
   D1 = D - d2
   R1 = A1 * b1 - C1 * D1
   r2 = (A2 * b1 - C1 * d2) + (A1 * b2 - C2 * D1)
   r3 = A2 * b2 - C2 * d2
   If (r3 < 0#) = (r2 < 0#) Then
      abMinuscd = r3 + (r2 + R1)
   Else
      abMinuscd = (r3 + r2) + R1
   End If
End Function

Private Function aTimes2Powerb(ByVal A As Double, ByVal B As Integer) As Double
   If B > 709 Then
      A = (A * scalefactor) * scalefactor
      B = B - 512
   ElseIf B < -709 Then
      A = (A * scalefactor2) * scalefactor2
      B = B + 512
   End If
   aTimes2Powerb = A * (2#) ^ B
End Function

Private Function GeneralabMinuscd(ByVal A As Double, ByVal B As Double, ByVal c As Double, ByVal D As Double) As Double
   Dim s As Double, CA As Double, CB As Double, CC As Double, cd As Double
   Dim L2 As Integer, PA As Integer, PB As Integer, pc As Integer, pd As Integer
   s = A * B - c * D
   If A <= 0# Or B <= 0# Or c <= 0# Or D <= 0# Then
      GeneralabMinuscd = s
      Exit Function
   ElseIf s < 0# Then
      GeneralabMinuscd = -GeneralabMinuscd(c, D, A, B)
      Exit Function
   End If
   L2 = Int(Log(A) / Log(2))
   PA = 51 - L2
   CA = aTimes2Powerb(A, PA)
   L2 = Int(Log(B) / Log(2))
   PB = 51 - L2
   CB = aTimes2Powerb(B, PB)
   L2 = Int(Log(c) / Log(2))
   pc = 51 - L2
   CC = aTimes2Powerb(c, pc)
   pd = PA + PB - pc
   cd = aTimes2Powerb(D, pd)
   GeneralabMinuscd = aTimes2Powerb(abMinuscd(CA, CB, CC, cd), -(PA + PB))
End Function

Private Function hypergeometricTerm(ByVal AI As Double, ByVal aji As Double, ByVal aki As Double, ByVal amkji As Double) As Double
'// Probability that hypergeometric variate from a population with total type Is of aki+ai, total type IIs of amkji+aji, has ai type Is and aji type IIs selected.
   Dim aj As Double, am As Double, ak As Double, amj As Double, amk As Double
   Dim cjkmi As Double, ai1 As Double, aj1 As Double, ak1 As Double, am1 As Double, aki1 As Double, aji1 As Double, amk1 As Double, amj1 As Double, amkji1 As Double
   Dim C1 As Double, c3 As Double, c4 As Double, c5 As Double, loghypergeometricTerm As Double

   ak = aki + AI
   amk = amkji + aji
   aj = aji + AI
   am = amk + ak
   amj = amkji + aki
   If (am > max_discrete) Then
      hypergeometricTerm = [#VALUE!]
      Exit Function
   End If
   If ((AI = 0#) And ((aji <= 0#) Or (aki <= 0#) Or (amj < 0#) Or (amk < 0#))) Then
      hypergeometricTerm = 1#
   ElseIf ((AI > 0#) And (Min(aki, aji) = 0#) And (max(amj, amk) = 0#)) Then
      hypergeometricTerm = 1#
   ElseIf ((AI >= 0#) And (amkji > -1#) And (aki > -1#) And (aji >= 0#)) Then
     C1 = logfbit(amkji) + logfbit(aki) + logfbit(aji) + logfbit(am) + logfbit(AI)
     C1 = logfbit(amk) + logfbit(ak) + logfbit(aj) + logfbit(amj) - C1
     ai1 = AI + 1#
     aj1 = aj + 1#
     ak1 = ak + 1#
     am1 = am + 1#
     aki1 = aki + 1#
     aji1 = aji + 1#
     amk1 = amk + 1#
     amj1 = amj + 1#
     amkji1 = amkji + 1#
     cjkmi = GeneralabMinuscd(aji, aki, AI, amkji)
     c5 = (cjkmi - AI) / (amkji1 * am1)
     If (c5 < minLog1Value) Then
        c3 = amkji * (Log((amj1 * amk1) / (amkji1 * am1)) - c5) - c5
     Else
        c3 = amkji * log1(c5) - c5
     End If

     c5 = (-cjkmi - aji) / (aki1 * am1)
     If (c5 < minLog1Value) Then
        c4 = aki * (Log((ak1 * amj1) / (aki1 * am1)) - c5) - c5
     Else
        c4 = aki * log1(c5) - c5
     End If

     c3 = c3 + c4
     c5 = (-cjkmi - aki) / (aji1 * am1)
     If (c5 < minLog1Value) Then
        c4 = aji * (Log((aj1 * amk1) / (aji1 * am1)) - c5) - c5
     Else
        c4 = aji * log1(c5) - c5
     End If

     c3 = c3 + c4
     c5 = (cjkmi - amkji) / (ai1 * am1)
     If (c5 < minLog1Value) Then
        c4 = AI * (Log((aj1 * ak1) / (ai1 * am1)) - c5) - c5
     Else
        c4 = AI * log1(c5) - c5
     End If

     c3 = c3 + c4
     loghypergeometricTerm = (C1 + 1# / am1) + c3

     hypergeometricTerm = Exp(loghypergeometricTerm) * Sqr((amk1 * ak1) * (aj1 * amj1) / ((amkji1 * aki1 * aji1) * (am1 * ai1))) * OneOverSqrTwoPi
   Else
     hypergeometricTerm = 0#
   End If

End Function

Private Function hypergeometric(ByVal AI As Double, ByVal aji As Double, ByVal aki As Double, ByVal amkji As Double, ByVal comp As Boolean, ByRef ha1 As Double, ByRef hprob As Double, ByRef hswap As Boolean) As Double
'// Probability that hypergeometric variate from a population with total type Is of aki+ai, total type IIs of amkji+aji, has up to ai type Is selected in a sample of size aji+ai.
     Dim prob As Double
     Dim A1 As Double, A2 As Double, b1 As Double, b2 As Double, an As Double, bn As Double, bnAdd As Double
     Dim C1 As Double, C2 As Double, c3 As Double, c4 As Double
     Dim i As Double, ji As Double, ki As Double, mkji As Double, njj As Double, Numb As Double, maxSums As Double, swapped As Boolean
     Dim ip1 As Double, must_do_cf As Boolean, allIntegral As Boolean, exact As Boolean
     If (amkji > -1#) And (amkji < 0#) Then
        ip1 = -amkji
        mkji = ip1 - 1#
        allIntegral = False
     Else
        ip1 = amkji + 1#
        mkji = amkji
        allIntegral = AI = Int(AI) And aji = Int(aji) And aki = Int(aki) And mkji = Int(mkji)
     End If

     If allIntegral Then
        swapped = (AI + 0.5) * (mkji + 0.5) >= (aki - 0.5) * (aji - 0.5)
     ElseIf AI < 100# And AI = Int(AI) Or mkji < 0# Then
        swapped = (AI + 0.5) * (mkji + 0.5) >= aki * aji + 1000#
     ElseIf AI < 1# Then
        swapped = (AI + 0.5) * (mkji + 0.5) >= aki * aji
     ElseIf aji < 1# Or aki < 1# Or (AI < 1# And AI > 0#) Then
        swapped = False
     Else
        swapped = (AI + 0.5) * (mkji + 0.5) >= (aki - 0.5) * (aji - 0.5)
     End If
     If Not swapped Then
       i = AI
       ji = aji
       ki = aki
     Else
       i = aji - 1#
       ji = AI + 1#
       ki = ip1
       ip1 = aki
       mkji = aki - 1#
     End If
     C2 = ji + i
     c4 = mkji + ki + C2
     If (c4 > max_discrete) Then
        hypergeometric = [#VALUE!]
        Exit Function
     End If
     If ((i >= 0#) And (ji <= 0#) Or (ki <= 0#) Or (ip1 + ki <= 0#) Or (ip1 + ji <= 0#)) Then
        exact = True
        If (i >= 0#) Then
           prob = 1#
        Else
           prob = 0#
        End If
     ElseIf (ip1 > 0#) And (ip1 < 1#) Then
        exact = False
        prob = hypergeometricTerm(i, ji, ki, ip1) * (ip1 * (c4 + 1#)) / ((ki + ip1) * (ji + ip1))
     Else
        exact = ((i = 0#) And ((ji <= 0#) Or (ki <= 0#) Or (mkji + ki < 0#) Or (mkji + ji < 0#))) Or ((i > 0#) And (Min(ki, ji) = 0#) And (max(mkji + ki, mkji + ji) = 0#))
        prob = hypergeometricTerm(i, ji, ki, mkji)
     End If
     hprob = prob
     hswap = swapped
     ha1 = 0#

     If (exact) Or (prob = 0#) Then
        If (swapped = comp) Then
           hypergeometric = prob
        Else
           hypergeometric = 1# - prob
        End If
        Exit Function
     End If

     A1 = 0#
     Dim sumAlways As Long, sumFactor As Long
     sumAlways = 0#
     sumFactor = 10#

     If i < mkji Then
        must_do_cf = i <> Int(i)
        maxSums = Int(i)
     Else
        must_do_cf = mkji <> Int(mkji)
        maxSums = Int(max(mkji, 0#))
     End If
     If must_do_cf Then
        sumAlways = 0#
        sumFactor = 5#
     Else
        sumAlways = 20#
        sumFactor = 10#
     End If
     If (maxSums > sumAlways Or must_do_cf) Then
        Numb = Int(sumFactor / c4 * Exp(Log((ki + i) * (ji + i) * (ip1 + ji) * (ip1 + ki)) / 3#))
        Numb = Int(i - (ki + i) * (ji + i) / c4 + Numb)
        If (Numb < 0#) Then
           Numb = 0#
        ElseIf Numb > maxSums Then
           Numb = maxSums
        End If
     Else
        Numb = maxSums
     End If

     If (2# * Numb <= maxSums Or must_do_cf) Then
        b1 = 1#
        C1 = 0#
        C2 = i - Numb
        c3 = mkji - Numb
        A2 = C2 * c3
        c3 = c3 - 1#
        b2 = GeneralabMinuscd(ki + Numb + 1#, ji + Numb + 1#, C2 - 1#, c3)
        bn = b2
        bnAdd = c3 + c4 + C2 - 2#
        While (b2 > 0# And (Abs(A2 * b1 - A1 * b2) > Abs(cfVSmall * b1 * A2)))
            C1 = C1 + 1#
            C2 = C2 - 1#
            an = (C1 * C2) * (c3 * c4)
            c3 = c3 - 1#
            c4 = c4 - 1#
            bn = bn + bnAdd
            bnAdd = bnAdd - 4#
            A1 = bn * A2 + an * A1
            b1 = bn * b2 + an * b1
            If (b1 > scalefactor) Then
              A1 = A1 * scalefactor2
              b1 = b1 * scalefactor2
              A2 = A2 * scalefactor2
              b2 = b2 * scalefactor2
            End If
            C1 = C1 + 1#
            C2 = C2 - 1#
            an = (C1 * C2) * (c3 * c4)
            c3 = c3 - 1#
            c4 = c4 - 1#
            bn = bn + bnAdd
            bnAdd = bnAdd - 4#
            A2 = bn * A1 + an * A2
            b2 = bn * b1 + an * b2
            If (b2 > scalefactor) Then
              A1 = A1 * scalefactor2
              b1 = b1 * scalefactor2
              A2 = A2 * scalefactor2
              b2 = b2 * scalefactor2
            End If
        Wend
        If b1 < 0# Or b2 < 0# Then
           hypergeometric = [#VALUE!]
           Exit Function
        Else
           A1 = A2 / b2
        End If
     Else
        Numb = maxSums
     End If

     C1 = i - Numb + 1#
     C2 = mkji - Numb + 1#
     c3 = ki + Numb
     c4 = ji + Numb
     For njj = 1 To Numb
       A1 = (1# + A1) * ((C1 * C2) / (c3 * c4))
       C1 = C1 + 1#
       C2 = C2 + 1#
       c3 = c3 - 1#
       c4 = c4 - 1#
     Next njj

     ha1 = A1
     A1 = (1# + A1) * prob
     If (swapped = comp) Then
        hypergeometric = A1
     Else
        If A1 > 0.99 Then
           hypergeometric = [#VALUE!]
        Else
           hypergeometric = 1# - A1
        End If
     End If
End Function

Private Function compgfunc(ByVal x As Double, ByVal A As Double) As Double
'//Calculates a*x(1/(a+1) - x/2*(1/(a+2) - x/3*(1/(a+3) - ...)))
'//Mainly for calculating the complement of gamma(x,a) for small a and x <= 1.
'//a should be close to 0, x >= 0 & x <=1
  Dim term As Double, D As Double, Sum As Double
  term = x
  D = 2#
  Sum = term / (A + 1#)
  While (Abs(term) > Abs(Sum * sumAcc))
      term = -term * x / D
      Sum = Sum + term / (A + D)
      D = D + 1#
  Wend
  compgfunc = A * Sum
End Function

Private Function lngammaexpansion(ByVal A As Double) As Double
'//Calculates log(gamma(a+1)) accurately for for small a (0 < a & a < 0.5).
'//Uses Abramowitz & Stegun's series 6.1.33
'//Mainly for calculating the complement of gamma(x,a) for small a and x <= 1.
'//
Dim coeffs As Variant                       ' "Variant" rather than  "coefs(40) as Double"  to permit use of Array assignment
'// for i < 40 coeffs[i] holds (zeta(i+2)-1)/(i+2), coeffs[40] holds (zeta(i+2)-1)
coeffs = Array( _
0.322467033424113, 6.73523010531981E-02, 2.05808084277845E-02, 7.38555102867399E-03, _
2.89051033074152E-03, 1.19275391170326E-03, 5.09669524743042E-04, 2.23154758453579E-04, _
9.94575127818085E-05, 4.49262367381331E-05, 2.05072127756707E-05, 9.4394882752684E-06, _
4.37486678990749E-06, 2.03921575380137E-06, 9.55141213040742E-07, 4.49246919876457E-07, _
2.12071848055547E-07, 1.00432248239681E-07, 4.76981016936398E-08, 2.27110946089432E-08, _
1.0838659214897E-08, 5.18347504197005E-09, 2.48367454380248E-09, 1.19214014058609E-09, _
5.73136724167886E-10, 2.75952288512423E-10, 1.33047643742445E-10, 6.4229645638381E-11, _
3.10442477473223E-11, 1.50213840807541E-11, 7.27597448023908E-12, 3.52774247657592E-12, _
1.71199179055962E-12, 8.31538584142029E-13, 4.04220052528944E-13, 1.96647563109662E-13, _
9.57363038783856E-14, 4.66407602642837E-14, 2.27373696006597E-14, 1.10913994708345E-14, _
2.27373684582465E-13)
Dim lgam As Double
Dim i As Integer
lgam = coeffs(40) * logcf(-A / 2#, 42#, 1#)
For i = 39 To 0 Step -1
   lgam = (coeffs(i) - A * lgam)
Next i
lngammaexpansion = (A * lgam - eulers_const) * A - log1(A)
End Function

Private Function incgamma(ByVal x As Double, ByVal A As Double, ByVal comp As Boolean) As Double
'//Calculates gamma-cdf for small a (complementary gamma-cdf if comp).
   Dim r As Double
   r = A * Log(x) - lngammaexpansion(A)
   If (comp) Then
      r = -expm1(r)
      incgamma = r + compgfunc(x, A) * (1# - r)
   Else
      incgamma = Exp(r) * (1# - compgfunc(x, A))
   End If
End Function

Private Function invincgamma(ByVal A As Double, ByVal prob As Double, ByVal comp As Boolean) As Double
'//Calculates inverse of gamma for small a (inverse of complementary gamma if comp).
Dim GA As Double, x As Double, deriv As Double, z As Double, W As Double, dif As Double, pr As Double, lpr As Double, small As Double, smalllpr As Double
   If (prob > 0.5) Then
       invincgamma = invincgamma(A, 1# - prob, Not comp)
       Exit Function
   End If
   lpr = -Log(prob)
   small = 0.00000000000001
   smalllpr = small * lpr * prob
   If (comp) Then
      GA = -expm1(lngammaexpansion(A))
      x = -Log(prob * (1# - GA) / A)
      If (x < 0.5) Then
         pr = Exp(log0(-(GA + prob * (1# - GA))) / A)
         If (x < pr) Then
            x = pr
         End If
      End If
      dif = x
      pr = -1#
      While ((Abs(pr - prob) > smalllpr) And (Abs(dif) > small * max(cSmall, x)))
         deriv = poissonTerm(A, x, x - A, 0#) * A             'value of derivative is actually deriv/x but it can overflow when x is denormal...
         If (x > 1#) Then
            pr = poisson1(-A, x, 0#)
         Else
            z = compgfunc(x, A)
            W = -expm1(A * Log(x))
            W = z + W * (1# - z)
            pr = (W - GA) / (1# - GA)
         End If
         dif = x * (pr / deriv) * logdif(pr, prob) '...so multiply by x in slightly different order
         x = x + dif
         If (x < 0#) Then
            invincgamma = 0#
            Exit Function
         End If
      Wend
   Else
      GA = Exp(lngammaexpansion(A))
      x = Log(prob * GA)
      If (x < -711# * A) Then
         invincgamma = 0#
         Exit Function
      End If
      x = Exp(x / A)
      z = 1# - compgfunc(x, A)
      deriv = poissonTerm(A, x, x - A, 0#) * A / x
      pr = prob * z
      dif = (pr / deriv) * logdif(pr, prob)
      x = x - dif
      While ((Abs(pr - prob) > smalllpr) And (Abs(dif) > small * max(cSmall, x)))
         deriv = poissonTerm(A, x, x - A, 0#) * A / x
         If (x > 1#) Then
            pr = 1# - poisson1(-A, x, 0#)
         Else
            pr = (1# - compgfunc(x, A)) * Exp(A * Log(x)) / GA
         End If
         dif = (pr / deriv) * logdif(pr, prob)
         x = x - dif
      Wend
   End If
   invincgamma = x
End Function

Private Function GAMMA(ByVal n As Double, ByVal A As Double) As Double
'Assumes n > 0 & a >= 0.  Only called by (comp)gamma_nc with a = 0.
   If (A = 0#) Then
      GAMMA = 1#
   ElseIf ((A < 1#) And (n < 1#)) Then
      GAMMA = incgamma(n, A, False)
   ElseIf (A >= 1#) Then
      GAMMA = comppoisson(A - 1#, n, n - A + 1#)
   Else
      GAMMA = 1# - poisson1(-A, n, 0#)
   End If
End Function

Private Function compgamma(ByVal n As Double, ByVal A As Double) As Double
'Assumes n > 0 & a >= 0. Only called by (comp)gamma_nc with a = 0.
   If (A = 0#) Then
      compgamma = 0#
   ElseIf ((A < 1#) And (n < 1#)) Then
      compgamma = incgamma(n, A, True)
   ElseIf (A >= 1#) Then
      compgamma = CPoisson(A - 1#, n, n - A + 1#)
   Else
      compgamma = poisson1(-A, n, 0#)
   End If
End Function

Private Function invgamma(ByVal A As Double, ByVal prob As Double) As Double
'//Inverse of gamma(x,a)
   If (A >= 1#) Then
      invgamma = invcomppoisson(A - 1#, prob)
   Else
      invgamma = invincgamma(A, prob, False)
   End If
End Function

Private Function invcompgamma(ByVal A As Double, ByVal prob As Double) As Double
'//Inverse of compgamma(x,a)
   If (A >= 1#) Then
      invcompgamma = invpoisson(A - 1#, prob)
   Else
      invcompgamma = invincgamma(A, prob, True)
   End If
End Function

Private Function logfbit1dif(ByVal x As Double) As Double
'// Calculation of logfbit1(x)-logfbit1(1+x).
  logfbit1dif = (logfbitdif(x) - 0.25 / ((x + 1#) * (x + 2#))) / (x + 1.5)
End Function

Private Function logfbit1(ByVal x As Double) As Double
'// Derivative of error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1 As Double, x2 As Double
  If (x >= 10000000000#) Then
     logfbit1 = -lfbc1 * ((x + 1#) ^ -2)
  ElseIf (x >= 6#) Then
     Dim x3 As Double
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = (11# * lfbc6 - x2 * (13# * lfbc7 - x2 * (15# * lfbc8 - x2 * 17# * lfbc9)))
     x3 = (5# * lfbc3 - x2 * (7# * lfbc4 - x2 * (9# * lfbc5 - x2 * x3)))
     x3 = x2 * (3# * lfbc2 - x2 * x3)
    logfbit1 = -lfbc1 * (1# - x3) * x2
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 6#)
        x2 = x2 + logfbit1dif(x1)
        x1 = x1 + 1#
     Wend
     logfbit1 = x2 + logfbit1(x1)
  Else
     logfbit1 = -1E+308
  End If
End Function

Private Function logfbit3dif(ByVal x As Double) As Double
'// Calculation of logfbit3(x)-logfbit3(1+x).
  logfbit3dif = -(2# * x + 3#) * (((x + 1#) * (x + 2#)) ^ -3)
End Function

Private Function logfbit3(ByVal x As Double) As Double
'// Third derivative of error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1 As Double, x2 As Double
  If (x >= 10000000000#) Then
     logfbit3 = -0.5 * ((x + 1#) ^ -4)
  ElseIf (x >= 6#) Then
     Dim x3 As Double
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = x2 * (4080# * lfbc8 - x2 * 5814# * lfbc9)
     x3 = x2 * (1716# * lfbc6 - x2 * (2730# * lfbc7 - x3))
     x3 = x2 * (504# * lfbc4 - x2 * (990# * lfbc5 - x3))
     x3 = x2 * (60# * lfbc2 - x2 * (210# * lfbc3 - x3))
     logfbit3 = -lfbc1 * (6# - x3) * x2 * x2
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 6#)
        x2 = x2 + logfbit3dif(x1)
        x1 = x1 + 1#
     Wend
     logfbit3 = x2 + logfbit3(x1)
  Else
     logfbit3 = -1E+308
  End If
End Function

Private Function logfbit5dif(ByVal x As Double) As Double
'// Calculation of logfbit5(x)-logfbit5(1+x).
  logfbit5dif = -6# * (2# * x + 3#) * ((5# * x + 15#) * x + 12#) * (((x + 1#) * (x + 2#)) ^ -5)
End Function

Private Function logfbit5(ByVal x As Double) As Double
'// Fifth derivative of error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1 As Double, x2 As Double
  If (x >= 10000000000#) Then
     logfbit5 = -10# * ((x + 1#) ^ -6)
  ElseIf (x >= 6#) Then
     Dim x3 As Double
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = x2 * (1395360# * lfbc8 - x2 * 2441880# * lfbc9)
     x3 = x2 * (360360# * lfbc6 - x2 * (742560# * lfbc7 - x3))
     x3 = x2 * (55440# * lfbc4 - x2 * (154440# * lfbc5 - x3))
     x3 = x2 * (2520# * lfbc2 - x2 * (15120# * lfbc3 - x3))
     logfbit5 = -lfbc1 * (120# - x3) * x2 * x2 * x2
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 6#)
        x2 = x2 + logfbit5dif(x1)
        x1 = x1 + 1#
     Wend
     logfbit5 = x2 + logfbit5(x1)
  Else
     logfbit5 = -1E+308
  End If
End Function

Private Function logfbit7dif(ByVal x As Double) As Double
'// Calculation of logfbit7(x)-logfbit7(1+x).
  logfbit7dif = -120# * (2# * x + 3#) * ((((14# * x + 84#) * x + 196#) * x + 210#) * x + 87#) * (((x + 1#) * (x + 2#)) ^ -7)
End Function

Private Function logfbit7(ByVal x As Double) As Double
'// Seventh derivative of error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1 As Double, x2 As Double
  If (x >= 10000000000#) Then
     logfbit7 = -420# * ((x + 1#) ^ -8)
  ElseIf (x >= 6#) Then
     Dim x3 As Double
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = x2 * (586051200# * lfbc8 - x2 * 1235591280# * lfbc9)
     x3 = x2 * (98017920# * lfbc6 - x2 * (253955520# * lfbc7 - x3))
     x3 = x2 * (8648640# * lfbc4 - x2 * (32432400# * lfbc5 - x3))
     x3 = x2 * (181440# * lfbc2 - x2 * (1663200# * lfbc3 - x3))
     logfbit7 = -lfbc1 * (5040# - x3) * x2 * x2 * x2 * x2
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 6#)
        x2 = x2 + logfbit7dif(x1)
        x1 = x1 + 1#
     Wend
     logfbit7 = x2 + logfbit7(x1)
  Else
     logfbit7 = -1E+308
  End If
End Function

Private Function lfbaccdif(ByVal A As Double, ByVal B As Double) As Double
'// This is now always reasonably accurate, although it is not always required to be so when called from incbeta.
   If (A > 0.03 * (A + B + 1#)) Then
      lfbaccdif = logfbit(A + B) - logfbit(B)
   Else
      Dim A2 As Double, ab2 As Double
      A2 = A * A
      ab2 = A / 2# + B
      lfbaccdif = A * (logfbit1(ab2) + A2 / 24# * (logfbit3(ab2) + A2 / 80# * (logfbit5(ab2) + A2 / 168# * logfbit7(ab2))))
   End If
End Function

Private Function compbfunc(ByVal x As Double, ByVal A As Double, ByVal B As Double) As Double
'// Calculates a*(b-1)*x(1/(a+1) - (b-2)*x/2*(1/(a+2) - (b-3)*x/3*(1/(a+3) - ...)))
'// Mainly for calculating the complement of BETA(x,a,b) for small a and b*x < 1.
'// a should be close to 0, x >= 0 & x <=1 & b*x < 1
  Dim term As Double, D As Double, Sum As Double
  term = x
  D = 2#
  Sum = term / (A + 1#)
  While (Abs(term) > Abs(Sum * sumAcc))
      term = -term * (B - D) * x / D
      Sum = Sum + term / (A + D)
      D = D + 1#
  Wend
  compbfunc = A * (B - 1#) * Sum
End Function

Private Function incbeta(ByVal x As Double, ByVal A As Double, ByVal B As Double, ByVal comp As Boolean) As Double
'// Calculates BETA for small a (complementary BETA if comp).
   Dim r As Double
   If (x > 0.5) Then
      incbeta = incbeta(1# - x, B, A, Not comp)
   Else
      r = (A + B + 0.5) * log1(A / (1# + B)) + A * ((A - 0.5) / (1# + B) + Log((1# + B) * x)) - lngammaexpansion(A)
      If (comp) Then
         r = -expm1(r + lfbaccdif(A, B))
         r = r + compbfunc(x, A, B) * (1# - r)
         r = r + (A / (A + B)) * (1# - r)
      Else
         r = Exp(r + (logfbit(A + B) - logfbit(B))) * (1# - compbfunc(x, A, B)) * (B / (A + B))
      End If
      incbeta = r
   End If
End Function

Private Function beta(ByVal x As Double, ByVal A As Double, ByVal B As Double) As Double
'Assumes x >= 0 & a >= 0 & b >= 0. Only called with a = 0 or b = 0 by (comp)BETA_nc
   If (A = 0# And B = 0#) Then
      beta = [#VALUE!]
   ElseIf (A = 0#) Then
      beta = 1#
   ElseIf (B = 0#) Then
      beta = 0#
   ElseIf (x <= 0#) Then
      beta = 0#
   ElseIf (x >= 1#) Then
      beta = 1#
   ElseIf (A < 1# And B < 1#) Then
      beta = incbeta(x, A, B, False)
   ElseIf (A < 1# And (1# + B) * x <= 1#) Then
      beta = incbeta(x, A, B, False)
   ElseIf (B < 1# And A <= (1# + A) * x) Then
      beta = incbeta(1# - x, B, A, True)
   ElseIf (A < 1#) Then
      beta = compbinomial(-A, B, x, 1# - x, 0#)
   ElseIf (B < 1#) Then
      beta = binomial(-B, A, 1# - x, x, 0#)
   Else
      beta = compbinomial(A - 1#, B, x, 1# - x, (A + B - 1#) * x - A + 1#)
   End If
End Function

Private Function compbeta(ByVal x As Double, ByVal A As Double, ByVal B As Double) As Double
'Assumes x >= 0 & a >= 0 & b >= 0. Only called with a = 0 or b = 0 by (comp)BETA_nc
   If (A = 0# And B = 0#) Then
      compbeta = [#VALUE!]
   ElseIf (A = 0#) Then
      compbeta = 0#
   ElseIf (B = 0#) Then
      compbeta = 1#
   ElseIf (x <= 0#) Then
      compbeta = 1#
   ElseIf (x >= 1#) Then
      compbeta = 0#
   ElseIf (A < 1# And B < 1#) Then
      compbeta = incbeta(x, A, B, True)
   ElseIf (A < 1# And (1# + B) * x <= 1#) Then
      compbeta = incbeta(x, A, B, True)
   ElseIf (B < 1# And A <= (1# + A) * x) Then
      compbeta = incbeta(1# - x, B, A, False)
   ElseIf (A < 1#) Then
      compbeta = binomial(-A, B, x, 1# - x, 0#)
   ElseIf (B < 1#) Then
      compbeta = compbinomial(-B, A, 1# - x, x, 0#)
   Else
      compbeta = binomial(A - 1#, B, x, 1# - x, (A + B - 1#) * x - A + 1#)
   End If
End Function

Private Function invincbeta(ByVal A As Double, ByVal B As Double, ByVal prob As Double, ByVal comp As Boolean, ByRef oneMinusP As Double) As Double
'// Calculates inverse of BETA for small a (inverse of complementary BETA if comp).
Dim r As Double, rb As Double, x As Double, OneOverDeriv As Double, dif As Double, pr As Double, mnab As Double, aplusbOvermxab As Double, lpr As Double, small As Double, smalllpr As Double
   If (Not comp And prob > B / (A + B)) Then
       invincbeta = invincbeta(A, B, 1# - prob, Not comp, oneMinusP)
       Exit Function
   ElseIf (comp And prob > A / (A + B) And prob > 0.1) Then
       invincbeta = invincbeta(A, B, 1# - prob, Not comp, oneMinusP)
       Exit Function
   End If
   lpr = max(-Log(prob), 1#)
   small = 0.00000000000001
   smalllpr = small * lpr * prob
   If A >= B Then
      mnab = B
      aplusbOvermxab = (A + B) / A
   Else
      mnab = A
      aplusbOvermxab = (A + B) / B
   End If
   If (comp) Then
      r = (A + B + 0.5) * log1(A / (1# + B)) + A * (A - 0.5) / (1# + B) + lfbaccdif(A, B) - lngammaexpansion(A)
      r = -expm1(r)
      r = r + (A / (A + B)) * (1# - r)
      If (B < 1#) Then
         rb = (A + B + 0.5) * log1(B / (1# + A)) + B * (B - 0.5) / (1# + A) + (logfbit(A + B) - logfbit(A)) - lngammaexpansion(B)
         rb = Exp(rb) * (A / (A + B))
         oneMinusP = Log(prob / rb) / B
         If (oneMinusP < 0#) Then
             oneMinusP = Exp(oneMinusP) / (1# + A)
         Else
             oneMinusP = 0.5
         End If
         If (oneMinusP = 0#) Then
            invincbeta = 1#
            Exit Function
         ElseIf (oneMinusP > 0.5) Then
            oneMinusP = 0.5
         End If
         x = 1# - oneMinusP
         pr = rb * (1# - compbfunc(oneMinusP, B, A)) * Exp(B * Log((1 + A) * oneMinusP))
         OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(A, B, x, oneMinusP, (A + B) * x - A, 0#) * mnab)
         dif = OneOverDeriv * pr * logdif(pr, prob)
         oneMinusP = oneMinusP - dif
         x = 1# - oneMinusP
         If (oneMinusP <= 0#) Then
            oneMinusP = 0#
            invincbeta = 1#
            Exit Function
         ElseIf (x < 0.25) Then
            x = Exp(log0((r - prob) / (1# - r)) / A) / (B + 1#)
            oneMinusP = 1# - x
            If (x = 0#) Then
               invincbeta = 0#
               Exit Function
            End If
            pr = compbfunc(x, A, B) * (1# - prob)
            OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(A, B, x, oneMinusP, (A + B) * x - A, 0#) * mnab)
            dif = OneOverDeriv * (prob + pr) * log0(pr / prob)
            x = x + dif
            If (x <= 0#) Then
               oneMinusP = 1#
               invincbeta = 0#
               Exit Function
            End If
            oneMinusP = 1# - x
         End If
      Else
         pr = Exp(log0((r - prob) / (1# - r)) / A) / (B + 1#)
         x = Log(B * prob / (A * (1# - r) * B * Exp(A * Log(1 + B)))) / B
         If (Abs(x) < 0.5) Then
            x = -expm1(x)
            oneMinusP = 1# - x
         Else
            oneMinusP = Exp(x)
            x = 1# - oneMinusP
            If (oneMinusP = 0#) Then
               invincbeta = x
               Exit Function
            End If
         End If
         If pr > x And pr < 1# Then
            x = pr
            oneMinusP = 1# - x
         End If
      End If
      dif = Min(x, oneMinusP)
      pr = -1#
      While ((Abs(pr - prob) > smalllpr) And (Abs(dif) > small * max(cSmall, Min(x, oneMinusP))))
         If (B < 1# And x > 0.5) Then
            pr = rb * (1# - compbfunc(oneMinusP, B, A)) * Exp(B * Log((1# + A) * oneMinusP))
         ElseIf ((1# + B) * x > 1#) Then
            pr = binomial(-A, B, x, oneMinusP, 0#)
         Else
            pr = r + compbfunc(x, A, B) * (1# - r)
            pr = pr - expm1(A * Log((1# + B) * x)) * (1# - pr)
         End If
         OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(A, B, x, oneMinusP, (A + B) * x - A, 0#) * mnab)
         dif = OneOverDeriv * pr * logdif(pr, prob)
         If (x > 0.5) Then
            oneMinusP = oneMinusP - dif
            x = 1# - oneMinusP
            If (oneMinusP <= 0#) Then
               oneMinusP = 0#
               invincbeta = 1#
               Exit Function
            End If
         Else
            x = x + dif
            oneMinusP = 1# - x
            If (x <= 0#) Then
               oneMinusP = 1#
               invincbeta = 0#
               Exit Function
            End If
         End If
      Wend
   Else
      r = (A + B + 0.5) * log1(A / (1# + B)) + A * (A - 0.5) / (1# + B) + (logfbit(A + B) - logfbit(B)) - lngammaexpansion(A)
      r = Exp(r) * (B / (A + B))
      x = logdif(prob, r)
      If (x < -711# * A) Then
         x = 0#
      Else
         x = Exp(x / A) / (1# + B)
      End If
      If (x = 0#) Then
         oneMinusP = 1#
         invincbeta = x
         Exit Function
      ElseIf (x >= 0.5) Then
         x = 0.5
      End If
      oneMinusP = 1# - x
      pr = r * (1# - compbfunc(x, A, B)) * Exp(A * Log((1# + B) * x))
      OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(A, B, x, oneMinusP, (A + B) * x - A, 0#) * mnab)
      dif = OneOverDeriv * pr * logdif(pr, prob)
      x = x - dif
      oneMinusP = oneMinusP + dif
      While ((Abs(pr - prob) > smalllpr) And (Abs(dif) > small * max(cSmall, Min(x, oneMinusP))))
         If ((1# + B) * x > 1#) Then
            pr = compbinomial(-A, B, x, oneMinusP, 0#)
         ElseIf (x > 0.5) Then
            pr = incbeta(oneMinusP, B, A, Not comp)
         Else
            pr = r * (1# - compbfunc(x, A, B)) * Exp(A * Log((1# + B) * x))
         End If
         OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(A, B, x, oneMinusP, (A + B) * x - A, 0#) * mnab)
         dif = OneOverDeriv * pr * logdif(pr, prob)
         If x < 0.5 Then
            x = x - dif
            oneMinusP = 1# - x
         Else
            oneMinusP = oneMinusP + dif
            x = 1# - oneMinusP
         End If
      Wend
   End If
   invincbeta = x
End Function

Private Function invbeta(ByVal A As Double, ByVal B As Double, ByVal prob As Double, ByRef oneMinusP As Double) As Double
   Dim swap As Double
   If (prob = 0#) Then
      oneMinusP = 1#
      invbeta = 0#
   ElseIf (prob = 1#) Then
      oneMinusP = 0#
      invbeta = 1#
   ElseIf (A = B And prob = 0.5) Then
      oneMinusP = 0.5
      invbeta = 0.5
   ElseIf (A < B And B < 1#) Then
      invbeta = invincbeta(A, B, prob, False, oneMinusP)
   ElseIf (B < A And A < 1#) Then
      swap = invincbeta(B, A, prob, True, oneMinusP)
      invbeta = oneMinusP
      oneMinusP = swap
   ElseIf (A < 1#) Then
      invbeta = invincbeta(A, B, prob, False, oneMinusP)
   ElseIf (B < 1#) Then
      swap = invincbeta(B, A, prob, True, oneMinusP)
      invbeta = oneMinusP
      oneMinusP = swap
   Else
      invbeta = invcompbinom(A - 1#, B, prob, oneMinusP)
   End If
End Function

Private Function invcompbeta(ByVal A As Double, ByVal B As Double, ByVal prob As Double, ByRef oneMinusP As Double) As Double
   Dim swap As Double
   If (prob = 0#) Then
      oneMinusP = 0#
      invcompbeta = 1#
   ElseIf (prob = 1#) Then
      oneMinusP = 1#
      invcompbeta = 0#
   ElseIf (A = B And prob = 0.5) Then
      oneMinusP = 0.5
      invcompbeta = 0.5
   ElseIf (A < B And B < 1#) Then
      invcompbeta = invincbeta(A, B, prob, True, oneMinusP)
   ElseIf (B < A And A < 1#) Then
      swap = invincbeta(B, A, prob, False, oneMinusP)
      invcompbeta = oneMinusP
      oneMinusP = swap
   ElseIf (A < 1#) Then
      invcompbeta = invincbeta(A, B, prob, True, oneMinusP)
   ElseIf (B < 1#) Then
      swap = invincbeta(B, A, prob, False, oneMinusP)
      invcompbeta = oneMinusP
      oneMinusP = swap
   Else
      invcompbeta = invbinom(A - 1#, B, prob, oneMinusP)
   End If
End Function

Private Function critpoiss(ByVal Mean As Double, ByVal cprob As Double) As Double
'//i such that Pr(poisson(mean,i)) >= cprob and  Pr(poisson(mean,i-1)) < cprob
   If (cprob > 0.5) Then
      critpoiss = critcomppoiss(Mean, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   dfm = invcnormal(cprob) * Sqr(Mean)
   i = Int(Mean + dfm + 0.5)
   While (True)
      i = Int(i)
      If (i < 0#) Then
         i = 0#
      End If
      If (i >= max_crit) Then
         critpoiss = i
         Exit Function
      End If
      dfm = Mean - i
      pr = CPoisson(i, Mean, dfm)
      tpr = 0
      If (pr >= cprob) Then
         If (i = 0#) Then
            critpoiss = i
            Exit Function
         End If
         tpr = poissonTerm(i, Mean, dfm, 0#)
         pr = pr - tpr
         If (pr < cprob) Then
            critpoiss = i
            Exit Function
         End If

         i = i - 1#
         Dim temp As Double, temp2 As Double
         temp = (pr - cprob) / tpr
         If (temp > 10) Then
            temp = Int(temp + 0.5)
            i = i - temp
            temp2 = poissonTerm(i, Mean, Mean - i, 0#)
            i = i - temp * (tpr - temp2) / (2 * temp2)
         Else
            tpr = tpr * (i + 1#) / Mean
            pr = pr - tpr
            If (pr < cprob) Then
               critpoiss = i
               Exit Function
            End If
            i = i - 1#
            If (i = 0#) Then
               critpoiss = i
               Exit Function
            End If
            temp2 = (pr - cprob) / tpr
            If (temp2 < temp - 0.9) Then
               While (pr >= cprob)
                  tpr = tpr * (i + 1#) / Mean
                  pr = pr - tpr
                  i = i - 1#
               Wend
               critpoiss = i + 1#
               Exit Function
            Else
               temp = Int(Log(cprob / pr) / Log((i + 1#) / Mean) + 0.5)
               i = i - temp
               If (i < 0#) Then
                  i = 0#
               End If
               temp2 = poissonTerm(i, Mean, Mean - i, 0#)
               If (temp2 > nearly_zero) Then
                  temp = Log((cprob / pr) * (tpr / temp2)) / Log((i + 1#) / Mean)
                  i = i - temp
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            dfm = dfm - 1#
            tpr = poissonTerm(i, Mean, dfm, 0#)
            pr = pr + tpr
         Wend
         While (pr < cprob)
            i = i + 1#
            tpr = tpr * Mean / i
            pr = pr + tpr
         Wend
         critpoiss = i
         Exit Function
      End If
   Wend
End Function

Private Function critcomppoiss(ByVal Mean As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(poisson(mean,i)) > cprob and  1-Pr(poisson(mean,i-1)) <= cprob
   If (cprob > 0.5) Then
      critcomppoiss = critpoiss(Mean, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   dfm = invcnormal(cprob) * Sqr(Mean)
   i = Int(Mean - dfm + 0.5)
   While (True)
      i = Int(i)
      If (i >= max_crit) Then
         critcomppoiss = i
         Exit Function
      End If
      dfm = Mean - i
      pr = comppoisson(i, Mean, dfm)
      tpr = 0
      If (pr > cprob) Then
         i = i + 1#
         dfm = dfm - 1#
         tpr = poissonTerm(i, Mean, dfm, 0#)
         If (pr < (1.00001) * tpr) Then
            While (tpr > cprob)
               i = i + 1#
               tpr = tpr * Mean / i
            Wend
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcomppoiss = i
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = poissonTerm(i, Mean, Mean - i, 0#)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            ElseIf (pr / tpr > 0.00001) Then
               i = i + 1#
               tpr = tpr * Mean / i
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcomppoiss = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * Mean / i
                     pr = pr - tpr
                  Wend
                  critcomppoiss = i
                  Exit Function
               Else
                  temp = Log(cprob / pr) / Log(Mean / i)
                  temp = Int((Log(cprob / pr) - temp * Log(i / (temp + i))) / Log(Mean / i) + 0.5)
                  i = i + temp
                  temp2 = poissonTerm(i, Mean, Mean - i, 0#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(Mean / i)
                     i = i + temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = poissonTerm(i, Mean, dfm, 0#)
            pr = pr + tpr
            i = i - 1#
            dfm = dfm + 1#
         Wend
         While (pr <= cprob)
            tpr = tpr * (i + 1#) / Mean
            pr = pr + tpr
            i = i - 1#
         Wend
         critcomppoiss = i + 1#
         Exit Function
      End If
   Wend
End Function

Private Function critbinomial(ByVal n As Double, ByVal eprob As Double, ByVal cprob As Double) As Double
'//i such that Pr(binomial(n,eprob,i)) >= cprob and  Pr(binomial(n,eprob,i-1)) < cprob
   If (cprob > 0.5) Then
      critbinomial = critcompbinomial(n, eprob, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   dfm = invcnormal(cprob) * Sqr(n * eprob * (1# - eprob))
   i = n * eprob + dfm
   While (True)
      i = Int(i)
      If (i < 0#) Then
         i = 0#
      ElseIf (i > n) Then
         i = n
      End If
      If (i >= max_crit) Then
         critbinomial = i
         Exit Function
      End If
      dfm = n * eprob - i
      pr = binomial(i, n - i, eprob, 1# - eprob, dfm)
      tpr = 0#
      If (pr >= cprob) Then
         If (i = 0#) Then
            critbinomial = i
            Exit Function
         End If
         tpr = binomialTerm(i, n - i, eprob, 1# - eprob, dfm, 0#)
         If (pr < (1.00001) * tpr) Then
            tpr = tpr * ((i + 1#) * (1# - eprob)) / ((n - i) * eprob)
            i = i - 1#
            While (tpr >= cprob)
               tpr = tpr * ((i + 1#) * (1# - eprob)) / ((n - i) * eprob)
               i = i - 1
            Wend
         Else
            pr = pr - tpr
            If (pr < cprob) Then
               critbinomial = i
               Exit Function
            End If
            i = i - 1#
            If (i = 0#) Then
               critbinomial = i
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = binomialTerm(i, n - i, eprob, 1# - eprob, n * eprob - i, 0#)
               i = i - temp * (tpr - temp2) / (2# * temp2)
            Else
               tpr = tpr * ((i + 1#) * (1# - eprob)) / ((n - i) * eprob)
               pr = pr - tpr
               If (pr < cprob) Then
                  critbinomial = i
                  Exit Function
               End If
               i = i - 1#
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     tpr = tpr * ((i + 1#) * (1# - eprob)) / ((n - i) * eprob)
                     pr = pr - tpr
                     i = i - 1#
                  Wend
                  critbinomial = i + 1#
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((i + 1#) * (1# - eprob)) / ((n - i) * eprob)) + 0.5)
                  i = i - temp
                  If (i < 0#) Then
                     i = 0#
                  End If
                  temp2 = binomialTerm(i, n - i, eprob, 1# - eprob, n * eprob - i, 0#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((i + 1#) * (1# - eprob)) / ((n - i) * eprob))
                     i = i - temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            dfm = dfm - 1#
            tpr = binomialTerm(i, n - i, eprob, 1# - eprob, dfm, 0#)
            pr = pr + tpr
         Wend
         While (pr < cprob)
            i = i + 1#
            tpr = tpr * ((n - i + 1#) * eprob) / (i * (1# - eprob))
            pr = pr + tpr
         Wend
         critbinomial = i
         Exit Function
      End If
   Wend
End Function

Private Function critcompbinomial(ByVal n As Double, ByVal eprob As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(binomial(n,eprob,i)) > cprob and  1-Pr(binomial(n,eprob,i-1)) <= cprob
   If (cprob > 0.5) Then
      critcompbinomial = critbinomial(n, eprob, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   dfm = invcnormal(cprob) * Sqr(n * eprob * (1# - eprob))
   i = n * eprob - dfm
   While (True)
      i = Int(i)
      If (i < 0#) Then
         i = 0
      ElseIf (i > n) Then
         i = n
      End If
      If (i >= max_crit) Then
         critcompbinomial = i
         Exit Function
      End If
      dfm = n * eprob - i
      pr = compbinomial(i, n - i, eprob, 1# - eprob, dfm)
      tpr = 0#
      If (pr > cprob) Then
         i = i + 1#
         dfm = dfm - 1#
         tpr = binomialTerm(i, n - i, eprob, 1# - eprob, dfm, 0#)
         If (pr < (1.00001) * tpr) Then
            While (tpr > cprob)
               i = i + 1#
               tpr = tpr * ((n - i + 1#) * eprob) / (i * (1# - eprob))
            Wend
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcompbinomial = i
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = binomialTerm(i, n - i, eprob, 1# - eprob, n * eprob - i, 0#)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            Else
               i = i + 1#
               tpr = tpr * ((n - i + 1#) * eprob) / (i * (1# - eprob))
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcompbinomial = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * ((n - i + 1#) * eprob) / (i * (1# - eprob))
                     pr = pr - tpr
                  Wend
                  critcompbinomial = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((n - i + 1#) * eprob) / (i * (1# - eprob))) + 0.5)
                  i = i + temp
                  If (i > n) Then
                     i = n
                  End If
                  temp2 = binomialTerm(i, n - i, eprob, 1# - eprob, n * eprob - i, 0#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((n - i + 1#) * eprob) / (i * (1# - eprob)))
                     i = i + temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = binomialTerm(i, n - i, eprob, 1# - eprob, dfm, 0#)
            pr = pr + tpr
            i = i - 1#
            dfm = dfm + 1#
         Wend
         While (pr <= cprob)
            tpr = tpr * ((i + 1#) * (1# - eprob)) / ((n - i) * eprob)
            pr = pr + tpr
            i = i - 1#
         Wend
         critcompbinomial = i + 1#
         Exit Function
      End If
   Wend
End Function

Private Function crithyperg(ByVal j As Double, ByVal k As Double, ByVal m As Double, ByVal cprob As Double) As Double
'//i such that Pr(hypergeometric(i,j,k,m)) >= cprob and  Pr(hypergeometric(i-1,j,k,m)) < cprob
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   If (cprob > 0.5) Then
      crithyperg = critcomphyperg(j, k, m, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double
   i = j * k / m + invcnormal(cprob) * Sqr(j * k * (m - j) * (m - k) / (m * m * (m - 1#)))
   Dim MX As Double, mn  As Double
   MX = Min(j, k)
   mn = max(0, j + k - m)
   While (True)
      If (i < mn) Then
         i = mn
      ElseIf (i > MX) Then
         i = MX
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         crithyperg = i
         Exit Function
      End If
      pr = hypergeometric(i, j - i, k - i, m - k - j + i, False, ha1, hprob, hswap)
      tpr = 0
      If (pr >= cprob) Then
         If (i = mn) Then
            crithyperg = mn
            Exit Function
         End If
         tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
         If (pr < (1 + 0.00001) * tpr) Then
            tpr = tpr * ((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))
            i = i - 1#
            While (tpr > cprob)
               tpr = tpr * ((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))
               i = i - 1#
            Wend
         Else
            pr = pr - tpr
            If (pr < cprob) Then
               crithyperg = i
               Exit Function
            End If
            i = i - 1#
            If (i = mn) Then
               crithyperg = mn
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10) Then
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
               i = i - temp * (tpr - temp2) / (2# * temp2)
            Else
               tpr = tpr * ((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))
               pr = pr - tpr
               If (pr < cprob) Then
                  crithyperg = i
                  Exit Function
               End If
               i = i - 1#
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     tpr = tpr * ((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))
                     pr = pr - tpr
                     i = i - 1#
                  Wend
                  crithyperg = i + 1#
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))) + 0.5)
                  i = i - temp
                  temp2 = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i)))
                     i = i - temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
            pr = pr + tpr
         Wend
         While (pr < cprob)
            i = i + 1#
            tpr = tpr * ((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i))
            pr = pr + tpr
         Wend
         crithyperg = i
         Exit Function
      End If
   Wend
End Function

Private Function critcomphyperg(ByVal j As Double, ByVal k As Double, ByVal m As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(hypergeometric(i,j,k,m)) > cprob and  1-Pr(hypergeometric(i-1,j,k,m)) <= cprob
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   If (cprob > 0.5) Then
      critcomphyperg = crithyperg(j, k, m, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double
   i = j * k / m - invcnormal(cprob) * Sqr(j * k * (m - j) * (m - k) / (m * m * (m - 1#)))
   Dim MX As Double, mn  As Double
   MX = Min(j, k)
   mn = max(0, j + k - m)
   While (True)
      If (i < mn) Then
         i = mn
      ElseIf (i > MX) Then
         i = MX
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         critcomphyperg = i
         Exit Function
      End If
      pr = hypergeometric(i, j - i, k - i, m - k - j + i, True, ha1, hprob, hswap)
      tpr = 0#
      If (pr > cprob) Then
         i = i + 1#
         tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
         If (pr < (1# + 0.00001) * tpr) Then
            While (tpr > cprob)
               i = i + 1
               tpr = tpr * ((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i))
            Wend
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcomphyperg = i
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            Else
               i = i + 1#
               tpr = tpr * ((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i))
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcomphyperg = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * ((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i))
                     pr = pr - tpr
                  Wend
                  critcomphyperg = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i))) + 0.5)
                  i = i + temp
                  temp2 = hypergeometricTerm(i, j - i, k, m - k)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i)))
                     i = i + temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
            pr = pr + tpr
            i = i - 1#
         Wend
         While (pr <= cprob)
            tpr = tpr * ((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))
            pr = pr + tpr
            i = i - 1#
         Wend
         critcomphyperg = i + 1#
         Exit Function
      End If
   Wend
End Function

Private Function critnegbinom(ByVal n As Double, ByVal eprob As Double, ByVal fprob As Double, ByVal cprob As Double) As Double
'//i such that Pr(negbinomial(n,eprob,i)) >= cprob and  Pr(negbinomial(n,eprob,i-1)) < cprob
   If (cprob > 0.5) Then
      critnegbinom = critcompnegbinom(n, eprob, fprob, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   i = invgamma(n * fprob, cprob) / eprob
   While (True)
      If (i < 0#) Then
         i = 0#
      End If
      i = Int(i)
      If (i >= max_crit) Then
         critnegbinom = i
         Exit Function
      End If
      If eprob <= fprob Then
         pr = beta(eprob, n, i + 1#)
      Else
         pr = compbeta(fprob, i + 1#, n)
      End If
      tpr = 0#
      If (pr >= cprob) Then
         If (i = 0#) Then
            critnegbinom = i
            Exit Function
         End If
         If eprob <= fprob Then
            dfm = n - (n + i) * eprob
         Else
            dfm = (n + i) * fprob - i
         End If
         tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
         If (pr < (1 + 0.00001) * tpr) Then
            tpr = tpr * (i + 1#) / ((n + i) * fprob)
            i = i - 1#
            While (tpr > cprob)
               tpr = tpr * (i + 1#) / ((n + i) * fprob)
               i = i - 1#
            Wend
         Else
            pr = pr - tpr
            If (pr < cprob) Then
               critnegbinom = i
               Exit Function
            End If
            i = i - 1#
            If (i = 0#) Then
               critnegbinom = i
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i - temp
               If eprob <= fprob Then
                  dfm = n - (n + i) * eprob
               Else
                  dfm = (n + i) * fprob - i
               End If
               temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
               i = i - temp * (tpr - temp2) / (2# * temp2)
            Else
               tpr = tpr * (i + 1#) / ((n + i) * fprob)
               pr = pr - tpr
               If (pr < cprob) Then
                  critnegbinom = i
                  Exit Function
               End If
               i = i - 1#
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     tpr = tpr * (i + 1#) / ((n + i) * fprob)
                     pr = pr - tpr
                     i = i - 1#
                  Wend
                  critnegbinom = i + 1#
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log((i + 1#) / ((n + i) * fprob)) + 0.5)
                  i = i - temp
                  If eprob <= fprob Then
                     dfm = n - (n + i) * eprob
                  Else
                     dfm = (n + i) * fprob - i
                  End If
                  temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log((i + 1#) / ((n + i) * fprob))
                     i = i - temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            If eprob <= fprob Then
               dfm = n - (n + i) * eprob
            Else
               dfm = (n + i) * fprob - i
            End If
            tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
            pr = pr + tpr
         Wend
         While (pr < cprob)
            i = i + 1#
            tpr = tpr * ((n + i - 1#) * fprob) / i
            pr = pr + tpr
         Wend
         critnegbinom = i
         Exit Function
      End If
   Wend
End Function

Private Function critcompnegbinom(ByVal n As Double, ByVal eprob As Double, ByVal fprob As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(negbinomial(n,eprob,i)) > cprob and  1-Pr(negbinomial(n,eprob,i-1)) <= cprob
   If (cprob > 0.5) Then
      critcompnegbinom = critnegbinom(n, eprob, fprob, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   i = invcompgamma(n * fprob, cprob) / eprob
   While (True)
      If (i < 0#) Then
         i = 0#
      End If
      i = Int(i)
      If (i >= max_crit) Then
         critcompnegbinom = i
         Exit Function
      End If
      If eprob <= fprob Then
         pr = compbeta(eprob, n, i + 1#)
      Else
         pr = beta(fprob, i + 1#, n)
      End If
      If (pr > cprob) Then
         i = i + 1#
         If eprob <= fprob Then
            dfm = n - (n + i) * eprob
         Else
            dfm = (n + i) * fprob - i
         End If
         tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
         If (pr < (1.00001) * tpr) Then
            While (tpr > cprob)
               i = i + 1#
               tpr = tpr * ((n + i - 1#) * fprob) / i
            Wend
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcompnegbinom = i
               Exit Function
            ElseIf (tpr < 0.000000000000001 * pr) Then
               If (tpr < cSmall) Then
                  critcompnegbinom = i
               Else
                  critcompnegbinom = i + Int((pr - cprob) / tpr)
               End If
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i + temp
               If eprob <= fprob Then
                  dfm = n - (n + i) * eprob
               Else
                  dfm = (n + i) * fprob - i
               End If
               temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            Else
               i = i + 1#
               tpr = tpr * ((n + i - 1#) * fprob) / i
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcompnegbinom = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * ((n + i - 1#) * fprob) / i
                     pr = pr - tpr
                  Wend
                  critcompnegbinom = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((n + i - 1#) * fprob) / i) + 0.5)
                  i = i + temp
                  If eprob <= fprob Then
                     dfm = n - (n + i) * eprob
                  Else
                     dfm = (n + i) * fprob - i
                  End If
                  temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((n + i - 1#) * fprob) / i)
                     i = i + temp
                  End If
               End If
            End If
         End If
      Else
         If eprob <= fprob Then
            dfm = n - (n + i) * eprob
         Else
            dfm = (n + i) * fprob - i
         End If
         tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
         If (tpr < 0.000000000000001 * pr) Then
            If (tpr < cSmall) Then
               critcompnegbinom = i
            Else
               critcompnegbinom = i - Int((cprob - pr) / tpr)
            End If
            Exit Function
         End If
         While ((tpr < cSmall) And (pr <= cprob))
            pr = pr + tpr
            i = i - 1#
            If eprob <= fprob Then
               dfm = n - (n + i) * eprob
            Else
               dfm = (n + i) * fprob - i
            End If
            tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
         Wend
         While (pr <= cprob)
            pr = pr + tpr
            i = i - 1#
            If i < 0# Then
               critcompnegbinom = 0#
               Exit Function
            End If
            tpr = tpr * (i + 1#) / ((n + i) * fprob)
         Wend
         critcompnegbinom = i + 1#
         Exit Function
      End If
   Wend
End Function

Private Function critneghyperg(ByVal j As Double, ByVal k As Double, ByVal m As Double, ByVal cprob As Double) As Double
'//i such that Pr(neghypergeometric(i,j,k,m)) >= cprob and  Pr(neghypergeometric(i-1,j,k,m)) < cprob
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   If (cprob > 0.5) Then
      critneghyperg = critcompneghyperg(j, k, m, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   pr = (m - k) / m
   i = invbeta(j * pr, pr * (k - j + 1#), cprob, oneMinusP) * (m - k)
   While (True)
      If (i < 0#) Then
         i = 0#
      ElseIf (i > m - k) Then
         i = m - k
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         critneghyperg = i
         Exit Function
      End If
      pr = hypergeometric(i, j, m - k - i, k - j, False, ha1, hprob, hswap)
      tpr = 0#
      If (pr >= cprob) Then
         If (i = 0#) Then
            critneghyperg = 0#
            Exit Function
         End If
         tpr = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
         If (pr < (1# + 0.00001) * tpr) Then
            tpr = tpr * ((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))
            i = i - 1#
            While (tpr > cprob)
               tpr = tpr * ((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))
               i = i - 1#
            Wend
         Else
            pr = pr - tpr
            If (pr < cprob) Then
               critneghyperg = i
               Exit Function
            End If
            i = i - 1#

            If (i = 0#) Then
               critneghyperg = 0#
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10) Then
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
               i = i - temp * (tpr - temp2) / (2 * temp2)
            Else
               tpr = tpr * ((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))
               pr = pr - tpr
               If (pr < cprob) Then
                  critneghyperg = i
                  Exit Function
               End If
               i = i - 1#
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     tpr = tpr * ((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))
                     pr = pr - tpr
                     i = i - 1#
                  Wend
                  critneghyperg = i + 1#
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))) + 0.5)
                  i = i - temp
                  temp2 = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k)))
                     i = i - temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            tpr = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
            pr = pr + tpr
         Wend
         While (pr < cprob)
            i = i + 1#
            tpr = tpr * ((j + i - 1#) * (m - i - k + 1#)) / (i * (m - j - i + 1#))
            pr = pr + tpr
         Wend
         critneghyperg = i
         Exit Function
      End If
   Wend
End Function

Private Function critcompneghyperg(ByVal j As Double, ByVal k As Double, ByVal m As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(neghypergeometric(i,j,k,m)) > cprob and  1-Pr(neghypergeometric(i-1,j,k,m)) <= cprob
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   If (cprob > 0.5) Then
      critcompneghyperg = critneghyperg(j, k, m, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   pr = (m - k) / m
   i = invcompbeta(j * pr, pr * (k - j + 1#), cprob, oneMinusP) * (m - k)
   While (True)
      If (i < 0#) Then
         i = 0#
      ElseIf (i > m - k) Then
         i = m - k
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         critcompneghyperg = i
         Exit Function
      End If
      pr = hypergeometric(i, j, m - k - i, k - j, True, ha1, hprob, hswap)
      tpr = 0#
      If (pr > cprob) Then
         i = i + 1#
         tpr = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
         If (pr < (1 + 0.00001) * tpr) Then
            Do While (tpr > cprob)
               i = i + 1#
               temp = m - j - i + 1#
               If temp = 0# Then Exit Do
               tpr = tpr * ((j + i - 1#) * (m - i - k + 1#)) / (i * temp)
            Loop
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcompneghyperg = i
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
               i = i + temp * (tpr - temp2) / (2 * temp2)
            Else
               i = i + 1#
               tpr = tpr * ((j + i - 1#) * (m - i - k + 1#)) / (i * (m - j - i + 1#))
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcompneghyperg = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * ((j + i - 1#) * (m - i - k + 1#)) / (i * (m - j - i + 1#))
                     pr = pr - tpr
                  Wend
                  critcompneghyperg = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((j + i - 1#) * (m - i - k + 1#)) / (i * (m - j - i + 1#))) + 0.5)
                  i = i + temp
                  temp2 = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((j + i - 1#) * (m - i - k + 1#)) / (i * (m - j - i + 1#)))
                     i = i + temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
            pr = pr + tpr
            i = i - 1#
         Wend
         While (pr <= cprob)
            tpr = tpr * ((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))
            pr = pr + tpr
            i = i - 1#
         Wend
         critcompneghyperg = i + 1#
         Exit Function
      End If
   Wend
End Function

Private Function AlterForIntegralChecks_Others(ByVal value As Double) As Double
   If NonIntegralValuesAllowed_Others Then
      AlterForIntegralChecks_Others = Int(value)
   ElseIf value <> Int(value) Then
      AlterForIntegralChecks_Others = [#VALUE!]
   Else
      AlterForIntegralChecks_Others = value
   End If
End Function

Private Function AlterForIntegralChecks_df(ByVal value As Double) As Double
   If NonIntegralValuesAllowed_df Then
      AlterForIntegralChecks_df = value
   Else
      AlterForIntegralChecks_df = AlterForIntegralChecks_Others(value)
   End If
End Function

Private Function AlterForIntegralChecks_NB(ByVal value As Double) As Double
   If NonIntegralValuesAllowed_NB Then
      AlterForIntegralChecks_NB = value
   Else
      AlterForIntegralChecks_NB = AlterForIntegralChecks_Others(value)
   End If
End Function

Private Function GetRidOfMinusZeroes(ByVal x As Double) As Double
   If x = 0# Then
      GetRidOfMinusZeroes = 0#
   Else
      GetRidOfMinusZeroes = x
   End If
End Function

Public Function pmf_geometric(ByVal failures As Double, ByVal success_prob As Double) As Double
   failures = AlterForIntegralChecks_Others(failures)
   If (success_prob < 0# Or success_prob > 1#) Then
      pmf_geometric = [#VALUE!]
   ElseIf failures < 0# Then
      pmf_geometric = 0#
   ElseIf success_prob = 1# Then
      If failures = 0# Then
         pmf_geometric = 1#
      Else
         pmf_geometric = 0#
      End If
   Else
      pmf_geometric = success_prob * Exp(failures * log0(-success_prob))
   End If
   pmf_geometric = GetRidOfMinusZeroes(pmf_geometric)
End Function

Public Function cdf_geometric(ByVal failures As Double, ByVal success_prob As Double) As Double
   failures = Int(failures)
   If (success_prob < 0# Or success_prob > 1#) Then
      cdf_geometric = [#VALUE!]
   ElseIf failures < 0# Then
      cdf_geometric = 0#
   ElseIf success_prob = 1# Then
      If failures >= 0# Then
         cdf_geometric = 1#
      Else
         cdf_geometric = 0#
      End If
   Else
      cdf_geometric = -expm1((failures + 1#) * log0(-success_prob))
   End If
   cdf_geometric = GetRidOfMinusZeroes(cdf_geometric)
End Function

Public Function comp_cdf_geometric(ByVal failures As Double, ByVal success_prob As Double) As Double
   failures = Int(failures)
   If (success_prob < 0# Or success_prob > 1#) Then
      comp_cdf_geometric = [#VALUE!]
   ElseIf failures < 0# Then
      comp_cdf_geometric = 1#
   ElseIf success_prob = 1# Then
      If failures >= 0# Then
         comp_cdf_geometric = 0#
      Else
         comp_cdf_geometric = 1#
      End If
   Else
      comp_cdf_geometric = Exp((failures + 1#) * log0(-success_prob))
   End If
   comp_cdf_geometric = GetRidOfMinusZeroes(comp_cdf_geometric)
End Function

Public Function crit_geometric(ByVal success_prob As Double, ByVal crit_prob As Double) As Double
   If (success_prob <= 0# Or success_prob > 1# Or crit_prob < 0# Or crit_prob > 1#) Then
      crit_geometric = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_geometric = [#VALUE!]
   ElseIf (success_prob = 1#) Then
      crit_geometric = 0#
   ElseIf (crit_prob = 1#) Then
      crit_geometric = [#VALUE!]
   Else
      crit_geometric = Int(log0(-crit_prob) / log0(-success_prob) - 1#)
      If -expm1((crit_geometric + 1#) * log0(-success_prob)) < crit_prob Then
         crit_geometric = crit_geometric + 1#
      End If
   End If
   crit_geometric = GetRidOfMinusZeroes(crit_geometric)
End Function

Public Function comp_crit_geometric(ByVal success_prob As Double, ByVal crit_prob As Double) As Double
   If (success_prob <= 0# Or success_prob > 1# Or crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_geometric = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_geometric = [#VALUE!]
   ElseIf (success_prob = 1#) Then
      comp_crit_geometric = 0#
   ElseIf (crit_prob = 0#) Then
      comp_crit_geometric = [#VALUE!]
   Else
      comp_crit_geometric = Int(Log(crit_prob) / log0(-success_prob) - 1#)
      If Exp((comp_crit_geometric + 1#) * log0(-success_prob)) > crit_prob Then
         comp_crit_geometric = comp_crit_geometric + 1#
      End If
   End If
   comp_crit_geometric = GetRidOfMinusZeroes(comp_crit_geometric)
End Function

Public Function lcb_geometric(ByVal failures As Double, ByVal prob As Double) As Double
   failures = AlterForIntegralChecks_Others(failures)
   If (prob < 0# Or prob > 1# Or failures < 0#) Then
      lcb_geometric = [#VALUE!]
   ElseIf (prob = 1#) Then
      lcb_geometric = 1#
   Else
      lcb_geometric = -expm1(log0(-prob) / (failures + 1#))
   End If
   lcb_geometric = GetRidOfMinusZeroes(lcb_geometric)
End Function

Public Function ucb_geometric(ByVal failures As Double, ByVal prob As Double) As Double
   failures = AlterForIntegralChecks_Others(failures)
   If (prob < 0# Or prob > 1# Or failures < 0#) Then
      ucb_geometric = [#VALUE!]
   ElseIf (prob = 0# Or failures = 0#) Then
      ucb_geometric = 1#
   ElseIf (prob = 1#) Then
      ucb_geometric = 0#
   Else
      ucb_geometric = -expm1(Log(prob) / failures)
   End If
   ucb_geometric = GetRidOfMinusZeroes(ucb_geometric)
End Function

Public Function pmf_negbinomial(ByVal failures As Double, ByVal success_prob As Double, ByVal successes_reqd As Double) As Double
   Dim q As Double, dfm As Double
   failures = AlterForIntegralChecks_Others(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (success_prob < 0# Or success_prob > 1# Or successes_reqd <= 0#) Then
      pmf_negbinomial = [#VALUE!]
   ElseIf (successes_reqd + failures > 0#) Then
      q = 1# - success_prob
      If success_prob <= q Then
         dfm = successes_reqd - (successes_reqd + failures) * success_prob
      Else
         dfm = (successes_reqd + failures) * q - failures
      End If
      pmf_negbinomial = successes_reqd / (successes_reqd + failures) * binomialTerm(failures, successes_reqd, q, success_prob, dfm, 0#)
   ElseIf (failures <> 0#) Then
      pmf_negbinomial = 0#
   Else
      pmf_negbinomial = 1#
   End If
   pmf_negbinomial = GetRidOfMinusZeroes(pmf_negbinomial)
End Function

Public Function cdf_negbinomial(ByVal failures As Double, ByVal success_prob As Double, ByVal successes_reqd As Double) As Double
   Dim q As Double
   failures = Int(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (success_prob < 0# Or success_prob > 1# Or successes_reqd <= 0#) Then
      cdf_negbinomial = [#VALUE!]
   Else
      q = 1# - success_prob
      If q < success_prob Then
         cdf_negbinomial = compbeta(q, failures + 1, successes_reqd)
      Else
         cdf_negbinomial = beta(success_prob, successes_reqd, failures + 1)
      End If
   End If
   cdf_negbinomial = GetRidOfMinusZeroes(cdf_negbinomial)
End Function

Public Function comp_cdf_negbinomial(ByVal failures As Double, ByVal success_prob As Double, ByVal successes_reqd As Double) As Double
   Dim q As Double
   failures = Int(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (success_prob < 0# Or success_prob > 1# Or successes_reqd <= 0#) Then
      comp_cdf_negbinomial = [#VALUE!]
   Else
      q = 1# - success_prob
      If q < success_prob Then
         comp_cdf_negbinomial = beta(q, failures + 1, successes_reqd)
      Else
         comp_cdf_negbinomial = compbeta(success_prob, successes_reqd, failures + 1)
      End If
   End If
   comp_cdf_negbinomial = GetRidOfMinusZeroes(comp_cdf_negbinomial)
End Function

Public Function crit_negbinomial(ByVal success_prob As Double, ByVal successes_reqd As Double, ByVal crit_prob As Double) As Double
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (success_prob <= 0# Or success_prob > 1# Or successes_reqd <= 0# Or crit_prob < 0# Or crit_prob > 1#) Then
      crit_negbinomial = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_negbinomial = [#VALUE!]
   ElseIf (success_prob = 1#) Then
      crit_negbinomial = 0#
   ElseIf (crit_prob = 1#) Then
      crit_negbinomial = [#VALUE!]
   Else
      Dim i As Double, pr As Double
      crit_negbinomial = critnegbinom(successes_reqd, success_prob, 1# - success_prob, crit_prob)
      i = crit_negbinomial
      pr = cdf_negbinomial(i, success_prob, successes_reqd)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         pr = cdf_negbinomial(i, success_prob, successes_reqd)
         If (pr >= crit_prob) Then
            crit_negbinomial = i
         End If
      Else
         crit_negbinomial = i + 1#
      End If
   End If
   crit_negbinomial = GetRidOfMinusZeroes(crit_negbinomial)
End Function

Public Function comp_crit_negbinomial(ByVal success_prob As Double, ByVal successes_reqd As Double, ByVal crit_prob As Double) As Double
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (success_prob <= 0# Or success_prob > 1# Or successes_reqd <= 0# Or crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_negbinomial = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_negbinomial = [#VALUE!]
   ElseIf (success_prob = 1#) Then
      comp_crit_negbinomial = 0#
   ElseIf (crit_prob = 0#) Then
      comp_crit_negbinomial = [#VALUE!]
   Else
      Dim i As Double, pr As Double
      comp_crit_negbinomial = critcompnegbinom(successes_reqd, success_prob, 1# - success_prob, crit_prob)
      i = comp_crit_negbinomial
      pr = comp_cdf_negbinomial(i, success_prob, successes_reqd)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         pr = comp_cdf_negbinomial(i, success_prob, successes_reqd)
         If (pr <= crit_prob) Then
            comp_crit_negbinomial = i
         End If
      Else
         comp_crit_negbinomial = i + 1#
      End If
   End If
   comp_crit_negbinomial = GetRidOfMinusZeroes(comp_crit_negbinomial)
End Function

Public Function lcb_negbinomial(ByVal failures As Double, ByVal successes_reqd As Double, ByVal prob As Double) As Double
   failures = AlterForIntegralChecks_Others(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (prob < 0# Or prob > 1# Or failures < 0# Or successes_reqd <= 0#) Then
      lcb_negbinomial = [#VALUE!]
   ElseIf (prob = 0#) Then
      lcb_negbinomial = 0#
   ElseIf (prob = 1#) Then
      lcb_negbinomial = 1#
   Else
      Dim oneMinusP As Double
      lcb_negbinomial = invbeta(successes_reqd, failures + 1, prob, oneMinusP)
   End If
   lcb_negbinomial = GetRidOfMinusZeroes(lcb_negbinomial)
End Function

Public Function ucb_negbinomial(ByVal failures As Double, ByVal successes_reqd As Double, ByVal prob As Double) As Double
   failures = AlterForIntegralChecks_Others(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (prob < 0# Or prob > 1# Or failures < 0# Or successes_reqd <= 0#) Then
      ucb_negbinomial = [#VALUE!]
   ElseIf (prob = 0# Or failures = 0#) Then
      ucb_negbinomial = 1#
   ElseIf (prob = 1#) Then
      ucb_negbinomial = 0#
   Else
      Dim oneMinusP As Double
      ucb_negbinomial = invcompbeta(successes_reqd, failures, prob, oneMinusP)
   End If
   ucb_negbinomial = GetRidOfMinusZeroes(ucb_negbinomial)
End Function

Public Function pmf_binomial(ByVal SAMPLE_SIZE As Double, ByVal successes As Double, ByVal success_prob As Double) As Double
   Dim q As Double, dfm As Double
   successes = AlterForIntegralChecks_Others(successes)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (success_prob < 0# Or success_prob > 1# Or SAMPLE_SIZE < 0#) Then
      pmf_binomial = [#VALUE!]
   Else
      q = 1# - success_prob
      If success_prob <= q Then
         dfm = SAMPLE_SIZE * success_prob - successes
      Else
         dfm = (SAMPLE_SIZE - successes) - SAMPLE_SIZE * q
      End If
      pmf_binomial = binomialTerm(successes, SAMPLE_SIZE - successes, success_prob, q, dfm, 0#)
   End If
   pmf_binomial = GetRidOfMinusZeroes(pmf_binomial)
End Function

Public Function cdf_binomial(ByVal SAMPLE_SIZE As Double, ByVal successes As Double, ByVal success_prob As Double) As Double
   Dim q As Double, dfm As Double
   successes = Int(successes)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (success_prob < 0# Or success_prob > 1# Or SAMPLE_SIZE < 0#) Then
      cdf_binomial = [#VALUE!]
   Else
      q = 1# - success_prob
      If success_prob <= q Then
         dfm = SAMPLE_SIZE * success_prob - successes
      Else
         dfm = (SAMPLE_SIZE - successes) - SAMPLE_SIZE * q
      End If
      cdf_binomial = binomial(successes, SAMPLE_SIZE - successes, success_prob, q, dfm)
   End If
   cdf_binomial = GetRidOfMinusZeroes(cdf_binomial)
End Function

Public Function comp_cdf_binomial(ByVal SAMPLE_SIZE As Double, ByVal successes As Double, ByVal success_prob As Double) As Double
   Dim q As Double, dfm As Double
   successes = Int(successes)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (success_prob < 0# Or success_prob > 1# Or SAMPLE_SIZE < 0#) Then
      comp_cdf_binomial = [#VALUE!]
   Else
      q = 1# - success_prob
      If success_prob <= q Then
         dfm = SAMPLE_SIZE * success_prob - successes
      Else
         dfm = (SAMPLE_SIZE - successes) - SAMPLE_SIZE * q
      End If
      comp_cdf_binomial = compbinomial(successes, SAMPLE_SIZE - successes, success_prob, q, dfm)
   End If
   comp_cdf_binomial = GetRidOfMinusZeroes(comp_cdf_binomial)
End Function

Public Function crit_binomial(ByVal SAMPLE_SIZE As Double, ByVal success_prob As Double, ByVal crit_prob As Double) As Double
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (success_prob < 0# Or success_prob > 1# Or SAMPLE_SIZE < 0# Or crit_prob < 0# Or crit_prob > 1#) Then
      crit_binomial = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_binomial = [#VALUE!]
   ElseIf (success_prob = 0#) Then
      crit_binomial = 0#
   ElseIf (crit_prob = 1# Or success_prob = 1#) Then
      crit_binomial = SAMPLE_SIZE
   Else
      Dim pr As Double, i As Double
      crit_binomial = critbinomial(SAMPLE_SIZE, success_prob, crit_prob)
      i = crit_binomial
      pr = cdf_binomial(SAMPLE_SIZE, i, success_prob)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         pr = cdf_binomial(SAMPLE_SIZE, i, success_prob)
         If (pr >= crit_prob) Then
            crit_binomial = i
         End If
      Else
         crit_binomial = i + 1#
      End If
   End If
   crit_binomial = GetRidOfMinusZeroes(crit_binomial)
End Function

Public Function comp_crit_binomial(ByVal SAMPLE_SIZE As Double, ByVal success_prob As Double, ByVal crit_prob As Double) As Double
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (success_prob < 0# Or success_prob > 1# Or SAMPLE_SIZE < 0# Or crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_binomial = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_binomial = [#VALUE!]
   ElseIf (crit_prob = 0# Or success_prob = 1#) Then
      comp_crit_binomial = SAMPLE_SIZE
   ElseIf (success_prob = 0#) Then
      comp_crit_binomial = 0#
   Else
      Dim pr As Double, i As Double
      comp_crit_binomial = critcompbinomial(SAMPLE_SIZE, success_prob, crit_prob)
      i = comp_crit_binomial
      pr = comp_cdf_binomial(SAMPLE_SIZE, i, success_prob)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         pr = comp_cdf_binomial(SAMPLE_SIZE, i, success_prob)
         If (pr <= crit_prob) Then
            comp_crit_binomial = i
         End If
      Else
         comp_crit_binomial = i + 1#
      End If
   End If
   comp_crit_binomial = GetRidOfMinusZeroes(comp_crit_binomial)
End Function

Public Function lcb_binomial(ByVal SAMPLE_SIZE As Double, ByVal successes As Double, ByVal prob As Double) As Double
   successes = AlterForIntegralChecks_Others(successes)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (prob < 0# Or prob > 1#) Then
      lcb_binomial = [#VALUE!]
   ElseIf (SAMPLE_SIZE < successes Or successes < 0#) Then
      lcb_binomial = [#VALUE!]
   ElseIf (prob = 0# Or successes = 0#) Then
      lcb_binomial = 0#
   ElseIf (prob = 1#) Then
      lcb_binomial = 1#
   Else
      Dim oneMinusP As Double
      lcb_binomial = invcompbinom(successes - 1#, SAMPLE_SIZE - successes + 1#, prob, oneMinusP)
   End If
   lcb_binomial = GetRidOfMinusZeroes(lcb_binomial)
End Function

Public Function ucb_binomial(ByVal SAMPLE_SIZE As Double, ByVal successes As Double, ByVal prob As Double) As Double
   successes = AlterForIntegralChecks_Others(successes)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (prob < 0# Or prob > 1#) Then
      ucb_binomial = [#VALUE!]
   ElseIf (SAMPLE_SIZE < successes Or successes < 0#) Then
      ucb_binomial = [#VALUE!]
   ElseIf (prob = 0# Or successes = SAMPLE_SIZE#) Then
      ucb_binomial = 1#
   ElseIf (prob = 1#) Then
      ucb_binomial = 0#
   Else
      Dim oneMinusP As Double
      ucb_binomial = invbinom(successes, SAMPLE_SIZE - successes, prob, oneMinusP)
   End If
   ucb_binomial = GetRidOfMinusZeroes(ucb_binomial)
End Function

Public Function pmf_poisson(ByVal Mean As Double, ByVal i As Double) As Double
   i = AlterForIntegralChecks_Others(i)
   If (Mean < 0#) Then
      pmf_poisson = [#VALUE!]
   ElseIf (i < 0#) Then
      pmf_poisson = 0#
   Else
      pmf_poisson = poissonTerm(i, Mean, Mean - i, 0#)
   End If
   pmf_poisson = GetRidOfMinusZeroes(pmf_poisson)
End Function

Public Function cdf_poisson(ByVal Mean As Double, ByVal i As Double) As Double
   i = Int(i)
   If (Mean < 0#) Then
      cdf_poisson = [#VALUE!]
   ElseIf (i < 0#) Then
      cdf_poisson = 0#
   Else
      cdf_poisson = CPoisson(i, Mean, Mean - i)
   End If
   cdf_poisson = GetRidOfMinusZeroes(cdf_poisson)
End Function

Public Function comp_cdf_poisson(ByVal Mean As Double, ByVal i As Double) As Double
   i = Int(i)
   If (Mean < 0#) Then
      comp_cdf_poisson = [#VALUE!]
   ElseIf (i < 0#) Then
      comp_cdf_poisson = 1#
   Else
      comp_cdf_poisson = comppoisson(i, Mean, Mean - i)
   End If
   comp_cdf_poisson = GetRidOfMinusZeroes(comp_cdf_poisson)
End Function

Public Function crit_poisson(ByVal Mean As Double, ByVal crit_prob As Double) As Double
   If (crit_prob < 0# Or crit_prob > 1# Or Mean < 0#) Then
      crit_poisson = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_poisson = [#VALUE!]
   ElseIf (Mean = 0#) Then
      crit_poisson = 0#
   ElseIf (crit_prob = 1#) Then
      crit_poisson = [#VALUE!]
   Else
      Dim pr As Double
      crit_poisson = critpoiss(Mean, crit_prob)
      pr = CPoisson(crit_poisson, Mean, Mean - crit_poisson)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         crit_poisson = crit_poisson - 1#
         pr = CPoisson(crit_poisson, Mean, Mean - crit_poisson)
         If (pr < crit_prob) Then
            crit_poisson = crit_poisson + 1#
         End If
      Else
         crit_poisson = crit_poisson + 1#
      End If
   End If
   crit_poisson = GetRidOfMinusZeroes(crit_poisson)
End Function

Public Function comp_crit_poisson(ByVal Mean As Double, ByVal crit_prob As Double) As Double
   If (crit_prob < 0# Or crit_prob > 1# Or Mean < 0#) Then
      comp_crit_poisson = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_poisson = [#VALUE!]
   ElseIf (Mean = 0#) Then
      comp_crit_poisson = 0#
   ElseIf (crit_prob = 0#) Then
      comp_crit_poisson = [#VALUE!]
   Else
      Dim pr As Double
      comp_crit_poisson = critcomppoiss(Mean, crit_prob)
      pr = comppoisson(comp_crit_poisson, Mean, Mean - comp_crit_poisson)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         comp_crit_poisson = comp_crit_poisson - 1#
         pr = comppoisson(comp_crit_poisson, Mean, Mean - comp_crit_poisson)
         If (pr > crit_prob) Then
            comp_crit_poisson = comp_crit_poisson + 1#
         End If
      Else
         comp_crit_poisson = comp_crit_poisson + 1#
      End If
   End If
   comp_crit_poisson = GetRidOfMinusZeroes(comp_crit_poisson)
End Function

Public Function lcb_poisson(ByVal i As Double, ByVal prob As Double) As Double
   i = AlterForIntegralChecks_Others(i)
   If (prob < 0# Or prob > 1# Or i < 0#) Then
      lcb_poisson = [#VALUE!]
   ElseIf (prob = 0# Or i = 0#) Then
      lcb_poisson = 0#
   ElseIf (prob = 1#) Then
      lcb_poisson = [#VALUE!]
   Else
      lcb_poisson = invcomppoisson(i - 1#, prob)
   End If
   lcb_poisson = GetRidOfMinusZeroes(lcb_poisson)
End Function

Public Function ucb_poisson(ByVal i As Double, ByVal prob As Double) As Double
   i = AlterForIntegralChecks_Others(i)
   If (prob <= 0# Or prob > 1#) Then
      ucb_poisson = [#VALUE!]
   ElseIf (i < 0#) Then
      ucb_poisson = [#VALUE!]
   ElseIf (prob = 1#) Then
      ucb_poisson = 0#
   Else
      ucb_poisson = invpoisson(i, prob)
   End If
   ucb_poisson = GetRidOfMinusZeroes(ucb_poisson)
End Function

Public Function pmf_hypergeometric(ByVal type1s As Double, ByVal SAMPLE_SIZE As Double, ByVal tot_type1 As Double, ByVal POP_SIZE As Double) As Double
   type1s = AlterForIntegralChecks_Others(type1s)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (SAMPLE_SIZE < 0# Or tot_type1 < 0# Or SAMPLE_SIZE > POP_SIZE Or tot_type1 > POP_SIZE) Then
      pmf_hypergeometric = [#VALUE!]
   Else
      pmf_hypergeometric = hypergeometricTerm(type1s, SAMPLE_SIZE - type1s, tot_type1 - type1s, POP_SIZE - tot_type1 - SAMPLE_SIZE + type1s)
   End If
   pmf_hypergeometric = GetRidOfMinusZeroes(pmf_hypergeometric)
End Function

Public Function cdf_hypergeometric(ByVal type1s As Double, ByVal SAMPLE_SIZE As Double, ByVal tot_type1 As Double, ByVal POP_SIZE As Double) As Double
   type1s = Int(type1s)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (SAMPLE_SIZE < 0# Or tot_type1 < 0# Or SAMPLE_SIZE > POP_SIZE Or tot_type1 > POP_SIZE) Then
      cdf_hypergeometric = [#VALUE!]
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      cdf_hypergeometric = hypergeometric(type1s, SAMPLE_SIZE - type1s, tot_type1 - type1s, POP_SIZE - tot_type1 - SAMPLE_SIZE + type1s, False, ha1, hprob, hswap)
   End If
   cdf_hypergeometric = GetRidOfMinusZeroes(cdf_hypergeometric)
End Function

Public Function comp_cdf_hypergeometric(ByVal type1s As Double, ByVal SAMPLE_SIZE As Double, ByVal tot_type1 As Double, ByVal POP_SIZE As Double) As Double
   type1s = Int(type1s)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (SAMPLE_SIZE < 0# Or tot_type1 < 0# Or SAMPLE_SIZE > POP_SIZE Or tot_type1 > POP_SIZE) Then
      comp_cdf_hypergeometric = [#VALUE!]
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      comp_cdf_hypergeometric = hypergeometric(type1s, SAMPLE_SIZE - type1s, tot_type1 - type1s, POP_SIZE - tot_type1 - SAMPLE_SIZE + type1s, True, ha1, hprob, hswap)
   End If
   comp_cdf_hypergeometric = GetRidOfMinusZeroes(comp_cdf_hypergeometric)
End Function

Public Function crit_hypergeometric(ByVal SAMPLE_SIZE As Double, ByVal tot_type1 As Double, ByVal POP_SIZE As Double, ByVal crit_prob As Double) As Double
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (crit_prob < 0# Or crit_prob > 1#) Then
      crit_hypergeometric = [#VALUE!]
   ElseIf (SAMPLE_SIZE < 0# Or tot_type1 < 0# Or SAMPLE_SIZE > POP_SIZE Or tot_type1 > POP_SIZE) Then
      crit_hypergeometric = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_hypergeometric = [#VALUE!]
   ElseIf (SAMPLE_SIZE = 0# Or tot_type1 = 0#) Then
      crit_hypergeometric = 0#
   ElseIf (SAMPLE_SIZE = POP_SIZE Or tot_type1 = POP_SIZE) Then
      crit_hypergeometric = Min(SAMPLE_SIZE, tot_type1)
   ElseIf (crit_prob = 1#) Then
      crit_hypergeometric = Min(SAMPLE_SIZE, tot_type1)
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      Dim i As Double, pr As Double
      crit_hypergeometric = crithyperg(SAMPLE_SIZE, tot_type1, POP_SIZE, crit_prob)
      i = crit_hypergeometric
      pr = hypergeometric(i, SAMPLE_SIZE - i, tot_type1 - i, POP_SIZE - tot_type1 - SAMPLE_SIZE + i, False, ha1, hprob, hswap)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         pr = hypergeometric(i, SAMPLE_SIZE - i, tot_type1 - i, POP_SIZE - tot_type1 - SAMPLE_SIZE + i, False, ha1, hprob, hswap)
         If (pr >= crit_prob) Then
            crit_hypergeometric = i
         End If
      Else
         crit_hypergeometric = i + 1#
      End If
   End If
   crit_hypergeometric = GetRidOfMinusZeroes(crit_hypergeometric)
End Function

Public Function comp_crit_hypergeometric(ByVal SAMPLE_SIZE As Double, ByVal tot_type1 As Double, ByVal POP_SIZE As Double, ByVal crit_prob As Double) As Double
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_hypergeometric = [#VALUE!]
   ElseIf (SAMPLE_SIZE < 0# Or tot_type1 < 0# Or SAMPLE_SIZE > POP_SIZE Or tot_type1 > POP_SIZE) Then
      comp_crit_hypergeometric = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_hypergeometric = [#VALUE!]
   ElseIf (SAMPLE_SIZE = 0# Or tot_type1 = 0#) Then
      comp_crit_hypergeometric = 0#
   ElseIf (SAMPLE_SIZE = POP_SIZE Or tot_type1 = POP_SIZE) Then
      comp_crit_hypergeometric = Min(SAMPLE_SIZE, tot_type1)
   ElseIf (crit_prob = 0#) Then
      comp_crit_hypergeometric = Min(SAMPLE_SIZE, tot_type1)
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      Dim i As Double, pr As Double
      comp_crit_hypergeometric = critcomphyperg(SAMPLE_SIZE, tot_type1, POP_SIZE, crit_prob)
      i = comp_crit_hypergeometric
      pr = hypergeometric(i, SAMPLE_SIZE - i, tot_type1 - i, POP_SIZE - tot_type1 - SAMPLE_SIZE + i, True, ha1, hprob, hswap)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         pr = hypergeometric(i, SAMPLE_SIZE - i, tot_type1 - i, POP_SIZE - tot_type1 - SAMPLE_SIZE + i, True, ha1, hprob, hswap)
         If (pr <= crit_prob) Then
            comp_crit_hypergeometric = i
         End If
      Else
         comp_crit_hypergeometric = i + 1#
      End If
   End If
   comp_crit_hypergeometric = GetRidOfMinusZeroes(comp_crit_hypergeometric)
End Function

Public Function lcb_hypergeometric(ByVal type1s As Double, ByVal SAMPLE_SIZE As Double, ByVal POP_SIZE As Double, ByVal prob As Double) As Double
   type1s = AlterForIntegralChecks_Others(type1s)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (prob < 0# Or prob > 1#) Then
      lcb_hypergeometric = [#VALUE!]
   ElseIf (type1s < 0# Or type1s > SAMPLE_SIZE Or SAMPLE_SIZE > POP_SIZE) Then
      lcb_hypergeometric = [#VALUE!]
   ElseIf (prob = 0# Or type1s = 0# Or POP_SIZE = SAMPLE_SIZE) Then
      lcb_hypergeometric = type1s
   ElseIf (prob = 1#) Then
      lcb_hypergeometric = POP_SIZE - (SAMPLE_SIZE - type1s)
   ElseIf (prob < 0.5) Then
      lcb_hypergeometric = critneghyperg(type1s, SAMPLE_SIZE, POP_SIZE, prob * (1.000000000001)) + type1s
   Else
      lcb_hypergeometric = critcompneghyperg(type1s, SAMPLE_SIZE, POP_SIZE, (1# - prob) * (1# - 0.000000000001)) + type1s
   End If
   lcb_hypergeometric = GetRidOfMinusZeroes(lcb_hypergeometric)
End Function

Public Function ucb_hypergeometric(ByVal type1s As Double, ByVal SAMPLE_SIZE As Double, ByVal POP_SIZE As Double, ByVal prob As Double) As Double
   type1s = AlterForIntegralChecks_Others(type1s)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (prob < 0# Or prob > 1#) Then
      ucb_hypergeometric = [#VALUE!]
   ElseIf (type1s < 0# Or type1s > SAMPLE_SIZE Or SAMPLE_SIZE > POP_SIZE) Then
      ucb_hypergeometric = [#VALUE!]
   ElseIf (prob = 0# Or type1s = SAMPLE_SIZE Or POP_SIZE = SAMPLE_SIZE) Then
      ucb_hypergeometric = POP_SIZE - (SAMPLE_SIZE - type1s)
   ElseIf (prob = 1#) Then
      ucb_hypergeometric = type1s
   ElseIf (prob < 0.5) Then
      ucb_hypergeometric = critcompneghyperg(type1s + 1#, SAMPLE_SIZE, POP_SIZE, prob * (1# - 0.000000000001)) + type1s
   Else
      ucb_hypergeometric = critneghyperg(type1s + 1#, SAMPLE_SIZE, POP_SIZE, (1# - prob) * (1.000000000001)) + type1s
   End If
   ucb_hypergeometric = GetRidOfMinusZeroes(ucb_hypergeometric)
End Function

Public Function pmf_neghypergeometric(ByVal type2s As Double, ByVal type1s_reqd As Double, ByVal tot_type1 As Double, ByVal POP_SIZE As Double) As Double
   type2s = AlterForIntegralChecks_Others(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (type1s_reqd <= 0# Or tot_type1 < type1s_reqd Or tot_type1 > POP_SIZE) Then
      pmf_neghypergeometric = [#VALUE!]
   ElseIf (type2s < 0# Or tot_type1 + type2s > POP_SIZE) Then
      If type2s = 0# Then
         pmf_neghypergeometric = 1#
      Else
         pmf_neghypergeometric = 0#
      End If
   Else
      pmf_neghypergeometric = hypergeometricTerm(type1s_reqd - 1#, type2s, tot_type1 - type1s_reqd + 1#, POP_SIZE - tot_type1 - type2s) * (tot_type1 - type1s_reqd + 1#) / (POP_SIZE - type1s_reqd - type2s + 1#)
   End If
   pmf_neghypergeometric = GetRidOfMinusZeroes(pmf_neghypergeometric)
End Function

Public Function cdf_neghypergeometric(ByVal type2s As Double, ByVal type1s_reqd As Double, ByVal tot_type1 As Double, ByVal POP_SIZE As Double) As Double
   type2s = Int(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (type1s_reqd <= 0# Or tot_type1 < type1s_reqd Or tot_type1 > POP_SIZE) Then
      cdf_neghypergeometric = [#VALUE!]
   ElseIf (tot_type1 + type2s > POP_SIZE) Then
      cdf_neghypergeometric = 1#
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      cdf_neghypergeometric = hypergeometric(type2s, type1s_reqd, POP_SIZE - tot_type1 - type2s, tot_type1 - type1s_reqd, False, ha1, hprob, hswap)
   End If
   cdf_neghypergeometric = GetRidOfMinusZeroes(cdf_neghypergeometric)
End Function

Public Function comp_cdf_neghypergeometric(ByVal type2s As Double, ByVal type1s_reqd As Double, ByVal tot_type1 As Double, ByVal POP_SIZE As Double) As Double
   type2s = Int(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (type1s_reqd <= 0# Or tot_type1 < type1s_reqd Or tot_type1 > POP_SIZE) Then
      comp_cdf_neghypergeometric = [#VALUE!]
   ElseIf (tot_type1 + type2s > POP_SIZE) Then
      comp_cdf_neghypergeometric = 0#
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      comp_cdf_neghypergeometric = hypergeometric(type2s, type1s_reqd, POP_SIZE - tot_type1 - type2s, tot_type1 - type1s_reqd, True, ha1, hprob, hswap)
   End If
   comp_cdf_neghypergeometric = GetRidOfMinusZeroes(comp_cdf_neghypergeometric)
End Function

Public Function crit_neghypergeometric(ByVal type1s_reqd As Double, ByVal tot_type1 As Double, ByVal POP_SIZE As Double, ByVal crit_prob As Double) As Double
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (crit_prob < 0# Or crit_prob > 1#) Then
      crit_neghypergeometric = [#VALUE!]
   ElseIf (type1s_reqd < 0# Or tot_type1 < type1s_reqd Or tot_type1 > POP_SIZE) Then
      crit_neghypergeometric = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_neghypergeometric = [#VALUE!]
   ElseIf (POP_SIZE = tot_type1) Then
      crit_neghypergeometric = 0#
   ElseIf (crit_prob = 1#) Then
      crit_neghypergeometric = POP_SIZE - tot_type1
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      Dim i As Double, pr As Double
      crit_neghypergeometric = critneghyperg(type1s_reqd, tot_type1, POP_SIZE, crit_prob)
      i = crit_neghypergeometric
      pr = hypergeometric(i, type1s_reqd, POP_SIZE - tot_type1 - i, tot_type1 - type1s_reqd, False, ha1, hprob, hswap)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         pr = hypergeometric(i, type1s_reqd, POP_SIZE - tot_type1 - i, tot_type1 - type1s_reqd, False, ha1, hprob, hswap)
         If (pr >= crit_prob) Then
            crit_neghypergeometric = i
         End If
      Else
         crit_neghypergeometric = i + 1#
      End If
   End If
   crit_neghypergeometric = GetRidOfMinusZeroes(crit_neghypergeometric)
End Function

Public Function comp_crit_neghypergeometric(ByVal type1s_reqd As Double, ByVal tot_type1 As Double, ByVal POP_SIZE As Double, ByVal crit_prob As Double) As Double
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_neghypergeometric = [#VALUE!]
   ElseIf (type1s_reqd <= 0# Or tot_type1 < type1s_reqd Or tot_type1 > POP_SIZE) Then
      comp_crit_neghypergeometric = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_neghypergeometric = [#VALUE!]
   ElseIf (crit_prob = 0# Or POP_SIZE = tot_type1) Then
      comp_crit_neghypergeometric = POP_SIZE - tot_type1
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      Dim i As Double, pr As Double
      comp_crit_neghypergeometric = critcompneghyperg(type1s_reqd, tot_type1, POP_SIZE, crit_prob)
      i = comp_crit_neghypergeometric
      pr = hypergeometric(i, type1s_reqd, POP_SIZE - tot_type1 - i, tot_type1 - type1s_reqd, True, ha1, hprob, hswap)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         pr = hypergeometric(i, type1s_reqd, POP_SIZE - tot_type1 - i, tot_type1 - type1s_reqd, True, ha1, hprob, hswap)
         If (pr <= crit_prob) Then
            comp_crit_neghypergeometric = i
         End If
      Else
         comp_crit_neghypergeometric = i + 1#
      End If
   End If
   comp_crit_neghypergeometric = GetRidOfMinusZeroes(comp_crit_neghypergeometric)
End Function

Public Function lcb_neghypergeometric(ByVal type2s As Double, ByVal type1s_reqd As Double, ByVal POP_SIZE As Double, ByVal prob As Double) As Double
   type2s = AlterForIntegralChecks_Others(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (prob < 0# Or prob > 1#) Then
      lcb_neghypergeometric = [#VALUE!]
   ElseIf (type1s_reqd <= 0# Or type1s_reqd > POP_SIZE Or type2s > POP_SIZE - type1s_reqd) Then
      lcb_neghypergeometric = [#VALUE!]
   ElseIf (prob = 0# Or POP_SIZE = type2s + type1s_reqd) Then
      lcb_neghypergeometric = type1s_reqd
   ElseIf (prob = 1#) Then
      lcb_neghypergeometric = POP_SIZE - type2s
   ElseIf (prob < 0.5) Then
      lcb_neghypergeometric = critneghyperg(type1s_reqd, type2s + type1s_reqd, POP_SIZE, prob * (1.000000000001)) + type1s_reqd
   Else
      lcb_neghypergeometric = critcompneghyperg(type1s_reqd, type2s + type1s_reqd, POP_SIZE, (1# - prob) * (1# - 0.000000000001)) + type1s_reqd
   End If
   lcb_neghypergeometric = GetRidOfMinusZeroes(lcb_neghypergeometric)
End Function

Public Function ucb_neghypergeometric(ByVal type2s As Double, ByVal type1s_reqd As Double, ByVal POP_SIZE As Double, ByVal prob As Double) As Double
   type2s = AlterForIntegralChecks_Others(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   POP_SIZE = AlterForIntegralChecks_Others(POP_SIZE)
   If (prob < 0# Or prob > 1#) Then
      ucb_neghypergeometric = [#VALUE!]
   ElseIf (type1s_reqd <= 0# Or type1s_reqd > POP_SIZE Or type2s > POP_SIZE - type1s_reqd) Then
      ucb_neghypergeometric = [#VALUE!]
   ElseIf (prob = 0# Or type2s = 0# Or POP_SIZE = type2s + type1s_reqd) Then
      ucb_neghypergeometric = POP_SIZE - type2s
   ElseIf (prob = 1#) Then
      ucb_neghypergeometric = type1s_reqd
   ElseIf (prob < 0.5) Then
      ucb_neghypergeometric = critcompneghyperg(type1s_reqd, type2s + type1s_reqd - 1#, POP_SIZE, prob * (1# - 0.000000000001)) + type1s_reqd - 1#
   Else
      ucb_neghypergeometric = critneghyperg(type1s_reqd, type2s + type1s_reqd - 1#, POP_SIZE, (1# - prob) * (1.000000000001)) + type1s_reqd - 1#
   End If
   ucb_neghypergeometric = GetRidOfMinusZeroes(ucb_neghypergeometric)
End Function

Public Function pdf_exponential(ByVal x As Double, ByVal LAMBDA As Double) As Double
   If (LAMBDA <= 0#) Then
      pdf_exponential = [#VALUE!]
   ElseIf (x < 0#) Then
      pdf_exponential = 0#
   Else
      pdf_exponential = Exp(-LAMBDA * x + Log(LAMBDA))
   End If
   pdf_exponential = GetRidOfMinusZeroes(pdf_exponential)
End Function

Public Function cdf_exponential(ByVal x As Double, ByVal LAMBDA As Double) As Double
   If (LAMBDA <= 0#) Then
      cdf_exponential = [#VALUE!]
   ElseIf (x < 0#) Then
      cdf_exponential = 0#
   Else
      cdf_exponential = -expm1(-LAMBDA * x)
   End If
   cdf_exponential = GetRidOfMinusZeroes(cdf_exponential)
End Function

Public Function comp_cdf_exponential(ByVal x As Double, ByVal LAMBDA As Double) As Double
   If (LAMBDA <= 0#) Then
      comp_cdf_exponential = [#VALUE!]
   ElseIf (x < 0#) Then
      comp_cdf_exponential = 1#
   Else
      comp_cdf_exponential = Exp(-LAMBDA * x)
   End If
   comp_cdf_exponential = GetRidOfMinusZeroes(comp_cdf_exponential)
End Function

Public Function inv_exponential(ByVal prob As Double, ByVal LAMBDA As Double) As Double
   If (LAMBDA <= 0# Or prob < 0# Or prob >= 1#) Then
      inv_exponential = [#VALUE!]
   Else
      inv_exponential = -log0(-prob) / LAMBDA
   End If
   inv_exponential = GetRidOfMinusZeroes(inv_exponential)
End Function

Public Function comp_inv_exponential(ByVal prob As Double, ByVal LAMBDA As Double) As Double
   If (LAMBDA <= 0# Or prob <= 0# Or prob > 1#) Then
      comp_inv_exponential = [#VALUE!]
   Else
      comp_inv_exponential = -Log(prob) / LAMBDA
   End If
   comp_inv_exponential = GetRidOfMinusZeroes(comp_inv_exponential)
End Function

Public Function pdf_normal(ByVal x As Double) As Double
   If (Abs(x) < 40#) Then
      pdf_normal = Exp(-x * x * 0.5 - lstpi)
   Else
      pdf_normal = 0#
   End If
   pdf_normal = GetRidOfMinusZeroes(pdf_normal)
End Function

Public Function cdf_normal(ByVal x As Double) As Double
   cdf_normal = CNormal(x)
   cdf_normal = GetRidOfMinusZeroes(cdf_normal)
End Function

Public Function inv_normal(ByVal prob As Double) As Double
   If (prob <= 0# Or prob >= 1#) Then
      inv_normal = [#VALUE!]
   Else
      inv_normal = invcnormal(prob)
   End If
   inv_normal = GetRidOfMinusZeroes(inv_normal)
End Function

Public Function pdf_chi_sq(ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   pdf_chi_sq = pdf_gamma(x, df / 2#, 2#)
   pdf_chi_sq = GetRidOfMinusZeroes(pdf_chi_sq)
End Function

Public Function cdf_chi_sq(ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0#) Then
      cdf_chi_sq = [#VALUE!]
   ElseIf (x <= 0#) Then
      cdf_chi_sq = 0#
   Else
      cdf_chi_sq = GAMMA(x / 2#, df / 2#)
   End If
   cdf_chi_sq = GetRidOfMinusZeroes(cdf_chi_sq)
End Function

Public Function comp_cdf_chi_sq(ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0#) Then
      comp_cdf_chi_sq = [#VALUE!]
   ElseIf (x <= 0#) Then
      comp_cdf_chi_sq = 1#
   Else
      comp_cdf_chi_sq = compgamma(x / 2#, df / 2#)
   End If
   comp_cdf_chi_sq = GetRidOfMinusZeroes(comp_cdf_chi_sq)
End Function

Public Function inv_chi_sq(ByVal prob As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0# Or prob < 0# Or prob >= 1#) Then
      inv_chi_sq = [#VALUE!]
   ElseIf (prob = 0#) Then
      inv_chi_sq = 0#
   Else
      inv_chi_sq = 2# * invgamma(df / 2#, prob)
   End If
   inv_chi_sq = GetRidOfMinusZeroes(inv_chi_sq)
End Function

Public Function comp_inv_chi_sq(ByVal prob As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0# Or prob <= 0# Or prob > 1#) Then
      comp_inv_chi_sq = [#VALUE!]
   ElseIf (prob = 1#) Then
      comp_inv_chi_sq = 0#
   Else
      comp_inv_chi_sq = 2# * invcompgamma(df / 2#, prob)
   End If
   comp_inv_chi_sq = GetRidOfMinusZeroes(comp_inv_chi_sq)
End Function

Public Function pdf_gamma(ByVal x As Double, ByVal shape_param As Double, ByVal scale_param As Double) As Double
   Dim XS As Double
   If (shape_param <= 0# Or scale_param <= 0#) Then
      pdf_gamma = [#VALUE!]
   ElseIf (x < 0#) Then
      pdf_gamma = 0#
   ElseIf (x = 0#) Then
      If (shape_param < 1#) Then
         pdf_gamma = [#VALUE!]
      ElseIf (shape_param = 1#) Then
         pdf_gamma = 1# / scale_param
      Else
         pdf_gamma = 0#
      End If
   Else
      XS = x / scale_param
      pdf_gamma = poissonTerm(shape_param, XS, XS - shape_param, Log(shape_param) - Log(x))
   End If
   pdf_gamma = GetRidOfMinusZeroes(pdf_gamma)
End Function

Public Function cdf_gamma(ByVal x As Double, ByVal shape_param As Double, ByVal scale_param As Double) As Double
   If (shape_param <= 0# Or scale_param <= 0#) Then
      cdf_gamma = [#VALUE!]
   ElseIf (x <= 0#) Then
      cdf_gamma = 0#
   Else
      cdf_gamma = GAMMA(x / scale_param, shape_param)
   End If
   cdf_gamma = GetRidOfMinusZeroes(cdf_gamma)
End Function

Public Function comp_cdf_gamma(ByVal x As Double, ByVal shape_param As Double, ByVal scale_param As Double) As Double
   If (shape_param <= 0# Or scale_param <= 0#) Then
      comp_cdf_gamma = [#VALUE!]
   ElseIf (x <= 0#) Then
      comp_cdf_gamma = 1#
   Else
      comp_cdf_gamma = compgamma(x / scale_param, shape_param)
   End If
   comp_cdf_gamma = GetRidOfMinusZeroes(comp_cdf_gamma)
End Function

Public Function inv_gamma(ByVal prob As Double, ByVal shape_param As Double, ByVal scale_param As Double) As Double
   If (shape_param <= 0# Or scale_param <= 0# Or prob < 0# Or prob >= 1#) Then
      inv_gamma = [#VALUE!]
   ElseIf (prob = 0#) Then
      inv_gamma = 0#
   Else
      inv_gamma = scale_param * invgamma(shape_param, prob)
   End If
   inv_gamma = GetRidOfMinusZeroes(inv_gamma)
End Function

Public Function comp_inv_gamma(ByVal prob As Double, ByVal shape_param As Double, ByVal scale_param As Double) As Double
   If (shape_param <= 0# Or scale_param <= 0# Or prob <= 0# Or prob > 1#) Then
      comp_inv_gamma = [#VALUE!]
   ElseIf (prob = 1#) Then
      comp_inv_gamma = 0#
   Else
      comp_inv_gamma = scale_param * invcompgamma(shape_param, prob)
   End If
   comp_inv_gamma = GetRidOfMinusZeroes(comp_inv_gamma)
End Function

Private Function pdftdist(ByVal x As Double, ByVal k As Double) As Double
'//Probability density for a variate from t-distribution with k degress of freedom
   Dim x2 As Double, k2 As Double, logterm As Double
   If (k <= 0#) Then
      pdftdist = [#VALUE!]
   ElseIf (k > 1E+30) Then
      pdftdist = pdf_normal(x)
   Else
      If Abs(x) >= Min(1#, k) Then
         k2 = k / x
         x2 = x + k2
         k2 = k2 / x2
         x2 = x / x2
      Else
         x2 = x * x
         k2 = k + x2
         x2 = x2 / k2
         k2 = k / k2
      End If
      If (k2 < cSmall) Then
         logterm = Log(k) - 2# * Log(Abs(x))
      ElseIf (Abs(x2) < 0.5) Then
         logterm = log0(-x2)
      Else
         logterm = Log(k2)
      End If
      x2 = k * 0.5
      pdftdist = Exp(0.5 + (k + 1#) * 0.5 * logterm + x2 * log0(-0.5 / (x2 + 1)) + logfbit(x2 - 0.5) - logfbit(x2)) * Sqr(x2 / ((1# + x2))) * OneOverSqrTwoPi
   End If
End Function

Public Function pdf_tdist(ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   pdf_tdist = pdftdist(x, df)
   pdf_tdist = GetRidOfMinusZeroes(pdf_tdist)
End Function

Public Function cdf_tdist(ByVal x As Double, ByVal df As Double) As Double
   Dim tdistDensity As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0#) Then
      cdf_tdist = [#VALUE!]
   Else
      cdf_tdist = tdist(x, df, tdistDensity)
   End If
   cdf_tdist = GetRidOfMinusZeroes(cdf_tdist)
End Function

Public Function inv_tdist(ByVal prob As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0#) Then
      inv_tdist = [#VALUE!]
   ElseIf (prob <= 0# Or prob >= 1#) Then
      inv_tdist = [#VALUE!]
   Else
      inv_tdist = invtdist(prob, df)
   End If
   inv_tdist = GetRidOfMinusZeroes(inv_tdist)
End Function

Public Function pdf_fdist(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0#) Then
      pdf_fdist = [#VALUE!]
   ElseIf (x < 0#) Then
      pdf_fdist = 0#
   ElseIf (x = 0# And df1 > 2#) Then
      pdf_fdist = 0#
   ElseIf (x = 0# And df1 < 2#) Then
      pdf_fdist = [#VALUE!]
   ElseIf (x = 0#) Then
      pdf_fdist = 1#
   Else
      Dim p As Double, q As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      'If p < cSmall And x <> 0# Or q < cSmall Then
      '   pdf_fdist = [#VALUE!]
      '   Exit Function
      'End If
      df2 = df2 / 2#
      df1 = df1 / 2#
      If (df1 >= 1#) Then
         df1 = df1 - 1#
         pdf_fdist = binomialTerm(df1, df2, p, q, df2 * p - df1 * q, Log((df1 + 1#) * q))
      Else
         pdf_fdist = df1 * df1 * q / (p * (df1 + df2)) * binomialTerm(df1, df2, p, q, df2 * p - df1 * q, 0#)
      End If
   End If
   pdf_fdist = GetRidOfMinusZeroes(pdf_fdist)
End Function

Public Function cdf_fdist(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0#) Then
      cdf_fdist = [#VALUE!]
   ElseIf (x <= 0#) Then
      cdf_fdist = 0#
   Else
      Dim p As Double, q As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      'If p < cSmall And x <> 0# Or q < cSmall Then
      '   cdf_fdist = [#VALUE!]
      '   Exit Function
      'End If
      df2 = df2 / 2#
      df1 = df1 / 2#
      If (p < 0.5) Then
          cdf_fdist = beta(p, df1, df2)
      Else
          cdf_fdist = compbeta(q, df2, df1)
      End If
   End If
   cdf_fdist = GetRidOfMinusZeroes(cdf_fdist)
End Function

Public Function comp_cdf_fdist(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0#) Then
      comp_cdf_fdist = [#VALUE!]
   ElseIf (x <= 0#) Then
      comp_cdf_fdist = 1#
   Else
      Dim p As Double, q As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      'If p < cSmall And x <> 0# Or q < cSmall Then
      '   comp_cdf_fdist = [#VALUE!]
      '   Exit Function
      'End If
      df2 = df2 / 2#
      df1 = df1 / 2#
      If (p < 0.5) Then
          comp_cdf_fdist = compbeta(p, df1, df2)
      Else
          comp_cdf_fdist = beta(q, df2, df1)
      End If
   End If
   comp_cdf_fdist = GetRidOfMinusZeroes(comp_cdf_fdist)
End Function

Public Function inv_fdist(ByVal prob As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or prob < 0# Or prob >= 1#) Then
      inv_fdist = [#VALUE!]
   ElseIf (prob = 0#) Then
      inv_fdist = 0#
   Else
      Dim temp As Double, oneMinusP As Double
      df1 = df1 / 2#
      df2 = df2 / 2#
      temp = invbeta(df1, df2, prob, oneMinusP)
      inv_fdist = df2 * temp / (df1 * oneMinusP)
      'If oneMinusP < cSmall Then inv_fdist = [#VALUE!]
   End If
   inv_fdist = GetRidOfMinusZeroes(inv_fdist)
End Function

Public Function comp_inv_fdist(ByVal prob As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or prob <= 0# Or prob > 1#) Then
      comp_inv_fdist = [#VALUE!]
   ElseIf (prob = 1#) Then
      comp_inv_fdist = 0#
   Else
      Dim temp As Double, oneMinusP As Double
      df1 = df1 / 2#
      df2 = df2 / 2#
      temp = invcompbeta(df1, df2, prob, oneMinusP)
      comp_inv_fdist = df2 * temp / (df1 * oneMinusP)
      'If oneMinusP < cSmall Then comp_inv_fdist = [#VALUE!]
   End If
   comp_inv_fdist = GetRidOfMinusZeroes(comp_inv_fdist)
End Function

Public Function pdf_BETA(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
   If (shape_param1 <= 0# Or shape_param2 <= 0#) Then
      pdf_BETA = [#VALUE!]
   ElseIf (x < 0# Or x > 1#) Then
      pdf_BETA = 0#
   ElseIf (x = 0# And shape_param1 < 1# Or x = 1# And shape_param2 < 1#) Then
      pdf_BETA = [#VALUE!]
   ElseIf (x = 0# And shape_param1 = 1#) Then
      pdf_BETA = shape_param2
   ElseIf (x = 1# And shape_param2 = 1#) Then
      pdf_BETA = shape_param1
   ElseIf ((x = 0#) Or (x = 1#)) Then
      pdf_BETA = 0#
   Else
      Dim MX As Double, mn As Double
      MX = max(shape_param1, shape_param2)
      mn = Min(shape_param1, shape_param2)
      pdf_BETA = binomialTerm(shape_param1, shape_param2, x, 1# - x, (shape_param1 + shape_param2) * x - shape_param1, Log(MX / (mn + MX)) + Log(mn) - Log(x * (1# - x)))
   End If
   pdf_BETA = GetRidOfMinusZeroes(pdf_BETA)
End Function

Public Function cdf_BETA(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
   If (shape_param1 <= 0# Or shape_param2 <= 0#) Then
      cdf_BETA = [#VALUE!]
   ElseIf (x <= 0#) Then
      cdf_BETA = 0#
   ElseIf (x >= 1#) Then
      cdf_BETA = 1#
   Else
      cdf_BETA = beta(x, shape_param1, shape_param2)
   End If
   cdf_BETA = GetRidOfMinusZeroes(cdf_BETA)
End Function

Public Function comp_cdf_BETA(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
   If (shape_param1 <= 0# Or shape_param2 <= 0#) Then
      comp_cdf_BETA = [#VALUE!]
   ElseIf (x <= 0#) Then
      comp_cdf_BETA = 1#
   ElseIf (x >= 1#) Then
      comp_cdf_BETA = 0#
   Else
      comp_cdf_BETA = compbeta(x, shape_param1, shape_param2)
   End If
   comp_cdf_BETA = GetRidOfMinusZeroes(comp_cdf_BETA)
End Function

Public Function inv_BETA(ByVal prob As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
   If (shape_param1 <= 0# Or shape_param2 <= 0# Or prob < 0# Or prob > 1#) Then
      inv_BETA = [#VALUE!]
   Else
      Dim oneMinusP As Double
      inv_BETA = invbeta(shape_param1, shape_param2, prob, oneMinusP)
   End If
   inv_BETA = GetRidOfMinusZeroes(inv_BETA)
End Function

Public Function comp_inv_BETA(ByVal prob As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
   If (shape_param1 <= 0# Or shape_param2 <= 0# Or prob < 0# Or prob > 1#) Then
      comp_inv_BETA = [#VALUE!]
   Else
      Dim oneMinusP As Double
      comp_inv_BETA = invcompbeta(shape_param1, shape_param2, prob, oneMinusP)
   End If
   comp_inv_BETA = GetRidOfMinusZeroes(comp_inv_BETA)
End Function

Private Function gamma_nc1(ByVal x As Double, ByVal A As Double, ByVal nc As Double, ByRef nc_derivative As Double) As Double
   Dim AA As Double, BB As Double, nc_dtemp As Double
   Dim n As Double, p As Double, W As Double, s As Double, ps As Double
   Dim Result As Double, term As Double, ptx As Double, ptnc As Double
   n = A + Sqr(A ^ 2 + 4# * nc * x)
   If n > 0# Then n = Int(2# * nc * x / n)
   AA = n + A
   BB = n
   ptnc = poissonTerm(n, nc, nc - n, 0#)
   ptx = poissonTerm(AA, x, x - AA, 0#)
   AA = AA + 1#
   BB = BB + 1#
   p = nc / BB
   ps = p
   nc_derivative = ps
   s = x / AA
   W = p
   term = s * W
   Result = term
   If ptx > 0# Then
     While (((term > 0.000000000000001 * Result) And (p > 1E-16 * W)) Or (ps > 1E-16 * nc_derivative))
       AA = AA + 1#
       BB = BB + 1#
       p = nc / BB * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / AA * s
       W = W + p
       term = s * W
       Result = Result + term
     Wend
     W = W * ptnc
   Else
     W = comppoisson(n, nc, nc - n)
   End If
   gamma_nc1 = Result * ptx * ptnc + comppoisson(A + BB, x, (x - A) - BB) * W
   ps = 1#
   nc_dtemp = 0#
   AA = n + A
   BB = n
   p = 1#
   s = ptx
   W = GAMMA(x, AA)
   term = p * W
   Result = term
   While BB > 0# And ((term > 0.000000000000001 * Result) Or (ps > 1E-16 * nc_dtemp))
       s = AA / x * s
       ps = p * s
       nc_dtemp = nc_dtemp + ps
       p = BB / nc * p
       W = W + s
       term = p * W
       Result = Result + term
       AA = AA - 1#
       BB = BB - 1#
   Wend
   If BB = 0# Then AA = A
   If n > 0# Then
      nc_dtemp = nc_derivative * ptx + nc_dtemp + p * AA / x * s
   Else
      nc_dtemp = poissonTerm(AA, x, x - AA, Log(nc_derivative * x + AA) - Log(x))
   End If
   gamma_nc1 = gamma_nc1 + Result * ptnc + CPoisson(BB - 1#, nc, nc - BB + 1#) * W
   If nc_dtemp = 0# Then
      nc_derivative = 0#
   Else
      nc_derivative = poissonTerm(n, nc, nc - n, Log(nc_dtemp))
   End If
End Function

Private Function comp_gamma_nc1(ByVal x As Double, ByVal A As Double, ByVal nc As Double, ByRef nc_derivative As Double) As Double
   Dim AA As Double, BB As Double, nc_dtemp As Double
   Dim n As Double, p As Double, W As Double, s As Double, ps As Double
   Dim Result As Double, term As Double, ptx As Double, ptnc As Double
   n = A + Sqr(A ^ 2 + 4# * nc * x)
   If n > 0# Then n = Int(2# * nc * x / n)
   AA = n + A
   BB = n
   ptnc = poissonTerm(n, nc, nc - n, 0#)
   ptx = poissonTerm(AA, x, x - AA, 0#)
   s = 1#
   ps = 1#
   nc_dtemp = 0#
   p = 1#
   W = p
   term = 1#
   Result = 0#
   If ptx > 0# Then
     While BB > 0# And (((term > 0.000000000000001 * Result) And (p > 1E-16 * W)) Or (ps > 1E-16 * nc_dtemp))
      s = AA / x * s
      ps = p * s
      nc_dtemp = nc_dtemp + ps
      p = BB / nc * p
      term = s * W
      Result = Result + term
      W = W + p
      AA = AA - 1#
      BB = BB - 1#
     Wend
     W = W * ptnc
   Else
     W = CPoisson(n, nc, nc - n)
   End If
   If BB = 0# Then AA = A
   If n > 0# Then
      nc_dtemp = (nc_dtemp + p * AA / x * s) * ptx
   ElseIf AA = 0 And x > 0 Then
      nc_dtemp = 0#
   Else
      nc_dtemp = poissonTerm(AA, x, x - AA, Log(AA) - Log(x))
   End If
   comp_gamma_nc1 = Result * ptx * ptnc + compgamma(x, AA) * W
   AA = n + A
   BB = n
   ps = 1#
   nc_derivative = 0#
   p = 1#
   s = ptx
   W = compgamma(x, AA)
   term = 0#
   Result = term
   Do
       W = W + s
       AA = AA + 1#
       BB = BB + 1#
       p = nc / BB * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / AA * s
       term = p * W
       Result = Result + term
   Loop While (((term > 0.000000000000001 * Result) And (s > 1E-16 * W)) Or (ps > 1E-16 * nc_derivative))
   comp_gamma_nc1 = comp_gamma_nc1 + Result * ptnc + comppoisson(BB, nc, nc - BB) * W
   nc_dtemp = nc_derivative + nc_dtemp
   If nc_dtemp = 0# Then
      nc_derivative = 0#
   Else
      nc_derivative = poissonTerm(n, nc, nc - n, Log(nc_dtemp))
   End If
End Function

Private Function inv_gamma_nc1(ByVal prob As Double, ByVal A As Double, ByVal nc As Double) As Double
'Uses approx in A&S 26.4.27 for to get initial estimate the modified NR to improve it.
Dim x As Double, pr As Double, dif As Double
Dim hi As Double, lo As Double, nc_derivative As Double
   If (prob > 0.5) Then
      inv_gamma_nc1 = comp_inv_gamma_nc1(1# - prob, A, nc)
      Exit Function
   End If

   lo = 0#
   hi = 1E+308
   pr = Exp(-nc)
   If pr > prob Then
      If 2# * prob > pr Then
         x = comp_inv_gamma((pr - prob) / pr, A + cSmall, 1#)
      Else
         x = inv_gamma(prob / pr, A + cSmall, 1#)
      End If
      If x < cSmall Then
         x = cSmall
         pr = gamma_nc1(x, A, nc, nc_derivative)
         If pr > prob Then
            inv_gamma_nc1 = 0#
            Exit Function
         End If
      End If
   Else
      x = inv_gamma(prob, (A + nc) / (1# + nc / (A + nc)), 1#)
      x = x * (1# + nc / (A + nc))
   End If
   dif = x
   Do
      pr = gamma_nc1(x, A, nc, nc_derivative)
      If pr < 3E-308 And nc_derivative = 0# Then
         lo = x
         dif = dif / 2#
         x = x - dif
      ElseIf nc_derivative = 0# Then
         hi = x
         dif = dif / 2#
         x = x - dif
      Else
         If pr < prob Then
            lo = x
         Else
            hi = x
         End If
         dif = -(pr / nc_derivative) * logdif(pr, prob)
         If x + dif < lo Then
            dif = (lo - x) / 2#
         ElseIf x + dif > hi Then
            dif = (hi - x) / 2#
         End If
         x = x + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(x) * 0.0000000001))
   inv_gamma_nc1 = x
End Function

Private Function comp_inv_gamma_nc1(ByVal prob As Double, ByVal A As Double, ByVal nc As Double) As Double
'Uses approx in A&S 26.4.27 for to get initial estimate the modified NR to improve it.
Dim x As Double, pr As Double, dif As Double
Dim hi As Double, lo As Double, nc_derivative As Double
   If (prob > 0.5) Then
      comp_inv_gamma_nc1 = inv_gamma_nc1(1# - prob, A, nc)
      Exit Function
   End If

   lo = 0#
   hi = 1E+308
   pr = Exp(-nc)
   If pr > prob Then
      x = comp_inv_gamma(prob / pr, A + cSmall, 1#) ' Is this as small as x could be?
   Else
      x = comp_inv_gamma(prob, (A + nc) / (1# + nc / (A + nc)), 1#)
      x = x * (1# + nc / (A + nc))
   End If
   If x < cSmall Then x = cSmall
   dif = x
   Do
      pr = comp_gamma_nc1(x, A, nc, nc_derivative)
      If pr < 3E-308 And nc_derivative = 0# Then
         hi = x
         dif = dif / 2#
         x = x - dif
      ElseIf nc_derivative = 0# Then
         lo = x
         dif = dif / 2#
         x = x - dif
      Else
         If pr < prob Then
            hi = x
         Else
            lo = x
         End If
         dif = (pr / nc_derivative) * logdif(pr, prob)
         If x + dif < lo Then
            dif = (lo - x) / 2#
         ElseIf x + dif > hi Then
            dif = (hi - x) / 2#
         End If
         x = x + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(x) * 0.0000000001))
   comp_inv_gamma_nc1 = x
End Function

Private Function ncp_gamma_nc1(ByVal prob As Double, ByVal x As Double, ByVal A As Double) As Double
'Uses Normal approx for difference of 2 poisson distributed variables  to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, deriv As Double, B As Double, sqarg As Double, checked_nc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double
   If (prob > 0.5) Then
      ncp_gamma_nc1 = comp_ncp_gamma_nc1(1# - prob, x, A)
      Exit Function
   End If

   lo = 0#
   hi = nc_limit
   checked_0_limit = False
   checked_nc_limit = False
   temp = inv_normal(prob) ^ 2
   B = 2# * (x - A) + temp
   sqarg = B ^ 2 - 4 * ((x - A) ^ 2 - temp * x)
   If sqarg < 0 Then
      ncp = B / 2
   Else
      ncp = (B + Sqr(sqarg)) / 2
   End If
   ncp = max(0#, Min(ncp, nc_limit))
   If ncp = 0# Then
      pr = cdf_gamma_nc(x, A, 0#)
      If pr < prob Then
         If (inv_gamma(prob, A, 1) <= x) Then
            ncp_gamma_nc1 = 0#
         Else
            ncp_gamma_nc1 = [#VALUE!]
         End If
         Exit Function
      Else
         checked_0_limit = True
      End If
   ElseIf ncp = nc_limit Then
      pr = cdf_gamma_nc(x, A, ncp)
      If pr > prob Then
         ncp_gamma_nc1 = [#VALUE!]
         Exit Function
      Else
         checked_nc_limit = True
      End If
   End If
   dif = ncp
   Do
      pr = cdf_gamma_nc(x, A, ncp)
      'Debug.Print ncp, pr, prob
      deriv = pdf_gamma_nc(x, A + 1#, ncp)
      If pr < 3E-308 And deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      ElseIf deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr < prob Then
            hi = ncp
         Else
            lo = ncp
         End If
         dif = (pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = cdf_gamma_nc(x, A, lo)
               If temp < prob Then
                  If (inv_gamma(prob, A, 1) <= x) Then
                     ncp_gamma_nc1 = 0#
                  Else
                     ncp_gamma_nc1 = [#VALUE!]
                  End If
                  Exit Function
               Else
                  checked_0_limit = True
               End If
            End If
         ElseIf ncp + dif > hi Then
            dif = (hi - ncp) / 2#
            If Not checked_nc_limit And (hi = nc_limit) Then
               pr = cdf_gamma_nc(x, A, hi)
               If pr > prob Then
                  ncp_gamma_nc1 = [#VALUE!]
                  Exit Function
               Else
                  ncp = hi
                  deriv = pdf_gamma_nc(x, A + 1#, ncp)
                  dif = (pr / deriv) * logdif(pr, prob)
                  If ncp + dif < lo Then
                     dif = (lo - ncp) / 2#
                  End If
                  checked_nc_limit = True
               End If
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   ncp_gamma_nc1 = ncp
   'Debug.Print "ncp_gamma_nc1", ncp_gamma_nc1
End Function

Private Function comp_ncp_gamma_nc1(ByVal prob As Double, ByVal x As Double, ByVal A As Double) As Double
'Uses Normal approx for difference of 2 poisson distributed variables  to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, deriv As Double, B As Double, sqarg As Double, checked_nc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double
   If (prob > 0.5) Then
      comp_ncp_gamma_nc1 = ncp_gamma_nc1(1# - prob, x, A)
      Exit Function
   End If

   lo = 0#
   hi = nc_limit
   checked_0_limit = False
   checked_nc_limit = False
   temp = inv_normal(prob) ^ 2
   B = 2# * (x - A) + temp
   sqarg = B ^ 2 - 4 * ((x - A) ^ 2 - temp * x)
   If sqarg < 0 Then
      ncp = B / 2
   Else
      ncp = (B - Sqr(sqarg)) / 2
   End If
   ncp = max(0#, ncp)
   If ncp <= 1# Then
      pr = comp_cdf_gamma_nc(x, A, 0#)
      If pr > prob Then
         If (comp_inv_gamma(prob, A, 1) <= x) Then
            comp_ncp_gamma_nc1 = 0#
         Else
            comp_ncp_gamma_nc1 = [#VALUE!]
         End If
         Exit Function
      Else
         checked_0_limit = True
      End If
      deriv = pdf_gamma_nc(x, A + 1#, ncp)
      If deriv = 0# Then
         ncp = nc_limit
      ElseIf A < 1 Then
         ncp = (prob - pr) / deriv
         If ncp >= nc_limit Then
            ncp = -(pr / deriv) * logdif(pr, prob)
         End If
      Else
         ncp = -(pr / deriv) * logdif(pr, prob)
      End If
   End If
   ncp = Min(ncp, nc_limit)
   If ncp = nc_limit Then
      pr = comp_cdf_gamma_nc(x, A, ncp)
      If pr < prob Then
         comp_ncp_gamma_nc1 = [#VALUE!]
         Exit Function
      Else
         deriv = pdf_gamma_nc(x, A + 1#, ncp)
         dif = -(pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
         End If
         checked_nc_limit = True
      End If
   End If
   dif = ncp
   Do
      pr = comp_cdf_gamma_nc(x, A, ncp)
      'Debug.Print ncp, pr, prob
      deriv = pdf_gamma_nc(x, A + 1#, ncp)
      If pr < 3E-308 And deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp + dif
      ElseIf deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr < prob Then
            lo = ncp
         Else
            hi = ncp
         End If
         dif = -(pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = comp_cdf_gamma_nc(x, A, lo)
               If temp > prob Then
                  If (comp_inv_gamma(prob, A, 1) <= x) Then
                     comp_ncp_gamma_nc1 = 0#
                  Else
                     comp_ncp_gamma_nc1 = [#VALUE!]
                  End If
                  Exit Function
               Else
                  checked_0_limit = True
               End If
            End If
         ElseIf ncp + dif > hi Then
            If Not checked_nc_limit And (hi = nc_limit) Then
               ncp = hi
               pr = comp_cdf_gamma_nc(x, A, ncp)
               If pr < prob Then
                  comp_ncp_gamma_nc1 = [#VALUE!]
                  Exit Function
               Else
                  deriv = pdf_gamma_nc(x, A + 1#, ncp)
                  dif = -(pr / deriv) * logdif(pr, prob)
                  If ncp + dif < lo Then
                     dif = (lo - ncp) / 2#
                  End If
                  checked_nc_limit = True
               End If
            Else
               dif = (hi - ncp) / 2#
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   comp_ncp_gamma_nc1 = ncp
   'Debug.Print "comp_ncp_gamma_nc1", comp_ncp_gamma_nc1
End Function

Public Function pdf_gamma_nc(ByVal x As Double, ByVal shape_param As Double, ByVal nc_param As Double) As Double
'// Calculate pdf of noncentral gamma
  Dim nc_derivative As Double
  If (shape_param < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Then
     pdf_gamma_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     pdf_gamma_nc = 0#
  ElseIf (shape_param = 0# And nc_param = 0# And x > 0#) Then
     pdf_gamma_nc = 0#
  ElseIf (x = 0# Or nc_param = 0#) Then
     pdf_gamma_nc = Exp(-nc_param) * pdf_gamma(x, shape_param, 1#)
  ElseIf shape_param >= 1# Then
     If x >= nc_param Then
        If (x < 1# Or x <= shape_param + nc_param) Then
           pdf_gamma_nc = gamma_nc1(x, shape_param, nc_param, nc_derivative)
        Else
           pdf_gamma_nc = comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
        End If
        pdf_gamma_nc = nc_derivative
     Else
        If (nc_param < 1# Or nc_param <= shape_param + x) Then
           pdf_gamma_nc = gamma_nc1(nc_param, shape_param, x, nc_derivative)
        Else
           pdf_gamma_nc = comp_gamma_nc1(nc_param, shape_param, x, nc_derivative)
        End If
        If nc_derivative = 0# Then
           pdf_gamma_nc = 0#
        Else
           pdf_gamma_nc = Exp(Log(nc_derivative) + (shape_param - 1#) * (Log(x) - Log(nc_param)))
        End If
     End If
  Else
     If x < nc_param Then
        If (x < 1# Or x <= shape_param + nc_param) Then
           pdf_gamma_nc = gamma_nc1(x, shape_param, nc_param, nc_derivative)
        Else
           pdf_gamma_nc = comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
        End If
        pdf_gamma_nc = nc_derivative
     Else
        If (nc_param < 1# Or nc_param <= shape_param + x) Then
           pdf_gamma_nc = gamma_nc1(nc_param, shape_param, x, nc_derivative)
        Else
           pdf_gamma_nc = comp_gamma_nc1(nc_param, shape_param, x, nc_derivative)
        End If
        If nc_derivative = 0# Then
           pdf_gamma_nc = 0#
        Else
           pdf_gamma_nc = Exp(Log(nc_derivative) + (shape_param - 1#) * (Log(x) - Log(nc_param)))
        End If
     End If
  End If
  pdf_gamma_nc = GetRidOfMinusZeroes(pdf_gamma_nc)
End Function

Public Function cdf_gamma_nc(ByVal x As Double, ByVal shape_param As Double, ByVal nc_param As Double) As Double
'// Calculate cdf of noncentral gamma
  Dim nc_derivative As Double
  If (shape_param < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Then
     cdf_gamma_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     cdf_gamma_nc = 0#
  ElseIf (x = 0# And shape_param = 0#) Then
     cdf_gamma_nc = Exp(-nc_param)
  ElseIf (shape_param + nc_param = 0#) Then    ' limit as shape_param+nc_param->0 is degenerate point mass at zero
     cdf_gamma_nc = 1#                         ' if fix central gamma, then works for degenerate poisson
  ElseIf (x = 0#) Then
     cdf_gamma_nc = 0#
  ElseIf (nc_param = 0#) Then
     cdf_gamma_nc = GAMMA(x, shape_param)
  'ElseIf (shape_param = 0#) Then              ' extends Ruben (1974) and Cohen (1988) recurrence
  '   cdf_gamma_nc = ((x + shape_param + 2#) * gamma_nc1(x, shape_param + 2#, nc_param) + (nc_param - shape_param - 2#) * gamma_nc1(x, shape_param + 4#, nc_param) - nc_param * gamma_nc1(x, shape_param + 6#, nc_param)) / x
  ElseIf (x < 1# Or x <= shape_param + nc_param) Then
     cdf_gamma_nc = gamma_nc1(x, shape_param, nc_param, nc_derivative)
  Else
     cdf_gamma_nc = 1# - comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
  End If
  cdf_gamma_nc = GetRidOfMinusZeroes(cdf_gamma_nc)
End Function

Public Function comp_cdf_gamma_nc(ByVal x As Double, ByVal shape_param As Double, ByVal nc_param As Double) As Double
'// Calculate 1-cdf of noncentral gamma
  Dim nc_derivative As Double
  If (shape_param < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Then
     comp_cdf_gamma_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     comp_cdf_gamma_nc = 1#
  ElseIf (x = 0# And shape_param = 0#) Then
     comp_cdf_gamma_nc = -expm1(-nc_param)
  ElseIf (shape_param + nc_param = 0#) Then     ' limit as shape_param+nc_param->0 is degenerate point mass at zero
     comp_cdf_gamma_nc = 0#                     ' if fix central gamma, then works for degenerate poisson
  ElseIf (x = 0#) Then
     comp_cdf_gamma_nc = 1
  ElseIf (nc_param = 0#) Then
     comp_cdf_gamma_nc = compgamma(x, shape_param)
  'ElseIf (shape_param = 0#) Then              ' extends Ruben (1974) and Cohen (1988) recurrence
  '   comp_cdf_gamma_nc = ((x + shape_param + 2#) * comp_gamma_nc1(x, shape_param + 2#, nc_param) + (nc_param - shape_param - 2#) * comp_gamma_nc1(x, shape_param + 4#, nc_param) - nc_param * comp_gamma_nc1(x, shape_param + 6#, nc_param)) / x
  ElseIf (x < 1# Or x >= shape_param + nc_param) Then
     comp_cdf_gamma_nc = comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
  Else
     comp_cdf_gamma_nc = 1# - gamma_nc1(x, shape_param, nc_param, nc_derivative)
  End If
  comp_cdf_gamma_nc = GetRidOfMinusZeroes(comp_cdf_gamma_nc)
End Function

Public Function inv_gamma_nc(ByVal prob As Double, ByVal shape_param As Double, ByVal nc_param As Double) As Double
   If (shape_param < 0# Or nc_param < 0# Or nc_param > nc_limit Or prob < 0# Or prob >= 1#) Then
      inv_gamma_nc = [#VALUE!]
   ElseIf (prob = 0# Or shape_param = 0# And prob <= Exp(-nc_param)) Then
      inv_gamma_nc = 0#
   Else
      inv_gamma_nc = inv_gamma_nc1(prob, shape_param, nc_param)
   End If
   inv_gamma_nc = GetRidOfMinusZeroes(inv_gamma_nc)
End Function

Public Function comp_inv_gamma_nc(ByVal prob As Double, ByVal shape_param As Double, ByVal nc_param As Double) As Double
   If (shape_param < 0# Or nc_param < 0# Or nc_param > nc_limit Or prob <= 0# Or prob > 1#) Then
      comp_inv_gamma_nc = [#VALUE!]
   ElseIf (prob = 1# Or shape_param = 0# And prob >= -expm1(-nc_param)) Then
      comp_inv_gamma_nc = 0#
   Else
      comp_inv_gamma_nc = comp_inv_gamma_nc1(prob, shape_param, nc_param)
   End If
   comp_inv_gamma_nc = GetRidOfMinusZeroes(comp_inv_gamma_nc)
End Function

Public Function ncp_gamma_nc(ByVal prob As Double, ByVal x As Double, ByVal shape_param As Double) As Double
   If (shape_param < 0# Or x < 0# Or prob <= 0# Or prob > 1#) Then
      ncp_gamma_nc = [#VALUE!]
   ElseIf (x = 0# And shape_param = 0#) Then
      ncp_gamma_nc = -Log(prob)
   ElseIf (shape_param = 0# And prob = 1#) Then
      ncp_gamma_nc = 0#
   ElseIf (x = 0# Or prob = 1#) Then
      ncp_gamma_nc = [#VALUE!]
   Else
      ncp_gamma_nc = ncp_gamma_nc1(prob, x, shape_param)
   End If
   ncp_gamma_nc = GetRidOfMinusZeroes(ncp_gamma_nc)
End Function

Public Function comp_ncp_gamma_nc(ByVal prob As Double, ByVal x As Double, ByVal shape_param As Double) As Double
   If (shape_param < 0# Or x < 0# Or prob < 0# Or prob >= 1#) Then
      comp_ncp_gamma_nc = [#VALUE!]
   ElseIf (x = 0# And shape_param = 0#) Then
      comp_ncp_gamma_nc = -log0(-prob)
   ElseIf (shape_param = 0# And prob = 0#) Then
      comp_ncp_gamma_nc = 0#
   ElseIf (x = 0# Or prob = 0#) Then
      comp_ncp_gamma_nc = [#VALUE!]
   Else
      comp_ncp_gamma_nc = comp_ncp_gamma_nc1(prob, x, shape_param)
   End If
   comp_ncp_gamma_nc = GetRidOfMinusZeroes(comp_ncp_gamma_nc)
End Function

Public Function pdf_Chi2_nc(ByVal x As Double, ByVal df As Double, ByVal nc As Double) As Double
'// Calculate pdf of noncentral chi-square
  df = AlterForIntegralChecks_df(df)
  pdf_Chi2_nc = 0.5 * pdf_gamma_nc(x / 2#, df / 2#, nc / 2#)
  pdf_Chi2_nc = GetRidOfMinusZeroes(pdf_Chi2_nc)
End Function

Public Function cdf_Chi2_nc(ByVal x As Double, ByVal df As Double, ByVal nc As Double) As Double
'// Calculate cdf of noncentral chi-square
'//   parametrized per Johnson & Kotz, SAS, etc. so that cdf_Chi2_nc(x,df,nc) = cdf_gamma_nc(x/2,df/2,nc/2)
'//   If Xi ~ N(Di,1) independent, then sum(Xi,i=1..n) ~ Chi2_nc(n,nc) with nc=sum(Di,i=1..n)
'//   Note that Knusel, Graybill, etc. use a different noncentrality parameter lambda=nc/2
  df = AlterForIntegralChecks_df(df)
  cdf_Chi2_nc = cdf_gamma_nc(x / 2#, df / 2#, nc / 2#)
  cdf_Chi2_nc = GetRidOfMinusZeroes(cdf_Chi2_nc)
End Function

Public Function comp_cdf_Chi2_nc(ByVal x As Double, ByVal df As Double, ByVal nc As Double) As Double
'// Calculate 1-cdf of noncentral chi-square
'//   parametrized per Johnson & Kotz, SAS, etc. so that cdf_Chi2_nc(x,df,nc) = cdf_gamma_nc(x/2,df/2,nc/2)
'//   If Xi ~ N(Di,1) independent, then sum(Xi,i=1..n) ~ Chi2_nc(n,nc) with nc=sum(Di,i=1..n)
'//   Note that Knusel, Graybill, etc. use a different noncentrality parameter lambda=nc/2
  df = AlterForIntegralChecks_df(df)
  comp_cdf_Chi2_nc = comp_cdf_gamma_nc(x / 2#, df / 2#, nc / 2#)
  comp_cdf_Chi2_nc = GetRidOfMinusZeroes(comp_cdf_Chi2_nc)
End Function

Public Function inv_Chi2_nc(ByVal prob As Double, ByVal df As Double, ByVal nc As Double) As Double
   df = AlterForIntegralChecks_df(df)
   inv_Chi2_nc = 2# * inv_gamma_nc(prob, df / 2#, nc / 2#)
   inv_Chi2_nc = GetRidOfMinusZeroes(inv_Chi2_nc)
End Function

Public Function comp_inv_Chi2_nc(ByVal prob As Double, ByVal df As Double, ByVal nc As Double) As Double
   df = AlterForIntegralChecks_df(df)
   comp_inv_Chi2_nc = 2# * comp_inv_gamma_nc(prob, df / 2#, nc / 2#)
   comp_inv_Chi2_nc = GetRidOfMinusZeroes(comp_inv_Chi2_nc)
End Function

Public Function ncp_Chi2_nc(ByVal prob As Double, ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   ncp_Chi2_nc = 2# * ncp_gamma_nc(prob, x / 2#, df / 2#)
   ncp_Chi2_nc = GetRidOfMinusZeroes(ncp_Chi2_nc)
End Function

Public Function comp_ncp_Chi2_nc(ByVal prob As Double, ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   comp_ncp_Chi2_nc = 2# * comp_ncp_gamma_nc(prob, x / 2#, df / 2#)
   comp_ncp_Chi2_nc = GetRidOfMinusZeroes(comp_ncp_Chi2_nc)
End Function

Private Function BETA_nc1(ByVal x As Double, ByVal Y As Double, ByVal A As Double, ByVal B As Double, ByVal nc As Double, ByRef nc_derivative As Double) As Double
'y is 1-x but held accurately to avoid possible cancellation errors
   Dim AA As Double, BB As Double, nc_dtemp As Double
   Dim n As Double, p As Double, W As Double, s As Double, ps As Double
   Dim Result As Double, term As Double, ptx As Double, ptnc As Double
   AA = A - nc * x * (A + B)
   BB = (x * nc - 1#) - A
   If (BB < 0#) Then
      n = BB - Sqr(BB ^ 2 - 4# * AA)
      n = Int(2# * AA / n)
   Else
      n = Int((BB + Sqr(BB ^ 2 - 4# * AA)) / 2#)
   End If
   If n < 0# Then
      n = 0#
   End If
   AA = n + A
   BB = n
   ptnc = poissonTerm(n, nc, nc - n, 0#)
   ptx = B * binomialTerm(AA, B, x, Y, B * x - AA * Y, 0#)  '  (aa + b)*(I(x, aa, b) - I(x, aa + 1, b))
   AA = AA + 1#
   BB = BB + 1#
   p = nc / BB
   ps = p
   nc_derivative = ps
   s = x / AA  ' (I(x, aa, b) - I(x, aa + 1, b)) / ptx
   W = p
   term = s * W
   Result = term
   If ptx > 0 Then
     While (((term > 0.000000000000001 * Result) And (p > 1E-16 * W)) Or (ps > 1E-16 * nc_derivative))
       s = (AA + B) * s
       AA = AA + 1#
       BB = BB + 1#
       p = nc / BB * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / AA * s ' (I(x, aa, b) - I(x, aa + 1, b)) / ptx
       W = W + p
       term = s * W
       Result = Result + term
     Wend
     W = W * ptnc
   Else
     W = comppoisson(n, nc, nc - n)
   End If
   If x > Y Then
      s = compbeta(Y, B, A + (BB + 1#))
   Else
      s = beta(x, A + (BB + 1#), B)
   End If
   BETA_nc1 = Result * ptx * ptnc + s * W
   ps = 1#
   nc_dtemp = 0#
   AA = n + A
   BB = n
   p = 1#
   s = ptx / (AA + B) ' I(x, aa, b) - I(x, aa + 1, b)
   If x > Y Then
      W = compbeta(Y, B, AA) ' I(x, aa, b)
   Else
      W = beta(x, AA, B) ' I(x, aa, b)
   End If
   term = p * W
   Result = term
   While BB > 0# And (((term > 0.000000000000001 * Result) And (s > 1E-16 * W)) Or (ps > 1E-16 * nc_dtemp))
       s = AA / x * s
       ps = p * s
       nc_dtemp = nc_dtemp + ps
       p = BB / nc * p
       AA = AA - 1#
       BB = BB - 1#
       If BB = 0# Then AA = A
       s = s / (AA + B) ' I(x, aa, b) - I(x, aa + 1, b)
       W = W + s ' I(x, aa, b)
       term = p * W
       Result = Result + term
   Wend
   If n > 0# Then
      nc_dtemp = nc_derivative * ptx + nc_dtemp + p * AA / x * s
   ElseIf B = 0# Then
      nc_dtemp = 0#
   Else
      nc_dtemp = binomialTerm(AA, B, x, Y, B * x - AA * Y, Log(B) + Log((nc_derivative + AA / (x * (AA + B)))))
   End If
   nc_dtemp = nc_dtemp / Y
   BETA_nc1 = BETA_nc1 + Result * ptnc + CPoisson(BB - 1#, nc, nc - BB + 1#) * W
   If nc_dtemp = 0# Then
      nc_derivative = 0#
   Else
      nc_derivative = poissonTerm(n, nc, nc - n, Log(nc_dtemp))
   End If
End Function

Private Function comp_BETA_nc1(ByVal x As Double, ByVal Y As Double, ByVal A As Double, ByVal B As Double, ByVal nc As Double, ByRef nc_derivative As Double) As Double
'y is 1-x but held accurately to avoid possible cancellation errors
   Dim AA As Double, BB As Double, nc_dtemp As Double
   Dim n As Double, p As Double, W As Double, s As Double, ps As Double
   Dim Result As Double, term As Double, ptx As Double, ptnc As Double
   AA = A - nc * x * (A + B)
   BB = (x * nc - 1#) - A
   If (BB < 0#) Then
      n = BB - Sqr(BB ^ 2 - 4# * AA)
      n = Int(2# * AA / n)
   Else
      n = Int((BB + Sqr(BB ^ 2 - 4# * AA)) / 2)
   End If
   If n < 0# Then
      n = 0#
   End If
   AA = n + A
   BB = n
   ptnc = poissonTerm(n, nc, nc - n, 0#)
   ptx = B / (AA + B) * binomialTerm(AA, B, x, Y, B * x - AA * Y, 0#) '(1 - I(x, aa + 1, b)) - (1 - I(x, aa, b))
   ps = 1#
   nc_dtemp = 0#
   p = 1#
   s = 1#
   W = p
   term = 1#
   Result = 0#
   If ptx > 0 Then
     While BB > 0# And (((term > 0.000000000000001 * Result) And (p > 1E-16 * W)) Or (ps > 1E-16 * nc_dtemp))
       s = AA / x * s
       ps = p * s
       nc_dtemp = nc_dtemp + ps
       p = BB / nc * p
       AA = AA - 1#
       BB = BB - 1#
       If BB = 0# Then AA = A
       s = s / (AA + B) ' (1 - I(x, aa + 1, b)) - (1 - I(x, aa + 1, b))
       term = s * W
       Result = Result + term
       W = W + p
     Wend
     W = W * ptnc
   Else
     W = CPoisson(n, nc, nc - n)
   End If
   If n > 0# Then
      nc_dtemp = (nc_dtemp + p * AA / x * s) * ptx
   ElseIf A = 0# Or B = 0# Then
      nc_dtemp = 0#
   Else
      nc_dtemp = binomialTerm(AA, B, x, Y, B * x - AA * Y, Log(B) + Log(AA / (x * (AA + B))))
   End If
   If x > Y Then
      s = beta(Y, B, AA)
   Else
      s = compbeta(x, AA, B)
   End If
   comp_BETA_nc1 = Result * ptx * ptnc + s * W
   AA = n + A
   BB = n
   p = 1#
   nc_derivative = 0#
   s = ptx
   If x > Y Then
      W = beta(Y, B, AA) '  1 - I(x, aa, b)
   Else
      W = compbeta(x, AA, B) ' 1 - I(x, aa, b)
   End If
   term = 0#
   Result = term
   Do
       W = W + s ' 1 - I(x, aa, b)
       s = (AA + B) * s
       AA = AA + 1#
       BB = BB + 1#
       p = nc / BB * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / AA * s ' (1 - I(x, aa + 1, b)) - (1 - I(x, aa, b))
       term = p * W
       Result = Result + term
   Loop While (((term > 0.000000000000001 * Result) And (s > 1E-16 * W)) Or (ps > 1E-16 * nc_derivative))
   nc_dtemp = (nc_derivative + nc_dtemp) / Y
   comp_BETA_nc1 = comp_BETA_nc1 + Result * ptnc + comppoisson(BB, nc, nc - BB) * W
   If nc_dtemp = 0# Then
      nc_derivative = 0#
   Else
      nc_derivative = poissonTerm(n, nc, nc - n, Log(nc_dtemp))
   End If
End Function

Private Function inv_BETA_nc1(ByVal prob As Double, ByVal A As Double, ByVal B As Double, ByVal nc As Double, ByRef oneMinusP As Double) As Double
'Uses approx in A&S 26.6.26 for to get initial estimate the modified NR to improve it.
Dim x As Double, Y As Double, pr As Double, dif As Double, temp As Double
Dim hip As Double, lop As Double
Dim hix As Double, lox As Double, nc_derivative As Double
   If (prob > 0.5) Then
      inv_BETA_nc1 = comp_inv_BETA_nc1(1# - prob, A, B, nc, oneMinusP)
      Exit Function
   End If

   lop = 0#
   hip = 1#
   lox = 0#
   hix = 1#
   pr = Exp(-nc)
   If pr > prob Then
      If 2# * prob > pr Then
         x = invcompbeta(A + cSmall, B, (pr - prob) / pr, oneMinusP)
      Else
         x = invbeta(A + cSmall, B, prob / pr, oneMinusP)
      End If
      If x = 0# Then
         inv_BETA_nc1 = 0#
         Exit Function
      Else
         temp = oneMinusP
         Y = invbeta((A + nc) ^ 2 / (A + 2# * nc), B, prob, oneMinusP)
         oneMinusP = (A + nc) * oneMinusP / (A + nc * (1# + Y))
         If temp > oneMinusP Then
            oneMinusP = temp
         Else
            x = (A + 2# * nc) * Y / (A + nc * (1# + Y))
         End If
      End If
   Else
      Y = invbeta((A + nc) ^ 2 / (A + 2# * nc), B, prob, oneMinusP)
      x = (A + 2# * nc) * Y / (A + nc * (1# + Y))
      oneMinusP = (A + nc) * oneMinusP / (A + nc * (1# + Y))
      If oneMinusP < cSmall Then
         oneMinusP = cSmall
         pr = BETA_nc1(x, oneMinusP, A, B, nc, nc_derivative)
         If pr < prob Then
            inv_BETA_nc1 = 1#
            oneMinusP = 0#
            Exit Function
         End If
      End If
   End If
   Do
      pr = BETA_nc1(x, oneMinusP, A, B, nc, nc_derivative)
      If pr < 3E-308 And nc_derivative = 0# Then
         hip = oneMinusP
         lox = x
         dif = dif / 2#
         x = x - dif
         oneMinusP = oneMinusP + dif
      ElseIf nc_derivative = 0# Then
         lop = oneMinusP
         hix = x
         dif = dif / 2#
         x = x - dif
         oneMinusP = oneMinusP + dif
      Else
         If pr < prob Then
            hip = oneMinusP
            lox = x
         Else
            lop = oneMinusP
            hix = x
         End If
         dif = -(pr / nc_derivative) * logdif(pr, prob)
         If x > oneMinusP Then
            If oneMinusP - dif < lop Then
               dif = (oneMinusP - lop) * 0.9
            ElseIf oneMinusP - dif > hip Then
               dif = (oneMinusP - hip) * 0.9
            End If
         ElseIf x + dif < lox Then
            dif = (lox - x) * 0.9
         ElseIf x + dif > hix Then
            dif = (hix - x) * 0.9
         End If
         x = x + dif
         oneMinusP = oneMinusP - dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(Min(x, oneMinusP)) * 0.0000000001))
   inv_BETA_nc1 = x
End Function

Private Function comp_inv_BETA_nc1(ByVal prob As Double, ByVal A As Double, ByVal B As Double, ByVal nc As Double, ByRef oneMinusP As Double) As Double
'Uses approx in A&S 26.6.26 for to get initial estimate the modified NR to improve it.
Dim x As Double, Y As Double, pr As Double, dif As Double, temp As Double
Dim hip As Double, lop As Double
Dim hix As Double, lox As Double, nc_derivative As Double
   If (prob > 0.5) Then
      comp_inv_BETA_nc1 = inv_BETA_nc1(1# - prob, A, B, nc, oneMinusP)
      Exit Function
   End If

   lop = 0#
   hip = 1#
   lox = 0#
   hix = 1#
   pr = Exp(-nc)
   If pr > prob Then
      If 2# * prob > pr Then
         x = invbeta(A + cSmall, B, (pr - prob) / pr, oneMinusP)
      Else
         x = invcompbeta(A + cSmall, B, prob / pr, oneMinusP)
      End If
      If oneMinusP < cSmall Then
         oneMinusP = cSmall
         pr = comp_BETA_nc1(x, oneMinusP, A, B, nc, nc_derivative)
         If pr > prob Then
            comp_inv_BETA_nc1 = 1#
            oneMinusP = 0#
            Exit Function
         End If
      Else
         temp = oneMinusP
         Y = invcompbeta((A + nc) ^ 2 / (A + 2# * nc), B, prob, oneMinusP)
         oneMinusP = (A + nc) * oneMinusP / (A + nc * (1# + Y))
         If temp < oneMinusP Then
            oneMinusP = temp
         Else
            x = (A + 2# * nc) * Y / (A + nc * (1# + Y))
         End If
         If oneMinusP < cSmall Then
            oneMinusP = cSmall
            pr = comp_BETA_nc1(x, oneMinusP, A, B, nc, nc_derivative)
            If pr > prob Then
               comp_inv_BETA_nc1 = 1#
               oneMinusP = 0#
               Exit Function
            End If
         ElseIf x < cSmall Then
            x = cSmall
            pr = comp_BETA_nc1(x, oneMinusP, A, B, nc, nc_derivative)
            If pr < prob Then
               comp_inv_BETA_nc1 = 0#
               oneMinusP = 1#
               Exit Function
            End If
         End If
      End If
   Else
      Y = invcompbeta((A + nc) ^ 2 / (A + 2# * nc), B, prob, oneMinusP)
      x = (A + 2# * nc) * Y / (A + nc * (1# + Y))
      oneMinusP = (A + nc) * oneMinusP / (A + nc * (1# + Y))
      If oneMinusP < cSmall Then
         oneMinusP = cSmall
         pr = comp_BETA_nc1(x, oneMinusP, A, B, nc, nc_derivative)
         If pr > prob Then
            comp_inv_BETA_nc1 = 1#
            oneMinusP = 0#
            Exit Function
         End If
      ElseIf x < cSmall Then
         x = cSmall
         pr = comp_BETA_nc1(x, oneMinusP, A, B, nc, nc_derivative)
         If pr < prob Then
            comp_inv_BETA_nc1 = 0#
            oneMinusP = 1#
            Exit Function
         End If
      End If
   End If
   dif = x
   Do
      pr = comp_BETA_nc1(x, oneMinusP, A, B, nc, nc_derivative)
      If pr < 3E-308 And nc_derivative = 0# Then
         lop = oneMinusP
         hix = x
         dif = dif / 2#
         x = x - dif
         oneMinusP = oneMinusP + dif
      ElseIf nc_derivative = 0# Then
         hip = oneMinusP
         lox = x
         dif = dif / 2#
         x = x - dif
         oneMinusP = oneMinusP + dif
      Else
         If pr < prob Then
            lop = oneMinusP
            hix = x
         Else
            hip = oneMinusP
            lox = x
         End If
         dif = (pr / nc_derivative) * logdif(pr, prob)
         If x > oneMinusP Then
            If oneMinusP - dif < lop Then
               dif = (oneMinusP - lop) * 0.9
            ElseIf oneMinusP - dif > hip Then
               dif = (oneMinusP - hip) * 0.9
            End If
         ElseIf x + dif < lox Then
            dif = (lox - x) * 0.9
         ElseIf x + dif > hix Then
            dif = (hix - x) * 0.9
         End If
         x = x + dif
         oneMinusP = oneMinusP - dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(Min(x, oneMinusP)) * 0.0000000001))
   comp_inv_BETA_nc1 = x
End Function

Private Function invBetaLessThanX(ByVal prob As Double, ByVal x As Double, ByVal Y As Double, ByVal A As Double, ByVal B As Double) As Double
   Dim oneMinusP As Double
   If x >= Y Then
      If invcompbeta(B, A, prob, oneMinusP) >= Y * (1# - 0.000000000000001) Then
         invBetaLessThanX = 0#
      Else
         invBetaLessThanX = [#VALUE!]
      End If
   ElseIf invbeta(A, B, prob, oneMinusP) <= x * (1# + 0.000000000000001) Then
      invBetaLessThanX = 0#
   Else
      invBetaLessThanX = [#VALUE!]
   End If
End Function

Private Function compInvBetaLessThanX(ByVal prob As Double, ByVal x As Double, ByVal Y As Double, ByVal A As Double, ByVal B As Double) As Double
   Dim oneMinusP As Double
   If x >= Y Then
      If invbeta(B, A, prob, oneMinusP) >= Y * (1# - 0.000000000000001) Then
         compInvBetaLessThanX = 0#
      Else
         compInvBetaLessThanX = [#VALUE!]
      End If
   ElseIf invcompbeta(A, B, prob, oneMinusP) <= x * (1# + 0.000000000000001) Then
      compInvBetaLessThanX = 0#
   Else
      compInvBetaLessThanX = [#VALUE!]
   End If
End Function

Private Function ncp_BETA_nc1(ByVal prob As Double, ByVal x As Double, ByVal Y As Double, ByVal A As Double, ByVal B As Double) As Double
'Uses Normal approx for difference of 2 a Negative Binomial and a poisson distributed variable to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, deriv As Double, c As Double, D As Double, E As Double, sqarg As Double, checked_nc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double, nc_derivative As Double
   If (prob > 0.5) Then
      ncp_BETA_nc1 = comp_ncp_BETA_nc1(1# - prob, x, Y, A, B)
      Exit Function
   End If

   lo = 0#
   hi = nc_limit
   checked_0_limit = False
   checked_nc_limit = False
   temp = inv_normal(prob) ^ 2
   c = B * x / Y
   D = temp - 2# * (A - c)
   If D < 2 * nc_limit Then
      E = (c - A) ^ 2 - temp * c / Y
      sqarg = D ^ 2 - 4 * E
      If sqarg < 0 Then
         ncp = D / 2
      Else
         ncp = (D + Sqr(sqarg)) / 2
      End If
   Else
      ncp = nc_limit
   End If
   ncp = Min(max(0#, ncp), nc_limit)
   If x > Y Then
      pr = compbeta(Y * (1 + ncp / (ncp + A)) / (1 + ncp / (ncp + A) * Y), B, A + ncp ^ 2 / (2 * ncp + A))
   Else
      pr = beta(x / (1 + ncp / (ncp + A) * Y), A + ncp ^ 2 / (2 * ncp + A), B)
   End If
   'Debug.Print "ncp_BETA_nc1 ncp1 ", ncp, pr
   If ncp = 0# Then
      If pr < prob Then
         ncp_BETA_nc1 = invBetaLessThanX(prob, x, Y, A, B)
         Exit Function
      Else
         checked_0_limit = True
      End If
   End If
   temp = Min(max(0#, invcompgamma(B * x, prob) / Y - A), nc_limit)
   If temp = ncp Then
      c = pr
   ElseIf x > Y Then
      c = compbeta(Y * (1 + temp / (temp + A)) / (1 + temp / (temp + A) * Y), B, A + temp ^ 2 / (2 * temp + A))
   Else
      c = beta(x / (1 + temp / (temp + A) * Y), A + temp ^ 2 / (2 * temp + A), B)
   End If
   'Debug.Print "ncp_BETA_nc1 ncp2 ", temp, c
   If temp = 0# Then
      If c < prob Then
         ncp_BETA_nc1 = invBetaLessThanX(prob, x, Y, A, B)
         Exit Function
      Else
         checked_0_limit = True
      End If
   End If
   If pr * c = 0# Then
      ncp = Min(ncp, temp)
      pr = max(pr, c)
      If pr = 0# Then
         c = compbeta(Y, B, A)
         If c < prob Then
            ncp_BETA_nc1 = invBetaLessThanX(prob, x, Y, A, B)
            Exit Function
         Else
            checked_0_limit = True
         End If
      End If
   ElseIf Abs(Log(pr / prob)) > Abs(Log(c / prob)) Then
      ncp = temp
      pr = c
   End If
   If ncp = 0# Then
      If B > 1 + 0.000001 Then
         deriv = BETA_nc1(x, Y, A + 1#, B - 1#, ncp, nc_derivative)
         deriv = nc_derivative * Y ^ 2 / (B - 1#)
      Else
         deriv = pr - BETA_nc1(x, Y, A + 1#, B, ncp, nc_derivative)
      End If
      If deriv = 0# Then
         ncp = nc_limit
      Else
         ncp = (pr - prob) / deriv
         If ncp >= nc_limit Then
            ncp = (pr / deriv) * logdif(pr, prob)
         End If
      End If
   Else
      If ncp = nc_limit Then
         If pr > prob Then
            ncp_BETA_nc1 = [#VALUE!]
            Exit Function
         Else
            checked_nc_limit = True
         End If
      End If
      If pr > 0 Then
         temp = ncp * 0.999999 'Use numerical derivative on approximation since cheap compared to evaluating non-central BETA
         If x > Y Then
            c = compbeta(Y * (1# + temp / (temp + A)) / (1 + temp / (temp + A) * Y), B, A + temp ^ 2 / (2 * temp + A))
         Else
            c = beta(x / (1 + temp / (temp + A) * Y), A + temp ^ 2 / (2 * temp + A), B)
         End If
         If pr <> c Then
            dif = (0.000001 * ncp * pr / (pr - c)) * logdif(pr, prob)
            If ncp - dif < 0# Then
               ncp = ncp / 2#
            ElseIf ncp - dif > nc_limit Then
               ncp = (ncp + nc_limit) / 2#
            Else
               ncp = ncp - dif
            End If
         End If
      Else
         ncp = ncp / 2#
      End If
   End If
   dif = ncp
   Do
      pr = BETA_nc1(x, Y, A, B, ncp, nc_derivative)
      'Debug.Print ncp, pr, prob
      If B > 1 + 0.000001 Then
         deriv = BETA_nc1(x, Y, A + 1#, B - 1#, ncp, nc_derivative)
         deriv = nc_derivative * Y ^ 2 / (B - 1#)
      Else
         deriv = pr - BETA_nc1(x, Y, A + 1#, B, ncp, nc_derivative)
      End If
      If pr < 3E-308 And deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      ElseIf deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr < prob Then
            hi = ncp
         Else
            lo = ncp
         End If
         dif = (pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = cdf_BETA_nc(x, A, B, lo)
               If temp < prob Then
                  ncp_BETA_nc1 = invBetaLessThanX(prob, x, Y, A, B)
                  Exit Function
               Else
                  checked_0_limit = True
               End If
            End If
         ElseIf ncp + dif > hi Then
            dif = (hi - ncp) / 2#
            If Not checked_nc_limit And (hi = nc_limit) Then
               temp = cdf_BETA_nc(x, A, B, hi)
               If temp > prob Then
                  ncp_BETA_nc1 = [#VALUE!]
                  Exit Function
               Else
                  checked_nc_limit = True
               End If
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   ncp_BETA_nc1 = ncp
   'Debug.Print "ncp_BETA_nc1", ncp_BETA_nc1
End Function

Private Function comp_ncp_BETA_nc1(ByVal prob As Double, ByVal x As Double, ByVal Y As Double, ByVal A As Double, ByVal B As Double) As Double
'Uses Normal approx for difference of 2 a Negative Binomial and a poisson distributed variable to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, deriv As Double, c As Double, D As Double, E As Double, sqarg As Double, checked_nc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double, nc_derivative As Double
   If (prob > 0.5) Then
      comp_ncp_BETA_nc1 = ncp_BETA_nc1(1# - prob, x, Y, A, B)
      Exit Function
   End If

   lo = 0#
   hi = nc_limit
   checked_0_limit = False
   checked_nc_limit = False
   temp = inv_normal(prob) ^ 2
   c = B * x / Y
   D = temp - 2# * (A - c)
   If D < 4 * nc_limit Then
      sqarg = D ^ 2 - 4 * E
      If sqarg < 0 Then
         ncp = D / 2
      Else
         ncp = (D - Sqr(sqarg)) / 2
      End If
   Else
      ncp = 0#
   End If
   ncp = Min(max(0#, ncp), nc_limit)
   If x > Y Then
      pr = beta(Y * (1 + ncp / (ncp + A)) / (1 + ncp / (ncp + A) * Y), B, A + ncp ^ 2 / (2 * ncp + A))
   Else
      pr = compbeta(x / (1 + ncp / (ncp + A) * Y), A + ncp ^ 2 / (2 * ncp + A), B)
   End If
   'Debug.Print "comp_ncp_BETA_nc1 ncp1 ", ncp, pr
   If ncp = 0# Then
      If pr > prob Then
         comp_ncp_BETA_nc1 = compInvBetaLessThanX(prob, x, Y, A, B)
         Exit Function
      Else
         checked_0_limit = True
      End If
   End If
   temp = Min(max(0#, invgamma(B * x, prob) / Y - A), nc_limit)
   If temp = ncp Then
      c = pr
   ElseIf x > Y Then
      c = beta(Y * (1 + temp / (temp + A)) / (1 + temp / (temp + A) * Y), B, A + temp ^ 2 / (2 * temp + A))
   Else
      c = compbeta(x / (1 + temp / (temp + A) * Y), A + temp ^ 2 / (2 * temp + A), B)
   End If
   'Debug.Print "comp_ncp_BETA_nc1 ncp2 ", temp, c
   If temp = 0# Then
      If c > prob Then
         comp_ncp_BETA_nc1 = compInvBetaLessThanX(prob, x, Y, A, B)
         Exit Function
      Else
         checked_0_limit = True
      End If
   End If
   If pr * c = 0# Then
      ncp = max(ncp, temp)
      pr = max(pr, c)
   ElseIf Abs(Log(pr / prob)) > Abs(Log(c / prob)) Then
      ncp = temp
      pr = c
   End If
   If ncp = 0# Then
      If pr > prob Then
         comp_ncp_BETA_nc1 = compInvBetaLessThanX(prob, x, Y, A, B)
         Exit Function
      Else
         If B > 1 + 0.000001 Then
            deriv = BETA_nc1(x, Y, A + 1#, B - 1#, 0#, nc_derivative)
            deriv = nc_derivative * Y ^ 2 / (B - 1#)
         Else
            deriv = comp_BETA_nc1(x, Y, A + 1#, B, 0#, nc_derivative) - pr
         End If
         If deriv = 0# Then
            ncp = nc_limit
         Else
            ncp = (prob - pr) / deriv
            If ncp >= nc_limit Then
               ncp = -(pr / deriv) * logdif(pr, prob)
            End If
         End If
         checked_0_limit = True
      End If
   Else
      If ncp = nc_limit Then
         If pr < prob Then
            comp_ncp_BETA_nc1 = [#VALUE!]
            Exit Function
         Else
            checked_nc_limit = True
         End If
      End If
      If pr > 0 Then
         temp = ncp * 0.999999 'Use numerical derivative on approximation since cheap compared to evaluating non-central BETA
         If x > Y Then
            c = beta(Y * (1# + temp / (temp + A)) / (1 + temp / (temp + A) * Y), B, A + temp ^ 2 / (2 * temp + A))
         Else
            c = compbeta(x / (1 + temp / (temp + A) * Y), A + temp ^ 2 / (2 * temp + A), B)
         End If
         If pr <> c Then
            dif = -(0.000001 * ncp * pr / (pr - c)) * logdif(pr, prob)
            If ncp + dif < 0 Then
               ncp = ncp / 2
            ElseIf ncp + dif > nc_limit Then
               ncp = (ncp + nc_limit) / 2
            Else
               ncp = ncp + dif
            End If
         End If
      Else
         ncp = (nc_limit + ncp) / 2#
      End If
   End If
   dif = ncp
   Do
      pr = comp_BETA_nc1(x, Y, A, B, ncp, nc_derivative)
      'Debug.Print ncp, pr, prob
      If B > 1 + 0.000001 Then
         deriv = BETA_nc1(x, Y, A + 1#, B - 1#, ncp, nc_derivative)
         deriv = nc_derivative * Y ^ 2 / (B - 1#)
      Else
         deriv = comp_BETA_nc1(x, Y, A + 1#, B, ncp, nc_derivative) - pr
      End If
      If pr < 3E-308 And deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp + dif
      ElseIf deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr < prob Then
            lo = ncp
         Else
            hi = ncp
         End If
         dif = -(pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = comp_cdf_BETA_nc(x, A, B, lo)
               If temp > prob Then
                  comp_ncp_BETA_nc1 = compInvBetaLessThanX(prob, x, Y, A, B)
                  Exit Function
               Else
                  checked_0_limit = True
               End If
            End If
         ElseIf ncp + dif > hi Then
            dif = (hi - ncp) / 2#
            If Not checked_nc_limit And (hi = nc_limit) Then
               temp = comp_cdf_BETA_nc(x, A, B, hi)
               If temp < prob Then
                  comp_ncp_BETA_nc1 = [#VALUE!]
                  Exit Function
               Else
                  checked_nc_limit = True
               End If
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   comp_ncp_BETA_nc1 = ncp
   'Debug.Print "comp_ncp_BETA_nc1", comp_ncp_BETA_nc1
End Function

Public Function pdf_BETA_nc(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, ByVal nc_param As Double) As Double
  If (shape_param1 < 0#) Or (shape_param2 < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Or ((shape_param1 = 0#) And (shape_param2 = 0#)) Then
     pdf_BETA_nc = [#VALUE!]
  ElseIf (x < 0# Or x > 1#) Then
     pdf_BETA_nc = 0#
  ElseIf (x = 0# Or nc_param = 0#) Then
     pdf_BETA_nc = Exp(-nc_param) * pdf_BETA(x, shape_param1, shape_param2)
  ElseIf (x = 1# And shape_param2 = 1#) Then
     pdf_BETA_nc = shape_param1 + nc_param
  ElseIf (x = 1#) Then
     pdf_BETA_nc = pdf_BETA(x, shape_param1, shape_param2)
  Else
     Dim nc_derivative As Double
     If (shape_param1 < 1# Or x * shape_param2 <= (1# - x) * (shape_param1 + nc_param)) Then
        pdf_BETA_nc = BETA_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
     Else
        pdf_BETA_nc = comp_BETA_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
     End If
     pdf_BETA_nc = nc_derivative
  End If
  pdf_BETA_nc = GetRidOfMinusZeroes(pdf_BETA_nc)
End Function

Public Function cdf_BETA_nc(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, ByVal nc_param As Double) As Double
  Dim nc_derivative As Double
  If (shape_param1 < 0#) Or (shape_param2 < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Or ((shape_param1 = 0#) And (shape_param2 = 0#)) Then
     cdf_BETA_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     cdf_BETA_nc = 0#
  ElseIf (x >= 1#) Then
     cdf_BETA_nc = 1#
  ElseIf (x = 0# And shape_param1 = 0#) Then
     cdf_BETA_nc = Exp(-nc_param)
  ElseIf (x = 0#) Then
     cdf_BETA_nc = 0#
  ElseIf (nc_param = 0#) Then
     cdf_BETA_nc = beta(x, shape_param1, shape_param2)
  ElseIf (shape_param1 < 1# Or x * shape_param2 <= (1# - x) * (shape_param1 + nc_param)) Then
     cdf_BETA_nc = BETA_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
  Else
     cdf_BETA_nc = 1# - comp_BETA_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
  End If
  cdf_BETA_nc = GetRidOfMinusZeroes(cdf_BETA_nc)
End Function

Public Function comp_cdf_BETA_nc(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, ByVal nc_param As Double) As Double
  Dim nc_derivative As Double
  If (shape_param1 < 0#) Or (shape_param2 < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Or ((shape_param1 = 0#) And (shape_param2 = 0#)) Then
     comp_cdf_BETA_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     comp_cdf_BETA_nc = 1#
  ElseIf (x >= 1#) Then
     comp_cdf_BETA_nc = 0#
  ElseIf (x = 0# And shape_param1 = 0#) Then
     comp_cdf_BETA_nc = -expm1(-nc_param)
  ElseIf (x = 0#) Then
     comp_cdf_BETA_nc = 1#
  ElseIf (nc_param = 0#) Then
     comp_cdf_BETA_nc = compbeta(x, shape_param1, shape_param2)
  ElseIf (shape_param1 < 1# Or x * shape_param2 >= (1# - x) * (shape_param1 + nc_param)) Then
     comp_cdf_BETA_nc = comp_BETA_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
  Else
     comp_cdf_BETA_nc = 1# - BETA_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
  End If
  comp_cdf_BETA_nc = GetRidOfMinusZeroes(comp_cdf_BETA_nc)
End Function

Public Function inv_BETA_nc(ByVal prob As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, ByVal nc_param As Double) As Double
  Dim oneMinusP As Double
  If (shape_param1 < 0#) Or (shape_param2 <= 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Or (prob < 0#) Or (prob > 1#) Then
     inv_BETA_nc = [#VALUE!]
  ElseIf (prob = 0# Or shape_param1 = 0# And prob <= Exp(-nc_param)) Then
     inv_BETA_nc = 0#
  ElseIf (prob = 1#) Then
     inv_BETA_nc = 1#
  ElseIf (nc_param = 0#) Then
     inv_BETA_nc = invbeta(shape_param1, shape_param2, prob, oneMinusP)
  Else
     inv_BETA_nc = inv_BETA_nc1(prob, shape_param1, shape_param2, nc_param, oneMinusP)
  End If
  inv_BETA_nc = GetRidOfMinusZeroes(inv_BETA_nc)
End Function

Public Function comp_inv_BETA_nc(ByVal prob As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, ByVal nc_param As Double) As Double
  Dim oneMinusP As Double
  If (shape_param1 < 0#) Or (shape_param2 <= 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Or (prob < 0#) Or (prob > 1#) Then
     comp_inv_BETA_nc = [#VALUE!]
  ElseIf (prob = 1# Or shape_param1 = 0# And prob >= -expm1(-nc_param)) Then
     comp_inv_BETA_nc = 0#
  ElseIf (prob = 0#) Then
     comp_inv_BETA_nc = 1#
  ElseIf (nc_param = 0#) Then
     comp_inv_BETA_nc = invcompbeta(shape_param1, shape_param2, prob, oneMinusP)
  Else
     comp_inv_BETA_nc = comp_inv_BETA_nc1(prob, shape_param1, shape_param2, nc_param, oneMinusP)
  End If
  comp_inv_BETA_nc = GetRidOfMinusZeroes(comp_inv_BETA_nc)
End Function

Public Function ncp_BETA_nc(ByVal prob As Double, ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
  If (shape_param1 < 0#) Or (shape_param2 <= 0#) Or (x < 0#) Or (x >= 1#) Or (prob <= 0#) Or (prob > 1#) Then
     ncp_BETA_nc = [#VALUE!]
  ElseIf (x = 0# And shape_param1 = 0#) Then
     ncp_BETA_nc = -Log(prob)
  ElseIf (x = 0# Or prob = 1#) Then
     ncp_BETA_nc = [#VALUE!]
  Else
     ncp_BETA_nc = ncp_BETA_nc1(prob, x, 1# - x, shape_param1, shape_param2)
  End If
  ncp_BETA_nc = GetRidOfMinusZeroes(ncp_BETA_nc)
End Function

Public Function comp_ncp_BETA_nc(ByVal prob As Double, ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
  If (shape_param1 < 0#) Or (shape_param2 <= 0#) Or (x < 0#) Or (x >= 1#) Or (prob < 0#) Or (prob >= 1#) Then
     comp_ncp_BETA_nc = [#VALUE!]
  ElseIf (x = 0# And shape_param1 = 0#) Then
     comp_ncp_BETA_nc = -log0(-prob)
  ElseIf (x = 0# Or prob = 0#) Then
     comp_ncp_BETA_nc = [#VALUE!]
  Else
     comp_ncp_BETA_nc = comp_ncp_BETA_nc1(prob, x, 1# - x, shape_param1, shape_param2)
  End If
  comp_ncp_BETA_nc = GetRidOfMinusZeroes(comp_ncp_BETA_nc)
End Function

Public Function pdf_fdist_nc(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, ByVal nc As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or (nc < 0#) Or (nc > 2# * nc_limit)) Then
      pdf_fdist_nc = [#VALUE!]
   ElseIf (x < 0#) Then
      pdf_fdist_nc = 0#
   ElseIf (x = 0# Or nc = 0#) Then
      pdf_fdist_nc = Exp(-nc / 2#) * pdf_fdist(x, df1, df2)
   Else
      Dim p As Double, q As Double, nc_derivative As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      If (df1 < 1# Or p * df2 <= q * (df1 + nc)) Then
         pdf_fdist_nc = BETA_nc1(p, q, df1 / 2#, df2 / 2#, nc / 2#, nc_derivative)
      Else
         pdf_fdist_nc = comp_BETA_nc1(p, q, df1 / 2#, df2 / 2#, nc / 2#, nc_derivative)
      End If
      pdf_fdist_nc = (nc_derivative * q) * (df1 * q / df2)
   End If
   pdf_fdist_nc = GetRidOfMinusZeroes(pdf_fdist_nc)
End Function

Public Function cdf_fdist_nc(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, ByVal nc As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or (nc < 0#) Or (nc > 2# * nc_limit)) Then
      cdf_fdist_nc = [#VALUE!]
   ElseIf (x <= 0#) Then
      cdf_fdist_nc = 0#
   Else
      Dim p As Double, q As Double, nc_derivative As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      'If p < cSmall And x <> 0# Or q < cSmall Then
      '   cdf_fdist_nc = [#VALUE!]
      '   Exit Function
      'End If
      df2 = df2 / 2#
      df1 = df1 / 2#
      nc = nc / 2#
      If (nc = 0# And p <= q) Then
         cdf_fdist_nc = beta(p, df1, df2)
      ElseIf (nc = 0#) Then
         cdf_fdist_nc = compbeta(q, df2, df1)
      ElseIf (df1 < 1# Or p * df2 <= q * (df1 + nc)) Then
         cdf_fdist_nc = BETA_nc1(p, q, df1, df2, nc, nc_derivative)
      Else
         cdf_fdist_nc = 1# - comp_BETA_nc1(p, q, df1, df2, nc, nc_derivative)
      End If
   End If
   cdf_fdist_nc = GetRidOfMinusZeroes(cdf_fdist_nc)
End Function

Public Function comp_cdf_fdist_nc(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, ByVal nc As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or (nc < 0#) Or (nc > 2# * nc_limit)) Then
      comp_cdf_fdist_nc = [#VALUE!]
   ElseIf (x <= 0#) Then
      comp_cdf_fdist_nc = 1#
   Else
      Dim p As Double, q As Double, nc_derivative As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      'If p < cSmall And x <> 0# Or q < cSmall Then
      '   comp_cdf_fdist_nc = [#VALUE!]
      '   Exit Function
      'End If
      df2 = df2 / 2#
      df1 = df1 / 2#
      nc = nc / 2#
      If (nc = 0# And p <= q) Then
         comp_cdf_fdist_nc = compbeta(p, df1, df2)
      ElseIf (nc = 0#) Then
         comp_cdf_fdist_nc = beta(q, df2, df1)
      ElseIf (df1 < 1# Or p * df2 >= q * (df1 + nc)) Then
         comp_cdf_fdist_nc = comp_BETA_nc1(p, q, df1, df2, nc, nc_derivative)
      Else
         comp_cdf_fdist_nc = 1# - BETA_nc1(p, q, df1, df2, nc, nc_derivative)
      End If
   End If
   comp_cdf_fdist_nc = GetRidOfMinusZeroes(comp_cdf_fdist_nc)
End Function

Public Function inv_fdist_nc(ByVal prob As Double, ByVal df1 As Double, ByVal df2 As Double, ByVal nc As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or (nc < 0#) Or (nc > 2# * nc_limit) Or prob < 0# Or prob >= 1#) Then
      inv_fdist_nc = [#VALUE!]
   ElseIf (prob = 0#) Then
      inv_fdist_nc = 0#
   Else
      Dim temp As Double, oneMinusP As Double
      df1 = df1 / 2#
      df2 = df2 / 2#
      If nc = 0# Then
         temp = invbeta(df1, df2, prob, oneMinusP)
      Else
         temp = inv_BETA_nc1(prob, df1, df2, nc / 2#, oneMinusP)
      End If
      inv_fdist_nc = df2 * temp / (df1 * oneMinusP)
      'If oneMinusP < cSmall Then inv_fdist_nc = [#VALUE!]
   End If
   inv_fdist_nc = GetRidOfMinusZeroes(inv_fdist_nc)
End Function

Public Function comp_inv_fdist_nc(ByVal prob As Double, ByVal df1 As Double, ByVal df2 As Double, ByVal nc As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or (nc < 0#) Or (nc > 2# * nc_limit) Or prob <= 0# Or prob > 1#) Then
      comp_inv_fdist_nc = [#VALUE!]
   ElseIf (prob = 1#) Then
      comp_inv_fdist_nc = 0#
   Else
      Dim temp As Double, oneMinusP As Double
      df1 = df1 / 2#
      df2 = df2 / 2#
      If nc = 0# Then
         temp = invcompbeta(df1, df2, prob, oneMinusP)
      Else
         temp = comp_inv_BETA_nc1(prob, df1, df2, nc / 2#, oneMinusP)
      End If
      comp_inv_fdist_nc = df2 * temp / (df1 * oneMinusP)
      'If oneMinusP < cSmall Then comp_inv_fdist_nc = [#VALUE!]
   End If
   comp_inv_fdist_nc = GetRidOfMinusZeroes(comp_inv_fdist_nc)
End Function

Public Function ncp_fdist_nc(ByVal prob As Double, ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
  df1 = AlterForIntegralChecks_df(df1)
  df2 = AlterForIntegralChecks_df(df2)
  If (df1 <= 0#) Or (df2 <= 0#) Or (x <= 0#) Or (prob <= 0#) Or (prob >= 1#) Then
     ncp_fdist_nc = [#VALUE!]
  Else
     Dim p As Double, q As Double
     If x > 1# Then
        q = df2 / x
        p = q + df1
        q = q / p
        p = df1 / p
     Else
        p = df1 * x
        q = df2 + p
        p = p / q
        q = df2 / q
     End If
     df2 = df2 / 2#
     df1 = df1 / 2#
     ncp_fdist_nc = ncp_BETA_nc1(prob, p, q, df1, df2) * 2#
  End If
  ncp_fdist_nc = GetRidOfMinusZeroes(ncp_fdist_nc)
End Function

Public Function comp_ncp_fdist_nc(ByVal prob As Double, ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
  df1 = AlterForIntegralChecks_df(df1)
  df2 = AlterForIntegralChecks_df(df2)
  If (df1 <= 0#) Or (df2 <= 0#) Or (x <= 0#) Or (prob <= 0#) Or (prob >= 1#) Then
     comp_ncp_fdist_nc = [#VALUE!]
  Else
     Dim p As Double, q As Double
     If x > 1# Then
        q = df2 / x
        p = q + df1
        q = q / p
        p = df1 / p
     Else
        p = df1 * x
        q = df2 + p
        p = p / q
        q = df2 / q
     End If
     df1 = df1 / 2#
     df2 = df2 / 2#
     comp_ncp_fdist_nc = comp_ncp_BETA_nc1(prob, p, q, df1, df2) * 2#
  End If
  comp_ncp_fdist_nc = GetRidOfMinusZeroes(comp_ncp_fdist_nc)
End Function

Private Function t_nc1(ByVal t As Double, ByVal df As Double, ByVal nct As Double, ByRef nc_derivative As Double) As Double
'y is 1-x but held accurately to avoid possible cancellation errors
'nc_derivative holds t * derivative
   Dim AA As Double, BB As Double, nc_dtemp As Double
   Dim n As Double, p As Double, q As Double, W As Double, V As Double, r As Double, s As Double, ps As Double
   Dim result1 As Double, result2 As Double, term1 As Double, term2 As Double, ptnc As Double, qtnc As Double, ptx As Double, qtx As Double
   Dim A As Double, B As Double, x As Double, Y As Double, nc As Double
   Dim save_result1 As Double, save_result2 As Double, phi As Double, vScale As Double
   phi = CNormal(-Abs(nct))
   A = 0.5
   B = df / 2#
   If Abs(t) >= Min(1#, df) Then
      Y = df / t
      x = t + Y
      Y = Y / x
      x = t / x
   Else
      x = t * t
      Y = df + x
      x = x / Y
      Y = df / Y
   End If
   If Y < cSmall Then
      t_nc1 = [#VALUE!]
      Exit Function
   End If
   nc = nct * nct / 2#
   AA = A - nc * x * (A + B)
   BB = (x * nc - 1#) - A
   If (BB < 0#) Then
      n = BB - Sqr(BB ^ 2 - 4# * AA)
      n = Int(2# * AA / n)
   Else
      n = Int((BB + Sqr(BB ^ 2 - 4# * AA)) / 2#)
   End If
   If n < 0# Then
      n = 0#
   End If
   AA = n + A
   BB = n + 0.5
   qtnc = poissonTerm(BB, nc, nc - BB, 0#)
   BB = n
   ptnc = poissonTerm(BB, nc, nc - BB, 0#)
   ptx = binomialTerm(AA, B, x, Y, B * x - AA * Y, 0#) / (AA + B) '(I(x, aa, b) - I(x, aa+1, b))/b
   qtx = binomialTerm(AA + 0.5, B, x, Y, B * x - (AA + 0.5) * Y, 0#) / (AA + B + 0.5) '(I(x, aa+1/2, b) - I(x, aa+3/2, b))/b
   If B > 1# Then
      ptx = B * ptx
      qtx = B * qtx
   End If
   vScale = max(ptx, qtx)
   If ptx = vScale Then
      s = 1#
   Else
      s = ptx / vScale
   End If
   If qtx = vScale Then
      r = 1#
   Else
      r = qtx / vScale
   End If
   s = (AA + B) * s
   r = (AA + B + 0.5) * r
   AA = AA + 1#
   BB = BB + 1#
   p = nc / BB * ptnc
   q = nc / (BB + 0.5) * qtnc
   ps = p * s + q * r
   nc_derivative = ps
   s = x / AA * s  ' I(x, aa, b) - I(x, aa+1, b)
   r = x / (AA + 0.5) * r ' I(x, aa+1/2, b) - I(x, aa+3/2, b)
   W = p
   V = q
   term1 = s * W
   term2 = r * V
   result1 = term1
   result2 = term2
   While ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) And (p > 1E-16 * W)) Or (ps > 1E-16 * nc_derivative))
       s = (AA + B) * s
       r = (AA + B + 0.5) * r
       AA = AA + 1#
       BB = BB + 1#
       p = nc / BB * p
       q = nc / (BB + 0.5) * q
       ps = p * s + q * r
       nc_derivative = nc_derivative + ps
       s = x / AA * s ' I(x, aa, b) - I(x, aa+1, b)
       r = x / (AA + 0.5) * r ' I(x, aa+1/2, b) - I(x, aa+3/2, b)
       W = W + p
       V = V + q
       term1 = s * W
       term2 = r * V
       result1 = result1 + term1
       result2 = result2 + term2
   Wend
   If x > Y Then
      s = compbeta(Y, B, A + (BB + 1#))
      r = compbeta(Y, B, A + (BB + 1.5))
   Else
      s = beta(x, A + (BB + 1#), B)
      r = beta(x, A + (BB + 1.5), B)
   End If
   nc_derivative = x * nc_derivative * vScale
   If B <= 1# Then vScale = vScale * B
   save_result1 = result1 * vScale + s * W
   save_result2 = result2 * vScale + r * V

   ps = 1#
   nc_dtemp = 0#
   AA = n + A
   BB = n
   vScale = max(ptnc, qtnc)
   If ptnc = vScale Then
      p = 1#
   Else
      p = ptnc / vScale
   End If
   If qtnc = vScale Then
      q = 1#
   Else
      q = qtnc / vScale
   End If
   s = ptx ' I(x, aa, b) - I(x, aa+1, b)
   r = qtx ' I(x, aa+1/2, b) - I(x, aa+3/2, b)
   If x > Y Then
      W = compbeta(Y, B, AA) ' I(x, aa, b)
      V = compbeta(Y, B, AA + 0.5) ' I(x, aa+1/2, b)
   Else
      W = beta(x, AA, B) ' I(x, aa, b)
      V = beta(x, AA + 0.5, B) ' I(x, aa+1/2, b)
   End If
   term1 = p * W
   term2 = q * V
   result1 = term1
   result2 = term2
   While BB > 0# And ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) And (s > 1E-16 * W)) Or (ps > 1E-16 * nc_dtemp))
       s = AA / x * s
       r = (AA + 0.5) / x * r
       ps = p * s + q * r
       nc_dtemp = nc_dtemp + ps
       p = BB / nc * p
       q = (BB + 0.5) / nc * q
       AA = AA - 1#
       BB = BB - 1#
       If BB = 0# Then AA = A
       s = s / (AA + B) ' I(x, aa, b) - I(x, aa+1, b)
       r = r / (AA + B + 0.5) ' I(x, aa+1/2, b) - I(x, aa+3/2, b)
       If B > 1# Then
          W = W + s ' I(x, aa, b)
          V = V + r ' I(x, aa+0.5, b)
       Else
          W = W + B * s
          V = V + B * r
       End If
       term1 = p * W
       term2 = q * V
       result1 = result1 + term1
       result2 = result2 + term2
   Wend
   nc_dtemp = x * nc_dtemp + p * AA * s + q * (AA + 0.5) * r
   p = CPoisson(BB - 1#, nc, nc - BB + 1#)
   q = CPoisson(BB - 0.5, nc, nc - BB + 0.5) - 2# * phi
   result1 = save_result1 + result1 * vScale + p * W
   result2 = save_result2 + result2 * vScale + q * V
   If t > 0# Then
      t_nc1 = phi + 0.5 * (result1 + result2)
      nc_derivative = nc_derivative + nc_dtemp * vScale
   Else
      t_nc1 = phi - 0.5 * (result1 - result2)
   End If
End Function

Private Function comp_t_nc1(ByVal t As Double, ByVal df As Double, ByVal nct As Double, ByRef nc_derivative As Double) As Double
'y is 1-x but held accurately to avoid possible cancellation errors
'nc_derivative holds t * derivative
   Dim AA As Double, BB As Double, nc_dtemp As Double
   Dim n As Double, p As Double, q As Double, W As Double, V As Double, r As Double, s As Double, ps As Double
   Dim result1 As Double, result2 As Double, term1 As Double, term2 As Double, ptnc As Double, qtnc As Double, ptx As Double, qtx As Double
   Dim A As Double, B As Double, x As Double, Y As Double, nc As Double
   Dim save_result1 As Double, save_result2 As Double, vScale As Double
   A = 0.5
   B = df / 2#
   If Abs(t) >= Min(1#, df) Then
      Y = df / t
      x = t + Y
      Y = Y / x
      x = t / x
   Else
      x = t * t
      Y = df + x
      x = x / Y
      Y = df / Y
   End If
   If Y < cSmall Then
      comp_t_nc1 = [#VALUE!]
      Exit Function
   End If
   nc = nct * nct / 2#
   AA = A - nc * x * (A + B)
   BB = (x * nc - 1#) - A
   If (BB < 0#) Then
      n = BB - Sqr(BB ^ 2 - 4# * AA)
      n = Int(2# * AA / n)
   Else
      n = Int((BB + Sqr(BB ^ 2 - 4# * AA)) / 2)
   End If
   If n < 0# Then
      n = 0#
   End If
   AA = n + A
   BB = n + 0.5
   qtnc = poissonTerm(BB, nc, nc - BB, 0#)
   BB = n
   ptnc = poissonTerm(BB, nc, nc - BB, 0#)
   ptx = binomialTerm(AA, B, x, Y, B * x - AA * Y, 0#) / (AA + B) '((1 - I(x, aa+1, b)) - (1 - I(x, aa, b)))/b
   qtx = binomialTerm(AA + 0.5, B, x, Y, B * x - (AA + 0.5) * Y, 0#) / (AA + B + 0.5) '((1 - I(x, aa+3/2, b)) - (1 - I(x, aa+1/2, b)))/b
   If B > 1# Then
      ptx = B * ptx
      qtx = B * qtx
   End If
   vScale = max(ptnc, qtnc)
   If ptnc = vScale Then
      p = 1#
   Else
      p = ptnc / vScale
   End If
   If qtnc = vScale Then
      q = 1#
   Else
      q = qtnc / vScale
   End If
   nc_derivative = 0#
   s = ptx
   r = qtx
   If x > Y Then
      V = beta(Y, B, AA + 0.5) '  1 - I(x, aa+1/2, b)
      W = beta(Y, B, AA) '  1 - I(x, aa, b)
   Else
      V = compbeta(x, AA + 0.5, B) ' 1 - I(x, aa+1/2, b)
      W = compbeta(x, AA, B) ' 1 - I(x, aa, b)
   End If
   term1 = 0#
   term2 = 0#
   result1 = term1
   result2 = term2
   Do
       If B > 1# Then
          W = W + s ' 1 - I(x, aa, b)
          V = V + r ' 1 - I(x, aa+1/2, b)
       Else
          W = W + B * s
          V = V + B * r
       End If
       s = (AA + B) * s
       r = (AA + B + 0.5) * r
       AA = AA + 1#
       BB = BB + 1#
       p = nc / BB * p
       q = nc / (BB + 0.5) * q
       ps = p * s + q * r
       nc_derivative = nc_derivative + ps
       s = x / AA * s ' (1 - I(x, aa+1, b)) - (1 - I(x, aa, b))
       r = x / (AA + 0.5) * r ' (1 - I(x, aa+3/2, b)) - (1 - I(x, aa+1/2, b))
       term1 = p * W
       term2 = q * V
       result1 = result1 + term1
       result2 = result2 + term2
   Loop While ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) And (s > 1E-16 * W)) Or (ps > 1E-16 * nc_derivative))
   p = comppoisson(BB, nc, nc - BB)
   BB = BB + 0.5
   q = comppoisson(BB, nc, nc - BB)
   nc_derivative = x * nc_derivative * vScale
   save_result1 = result1 * vScale + p * W
   save_result2 = result2 * vScale + q * V
   ps = 1#
   nc_dtemp = 0#
   AA = n + A
   BB = n
   p = ptnc
   q = qtnc
   vScale = max(ptx, qtx)
   If ptx = vScale Then
      s = 1#
   Else
      s = ptx / vScale
   End If
   If qtx = vScale Then
      r = 1#
   Else
      r = qtx / vScale
   End If
   W = p
   V = q
   term1 = 1#
   term2 = 1#
   result1 = 0#
   result2 = 0#
   While BB > 0# And ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) And (p > 1E-16 * W)) Or (ps > 1E-16 * nc_dtemp))
      r = (AA + 0.5) / x * r
      s = AA / x * s
      ps = p * s + q * r
      nc_dtemp = nc_dtemp + ps
      p = BB / nc * p
      q = (BB + 0.5) / nc * q
      AA = AA - 1#
      BB = BB - 1#
      If BB = 0# Then AA = A
      r = r / (AA + B + 0.5) ' (1 - I(x, aa+3/2, b)) - (1 - I(x, aa+1/2, b))
      s = s / (AA + B) ' (1 - I(x, aa + 1, b)) - (1 - I(x, aa, b))
      term1 = s * W
      term2 = r * V
      result1 = result1 + term1
      result2 = result2 + term2
      W = W + p
      V = V + q
   Wend
   nc_dtemp = (x * nc_dtemp + p * AA * s + q * (AA + 0.5) * r) * vScale
   If x > Y Then
      r = beta(Y, B, A + (BB + 0.5))
      s = beta(Y, B, A + BB)
   Else
      r = compbeta(x, A + (BB + 0.5), B)
      s = compbeta(x, A + BB, B)
   End If
   If B <= 1# Then vScale = vScale * B
   result1 = save_result1 + result1 * vScale + s * W
   result2 = save_result2 + result2 * vScale + r * V
   If t > 0# Then
      comp_t_nc1 = 0.5 * (result1 + result2)
      nc_derivative = nc_derivative + nc_dtemp
   Else
      comp_t_nc1 = 1# - 0.5 * (result1 - result2)
   End If
End Function

Private Function inv_t_nc1(ByVal prob As Double, ByVal df As Double, ByVal nc As Double, ByRef oneMinusP As Double) As Double
'Uses approximations in A&S 26.6.26 and 26.7.10 for to get initial estimate, the modified NR to improve it.
Dim x As Double, Y As Double, pr As Double, dif As Double, temp As Double, nc_BETA_param As Double
Dim hix As Double, lox As Double, test As Double, nc_derivative As Double
   If (prob > 0.5) Then
      inv_t_nc1 = comp_inv_t_nc1(1# - prob, df, nc, oneMinusP)
      Exit Function
   End If
   nc_BETA_param = nc ^ 2 / 2#
   lox = 0#
   hix = t_nc_limit * Sqr(df)
   pr = Exp(-nc_BETA_param)
   If pr > prob Then
      If 2# * prob > pr Then
         x = invcompbeta(0.5, df / 2#, (pr - prob) / pr, oneMinusP)
      Else
         x = invbeta(0.5, df / 2#, prob / pr, oneMinusP)
      End If
      If x = 0# Then
         inv_t_nc1 = 0#
         Exit Function
      Else
         temp = oneMinusP
         Y = invbeta((0.5 + nc_BETA_param) ^ 2 / (0.5 + 2# * nc_BETA_param), df / 2#, prob, oneMinusP)
         oneMinusP = (0.5 + nc_BETA_param) * oneMinusP / (0.5 + nc_BETA_param * (1# + Y))
         If temp > oneMinusP Then
            oneMinusP = temp
         Else
            x = (0.5 + 2# * nc_BETA_param) * Y / (0.5 + nc_BETA_param * (1# + Y))
         End If
         If oneMinusP < cSmall Then
            pr = t_nc1(hix, df, nc, nc_derivative)
            If pr < prob Then
               inv_t_nc1 = [#VALUE!]
               oneMinusP = 0#
               Exit Function
            End If
            oneMinusP = 4# * cSmall
         End If
      End If
   Else
      Y = invbeta((0.5 + nc_BETA_param) ^ 2 / (0.5 + 2# * nc_BETA_param), df / 2#, prob, oneMinusP)
      x = (0.5 + 2# * nc_BETA_param) * Y / (0.5 + nc_BETA_param * (1 + Y))
      oneMinusP = (0.5 + nc_BETA_param) * oneMinusP / (0.5 + nc_BETA_param * (1# + Y))
      If oneMinusP < cSmall Then
         pr = t_nc1(hix, df, nc, nc_derivative)
         If pr < prob Then
            inv_t_nc1 = [#VALUE!]
            oneMinusP = 0#
            Exit Function
         End If
         oneMinusP = 4# * cSmall
      End If
   End If
   test = Sqr(df * x) / Sqr(oneMinusP)
   Do
      pr = t_nc1(test, df, nc, nc_derivative)
      If pr < prob Then
         lox = test
      Else
         hix = test
      End If
      If nc_derivative = 0# Then
         If pr < prob Then
            dif = (hix - lox) / 2#
         Else
            dif = (lox - hix) / 2#
         End If
      Else
         dif = -(pr * test / nc_derivative) * logdif(pr, prob)
         If df < 2# Then dif = 2# * dif / df
         If test + dif < lox Then
            If lox = 0 Then
               dif = (lox - test) * 0.9999999999
            Else
               dif = (lox - test) * 0.9
            End If
         ElseIf test + dif > hix Then
            dif = (hix - test) * 0.9
         End If
      End If
      test = test + dif
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > test * 0.0000000001))
   inv_t_nc1 = test
End Function

Private Function comp_inv_t_nc1(ByVal prob As Double, ByVal df As Double, ByVal nc As Double, ByRef oneMinusP As Double) As Double
'Uses approximations in A&S 26.6.26 and 26.7.10 for to get initial estimate, the modified NR to improve it.
Dim x As Double, Y As Double, pr As Double, dif As Double, temp As Double, nc_BETA_param As Double
Dim hix As Double, lox As Double, test As Double, nc_derivative As Double
   If (prob > 0.5) Then
      comp_inv_t_nc1 = inv_t_nc1(1# - prob, df, nc, oneMinusP)
      Exit Function
   End If
   nc_BETA_param = nc ^ 2 / 2#
   lox = 0#
   hix = t_nc_limit * Sqr(df)
   pr = Exp(-nc_BETA_param)
   If pr > prob Then
      If 2# * prob > pr Then
         x = invbeta(0.5, df / 2#, (pr - prob) / pr, oneMinusP)
      Else
         x = invcompbeta(0.5, df / 2#, prob / pr, oneMinusP)
      End If
      If oneMinusP < cSmall Then
         pr = comp_t_nc1(hix, df, nc, nc_derivative)
         If pr > prob Then
            comp_inv_t_nc1 = [#VALUE!]
            oneMinusP = 0#
            Exit Function
         End If
         oneMinusP = 4# * cSmall
      Else
         temp = oneMinusP
         Y = invcompbeta((0.5 + nc_BETA_param) ^ 2 / (0.5 + 2# * nc_BETA_param), df / 2#, prob, oneMinusP)
         oneMinusP = (0.5 + nc_BETA_param) * oneMinusP / (0.5 + nc_BETA_param * (1# + Y))
         If temp < oneMinusP Then
            oneMinusP = temp
         Else
            x = (0.5 + 2# * nc_BETA_param) * Y / (0.5 + nc_BETA_param * (1# + Y))
         End If
         If oneMinusP < cSmall Then
            pr = comp_t_nc1(hix, df, nc, nc_derivative)
            If pr > prob Then
               comp_inv_t_nc1 = [#VALUE!]
               oneMinusP = 0#
               Exit Function
            End If
            oneMinusP = 4# * cSmall
         End If
      End If
   Else
      Y = invcompbeta((0.5 + nc_BETA_param) ^ 2 / (0.5 + 2# * nc_BETA_param), df / 2#, prob, oneMinusP)
      x = (0.5 + 2# * nc_BETA_param) * Y / (0.5 + nc_BETA_param * (1# + Y))
      oneMinusP = (0.5 + nc_BETA_param) * oneMinusP / (0.5 + nc_BETA_param * (1# + Y))
      If oneMinusP < cSmall Then
         pr = comp_t_nc1(hix, df, nc, nc_derivative)
         If pr > prob Then
            comp_inv_t_nc1 = [#VALUE!]
            oneMinusP = 0#
            Exit Function
         End If
         oneMinusP = 4# * cSmall
      End If
   End If
   test = Sqr(df * x) / Sqr(oneMinusP)
   dif = test
   Do
      pr = comp_t_nc1(test, df, nc, nc_derivative)
      If pr < prob Then
         hix = test
      Else
         lox = test
      End If
      If nc_derivative = 0# Then
         If pr < prob Then
            dif = (lox - hix) / 2#
         Else
            dif = (hix - lox) / 2#
         End If
      Else
         dif = (pr * test / nc_derivative) * logdif(pr, prob)
         If df < 2# Then dif = 2# * dif / df
         If test + dif < lox Then
            dif = (lox - test) * 0.9
         ElseIf test + dif > hix Then
            dif = (hix - test) * 0.9
         End If
      End If
      test = test + dif
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > test * 0.0000000001))
   comp_inv_t_nc1 = test
End Function

Private Function ncp_t_nc1(ByVal prob As Double, ByVal t As Double, ByVal df As Double) As Double
'Uses Normal approx for non-central t (A&S 26.7.10) to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, deriv As Double, checked_tnc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double, tnc_limit As Double, x As Double, Y As Double
   If (prob > 0.5) Then
      ncp_t_nc1 = comp_ncp_t_nc1(1# - prob, t, df)
      Exit Function
   End If

   lo = 0#
   tnc_limit = Sqr(2# * nc_limit)
   hi = tnc_limit
   checked_0_limit = False
   checked_tnc_limit = False
   If t >= Min(1#, df) Then
      Y = df / t
      x = t + Y
      Y = Y / x
      x = t / x
   Else
      x = t * t
      Y = df + x
      x = x / Y
      Y = df / Y
   End If
   temp = -inv_normal(prob)
   If t > df Then
        ncp = t * (1# - 0.25 / df) + temp * Sqr(t) * Sqr((1# / t + 0.5 * t / df))
   Else
        ncp = t * (1# - 0.25 / df) + temp * Sqr((1# + (0.5 * t / df) * t))
   End If
   ncp = max(temp, ncp)
   'Debug.Print "ncp_estimate1", ncp
   If x > 1E-200 Then 'I think we can put more accurate bounds on when this will not deliver a sensible answer
      temp = invcompgamma(0.5 * x * df, prob) / Y - 0.5
      If temp > 0 Then
         temp = Sqr(2# * temp)
         If temp > ncp Then
            ncp = temp
         End If
      End If
   End If
   'Debug.Print "ncp_estimate2", ncp
   ncp = Min(ncp, tnc_limit)
   If ncp = tnc_limit Then
      pr = cdf_t_nc(t, df, ncp)
      If pr > prob Then
         ncp_t_nc1 = [#VALUE!]
         Exit Function
      Else
         checked_tnc_limit = True
      End If
   End If
   dif = ncp
   Do
      pr = cdf_t_nc(t, df, ncp)
      'Debug.Print ncp, pr, prob
      If ncp > 1 Then
         deriv = cdf_t_nc(t, df, ncp * (1 - 0.000001))
         deriv = 1000000# * (deriv - pr) / ncp
      ElseIf ncp > 0.000001 Then
         deriv = cdf_t_nc(t, df, ncp + 0.000001)
         deriv = 1000000# * (pr - deriv)
      ElseIf x < Y Then
         deriv = comp_cdf_BETA(x, 1, df / 2) * OneOverSqrTwoPi
      Else
         deriv = cdf_BETA(Y, df / 2, 1) * OneOverSqrTwoPi
      End If
      If pr < 3E-308 And deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      ElseIf deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr < prob Then
            hi = ncp
         Else
            lo = ncp
         End If
         dif = (pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = cdf_t_nc(t, df, lo)
               If temp < prob Then
                  If invtdist(prob, df) <= t Then
                     ncp_t_nc1 = 0#
                  Else
                     ncp_t_nc1 = [#VALUE!]
                  End If
                  Exit Function
               Else
                  checked_0_limit = True
               End If
               dif = dif * 1.99999999
            End If
         ElseIf ncp + dif > hi Then
            dif = (hi - ncp) / 2#
            If Not checked_tnc_limit And (hi = tnc_limit) Then
               temp = cdf_t_nc(t, df, hi)
               If temp > prob Then
                  ncp_t_nc1 = [#VALUE!]
                  Exit Function
               Else
                  checked_tnc_limit = True
               End If
               dif = dif * 1.99999999
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   ncp_t_nc1 = ncp
   'Debug.Print "ncp_t_nc1", ncp_t_nc1
End Function

Private Function comp_ncp_t_nc1(ByVal prob As Double, ByVal t As Double, ByVal df As Double) As Double
'Uses Normal approx for non-central t (A&S 26.7.10) to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, temp1 As Double, temp2 As Double, deriv As Double, checked_tnc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double, tnc_limit As Double, x As Double, Y As Double
   If (prob > 0.5) Then
      comp_ncp_t_nc1 = ncp_t_nc1(1# - prob, t, df)
      Exit Function
   End If

   lo = 0#
   tnc_limit = Sqr(2# * nc_limit)
   hi = tnc_limit
   checked_0_limit = False
   checked_tnc_limit = False
   If t >= Min(1#, df) Then
      Y = df / t
      x = t + Y
      Y = Y / x
      x = t / x
   Else
      x = t * t
      Y = df + x
      x = x / Y
      Y = df / Y
   End If
   temp = -inv_normal(prob)
   temp1 = t * (1# - 0.25 / df)
   If t > df Then
        temp2 = temp * Sqr(t) * Sqr((1# / t + 0.5 * t / df))
   Else
        temp2 = temp * Sqr((1# + (0.5 * t / df) * t))
   End If
   ncp = max(temp, temp1 + temp2)
   'Debug.Print "comp_ncp ncp estimate1", ncp
   If x > 1E-200 Then 'I think we can put more accurate bounds on when this will not deliver a sensible answer
      temp = invcompgamma(0.5 * x * df, prob) / Y - 0.5
      If temp > 0 Then
         temp = Sqr(2# * temp)
         If temp > ncp Then
            temp = invgamma(0.5 * x * df, prob) / Y - 0.5
            If temp > 0 Then
               ncp = Sqr(2# * temp)
            Else
               ncp = 0
            End If
         Else
            ncp = temp1 - temp2
         End If
      Else
         ncp = temp1 - temp2
      End If
   Else
      ncp = temp1 - temp2
   End If
   ncp = Min(max(0#, ncp), tnc_limit)
   If ncp = 0# Then
      pr = comp_cdf_t_nc(t, df, 0#)
      If pr > prob Then
         If -invtdist(prob, df) <= t Then
            comp_ncp_t_nc1 = 0#
         Else
            comp_ncp_t_nc1 = [#VALUE!]
         End If
         Exit Function
      ElseIf Abs(pr - prob) <= -prob * 0.00000000000001 * Log(pr) Then
         comp_ncp_t_nc1 = 0#
         Exit Function
      Else
         checked_0_limit = True
      End If
      If x < Y Then
         deriv = -comp_cdf_BETA(x, 1, 0.5 * df) * OneOverSqrTwoPi
      Else
         deriv = -cdf_BETA(Y, 0.5 * df, 1) * OneOverSqrTwoPi
      End If
      If deriv = 0# Then
         ncp = tnc_limit
      Else
         ncp = (pr - prob) / deriv
         If ncp >= tnc_limit Then
            ncp = (pr / deriv) * logdif(pr, prob) 'If these two are miles apart then best to take invgamma estimate if > 0
         End If
      End If
   End If
   ncp = Min(ncp, tnc_limit)
   If ncp = tnc_limit Then
      pr = comp_cdf_t_nc(t, df, ncp)
      If pr < prob Then
         comp_ncp_t_nc1 = [#VALUE!]
         Exit Function
      Else
         checked_tnc_limit = True
      End If
   End If
   dif = ncp
   Do
      pr = comp_cdf_t_nc(t, df, ncp)
      'Debug.Print ncp, pr, prob
      If ncp > 1 Then
         deriv = comp_cdf_t_nc(t, df, ncp * (1 - 0.000001))
         deriv = 1000000# * (pr - deriv) / ncp
      ElseIf ncp > 0.000001 Then
         deriv = comp_cdf_t_nc(t, df, ncp + 0.000001)
         deriv = 1000000# * (deriv - pr)
      ElseIf x < Y Then
         deriv = comp_cdf_BETA(x, 1, 0.5 * df) * OneOverSqrTwoPi
      Else
         deriv = cdf_BETA(Y, 0.5 * df, 1) * OneOverSqrTwoPi
      End If
      If pr < 3E-308 And deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp - dif
      ElseIf deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr > prob Then
            hi = ncp
         Else
            lo = ncp
         End If
         dif = -(pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = comp_cdf_t_nc(t, df, lo)
               If temp > prob Then
                  If -invtdist(prob, df) <= t Then
                     comp_ncp_t_nc1 = 0#
                  Else
                     comp_ncp_t_nc1 = [#VALUE!]
                  End If
                  Exit Function
               Else
                  checked_0_limit = True
               End If
               dif = dif * 1.99999999
            End If
         ElseIf ncp + dif > hi Then
            dif = (hi - ncp) / 2#
            If Not checked_tnc_limit And (hi = tnc_limit) Then
               temp = comp_cdf_t_nc(t, df, hi)
               If temp < prob Then
                  comp_ncp_t_nc1 = [#VALUE!]
                  Exit Function
               Else
                  checked_tnc_limit = True
               End If
               dif = dif * 1.99999999
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   comp_ncp_t_nc1 = ncp
   'Debug.Print "comp_ncp_t_nc1", comp_ncp_t_nc1
End Function

Public Function pdf_t_nc(ByVal x As Double, ByVal df As Double, ByVal nc_param As Double) As Double
'// Calculate pdf of noncentral t
'// Deliberately set not to calculate when x and nc_param have opposite signs as the algorithm used is prone to cancellation error in these circumstances.
'// The user can access t_nc1,comp_t_nc1 directly and check on the accuracy of the results, if required
  Dim nc_derivative As Double
  df = AlterForIntegralChecks_df(df)
  If (x < 0#) And (nc_param <= 0#) Then
     pdf_t_nc = pdf_t_nc(-x, df, -nc_param)
  ElseIf (df <= 0#) Or (nc_param < 0#) Or (nc_param > Sqr(2# * nc_limit)) Then
     pdf_t_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     pdf_t_nc = [#VALUE!]
  ElseIf (x = 0# Or nc_param = 0#) Then
     pdf_t_nc = Exp(-nc_param ^ 2 / 2) * pdftdist(x, df)
  Else
     If (df < 1# Or x < 1# Or x <= nc_param) Then
        pdf_t_nc = t_nc1(x, df, nc_param, nc_derivative)
     Else
        pdf_t_nc = comp_t_nc1(x, df, nc_param, nc_derivative)
     End If
     If nc_derivative < cSmall Then
        pdf_t_nc = Exp(-nc_param ^ 2 / 2) * pdftdist(x, df)
     ElseIf df > 2# Then
        pdf_t_nc = nc_derivative / x
     Else
        pdf_t_nc = nc_derivative * (df / (2# * x))
     End If
  End If
  pdf_t_nc = GetRidOfMinusZeroes(pdf_t_nc)
End Function

Public Function cdf_t_nc(ByVal x As Double, ByVal df As Double, ByVal nc_param As Double) As Double
'// Calculate cdf of noncentral t
'// Deliberately set not to calculate when x and nc_param have opposite signs as the algorithm used is prone to cancellation error in these circumstances.
'// The user can access t_nc1,comp_t_nc1 directly and check on the accuracy of the results, if required
  Dim tdistDensity As Double, nc_derivative As Double
  df = AlterForIntegralChecks_df(df)
  If (nc_param = 0#) Then
     cdf_t_nc = tdist(x, df, tdistDensity)
  ElseIf (x <= 0#) And (nc_param < 0#) Then
     cdf_t_nc = comp_cdf_t_nc(-x, df, -nc_param)
  ElseIf (df <= 0#) Or (nc_param < 0#) Or (nc_param > Sqr(2# * nc_limit)) Then
     cdf_t_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     cdf_t_nc = [#VALUE!]
  ElseIf (df < 1# Or x < 1# Or x <= nc_param) Then
     cdf_t_nc = t_nc1(x, df, nc_param, nc_derivative)
  Else
     cdf_t_nc = 1# - comp_t_nc1(x, df, nc_param, nc_derivative)
  End If
  cdf_t_nc = GetRidOfMinusZeroes(cdf_t_nc)
End Function

Public Function comp_cdf_t_nc(ByVal x As Double, ByVal df As Double, ByVal nc_param As Double) As Double
'// Calculate 1-cdf of noncentral t
'// Deliberately set not to calculate when x and nc_param have opposite signs as the algorithm used is prone to cancellation error in these circumstances.
'// The user can access t_nc1,comp_t_nc1 directly and check on the accuracy of the results, if required
  Dim tdistDensity As Double, nc_derivative As Double
  df = AlterForIntegralChecks_df(df)
  If (nc_param = 0#) Then
     comp_cdf_t_nc = tdist(-x, df, tdistDensity)
  ElseIf (x <= 0#) And (nc_param < 0#) Then
     comp_cdf_t_nc = cdf_t_nc(-x, df, -nc_param)
  ElseIf (df <= 0#) Or (nc_param < 0#) Or (nc_param > Sqr(2# * nc_limit)) Then
     comp_cdf_t_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     comp_cdf_t_nc = [#VALUE!]
  ElseIf (df < 1# Or x < 1# Or x >= nc_param) Then
     comp_cdf_t_nc = comp_t_nc1(x, df, nc_param, nc_derivative)
  Else
     comp_cdf_t_nc = 1# - t_nc1(x, df, nc_param, nc_derivative)
  End If
  comp_cdf_t_nc = GetRidOfMinusZeroes(comp_cdf_t_nc)
End Function

Public Function inv_t_nc(ByVal prob As Double, ByVal df As Double, ByVal nc_param As Double) As Double
  df = AlterForIntegralChecks_df(df)
  If (nc_param = 0#) Then
     inv_t_nc = invtdist(prob, df)
  ElseIf (nc_param < 0#) Then
     inv_t_nc = -comp_inv_t_nc(prob, df, -nc_param)
  ElseIf (df <= 0# Or nc_param > Sqr(2# * nc_limit) Or prob <= 0# Or prob >= 1#) Then
     inv_t_nc = [#VALUE!]
  ElseIf (invcnormal(prob) < -nc_param) Then
     inv_t_nc = [#VALUE!]
  Else
     Dim oneMinusP As Double
     inv_t_nc = inv_t_nc1(prob, df, nc_param, oneMinusP)
  End If
  inv_t_nc = GetRidOfMinusZeroes(inv_t_nc)
End Function

Public Function comp_inv_t_nc(ByVal prob As Double, ByVal df As Double, ByVal nc_param As Double) As Double
  df = AlterForIntegralChecks_df(df)
  If (nc_param = 0#) Then
     comp_inv_t_nc = -invtdist(prob, df)
  ElseIf (nc_param < 0#) Then
     comp_inv_t_nc = -inv_t_nc(prob, df, -nc_param)
  ElseIf (df <= 0# Or nc_param > Sqr(2# * nc_limit) Or prob <= 0# Or prob >= 1#) Then
     comp_inv_t_nc = [#VALUE!]
  ElseIf (invcnormal(prob) > nc_param) Then
     comp_inv_t_nc = [#VALUE!]
  Else
     Dim oneMinusP As Double
     comp_inv_t_nc = comp_inv_t_nc1(prob, df, nc_param, oneMinusP)
  End If
  comp_inv_t_nc = GetRidOfMinusZeroes(comp_inv_t_nc)
End Function

Public Function ncp_t_nc(ByVal prob As Double, ByVal x As Double, ByVal df As Double) As Double
  df = AlterForIntegralChecks_df(df)
  If (x = 0# And prob > 0.5) Then
     ncp_t_nc = -invcnormal(prob)
  ElseIf (x < 0) Then
     ncp_t_nc = -comp_ncp_t_nc(prob, -x, df)
  ElseIf (df <= 0# Or prob <= 0# Or prob >= 1#) Then
     ncp_t_nc = [#VALUE!]
  Else
     ncp_t_nc = ncp_t_nc1(prob, x, df)
  End If
  ncp_t_nc = GetRidOfMinusZeroes(ncp_t_nc)
End Function

Public Function comp_ncp_t_nc(ByVal prob As Double, ByVal x As Double, ByVal df As Double) As Double
  df = AlterForIntegralChecks_df(df)
  If (x = 0#) Then
     comp_ncp_t_nc = invcnormal(prob)
  ElseIf (x < 0) Then
     comp_ncp_t_nc = -ncp_t_nc(prob, -x, df)
  ElseIf (df <= 0# Or prob <= 0# Or prob >= 1#) Then
     comp_ncp_t_nc = [#VALUE!]
  Else
     comp_ncp_t_nc = comp_ncp_t_nc1(prob, x, df)
  End If
  comp_ncp_t_nc = GetRidOfMinusZeroes(comp_ncp_t_nc)
End Function

Public Function pmf_GammaPoisson(i As Double, gamma_shape As Double, gamma_scale As Double) As Double
   Dim p As Double, q As Double, dfm As Double
   q = gamma_scale / (1# + gamma_scale)
   p = 1# / (1# + gamma_scale)
   i = AlterForIntegralChecks_Others(i)
   If (gamma_shape <= 0# Or gamma_scale <= 0#) Then
      pmf_GammaPoisson = [#VALUE!]
   ElseIf (i < 0#) Then
      pmf_GammaPoisson = 0
   Else
      If p < q Then
         dfm = gamma_shape - (gamma_shape + i) * p
      Else
         dfm = (gamma_shape + i) * q - i
      End If
      pmf_GammaPoisson = (gamma_shape / (gamma_shape + i)) * binomialTerm(i, gamma_shape, q, p, dfm, 0#)
   End If
   pmf_GammaPoisson = GetRidOfMinusZeroes(pmf_GammaPoisson)
End Function

Public Function cdf_GammaPoisson(i As Double, gamma_shape As Double, gamma_scale As Double) As Double
   Dim p As Double, q As Double
   q = gamma_scale / (1# + gamma_scale)
   p = 1# / (1# + gamma_scale)
   i = Int(i)
   If (gamma_shape <= 0# Or gamma_scale <= 0#) Then
      cdf_GammaPoisson = [#VALUE!]
   ElseIf (i < 0#) Then
      cdf_GammaPoisson = 0#
   ElseIf (p <= q) Then
      cdf_GammaPoisson = beta(p, gamma_shape, i + 1#)
   Else
      cdf_GammaPoisson = compbeta(q, i + 1#, gamma_shape)
   End If
   cdf_GammaPoisson = GetRidOfMinusZeroes(cdf_GammaPoisson)
End Function

Public Function comp_cdf_GammaPoisson(i As Double, gamma_shape As Double, gamma_scale As Double) As Double
   Dim p As Double, q As Double
   q = gamma_scale / (1# + gamma_scale)
   p = 1# / (1# + gamma_scale)
   i = Int(i)
   If (gamma_shape <= 0# Or gamma_scale <= 0#) Then
      comp_cdf_GammaPoisson = [#VALUE!]
   ElseIf (i < 0#) Then
      comp_cdf_GammaPoisson = 1#
   ElseIf (p <= q) Then
      comp_cdf_GammaPoisson = compbeta(p, gamma_shape, i + 1#)
   Else
      comp_cdf_GammaPoisson = beta(q, i + 1#, gamma_shape)
   End If
   comp_cdf_GammaPoisson = GetRidOfMinusZeroes(comp_cdf_GammaPoisson)
End Function

Public Function crit_GammaPoisson(gamma_shape As Double, gamma_scale As Double, crit_prob As Double) As Double
   Dim p As Double, q As Double
   q = gamma_scale / (1# + gamma_scale)
   p = 1# / (1# + gamma_scale)
   If (gamma_shape < 0# Or gamma_scale < 0#) Then
      crit_GammaPoisson = [#VALUE!]
   ElseIf (crit_prob < 0# Or crit_prob >= 1#) Then
      crit_GammaPoisson = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_GammaPoisson = [#VALUE!]
   Else
      Dim i As Double, pr As Double
      crit_GammaPoisson = critnegbinom(gamma_shape, p, q, crit_prob)
      i = crit_GammaPoisson
      If p <= q Then
         pr = beta(p, gamma_shape, i + 1#)
      Else
         pr = compbeta(q, i + 1#, gamma_shape)
      End If
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         If p <= q Then
            pr = beta(p, gamma_shape, i + 1#)
         Else
            pr = compbeta(q, i + 1#, gamma_shape)
         End If
         If (pr >= crit_prob) Then
            crit_GammaPoisson = i
         End If
      Else
         crit_GammaPoisson = i + 1#
      End If
   End If
   crit_GammaPoisson = GetRidOfMinusZeroes(crit_GammaPoisson)
End Function

Public Function comp_crit_GammaPoisson(gamma_shape As Double, gamma_scale As Double, crit_prob As Double) As Double
   Dim p As Double, q As Double
   q = gamma_scale / (1# + gamma_scale)
   p = 1# / (1# + gamma_scale)
   If (gamma_shape < 0# Or gamma_scale < 0#) Then
      comp_crit_GammaPoisson = [#VALUE!]
   ElseIf (crit_prob <= 0# Or crit_prob > 1#) Then
      comp_crit_GammaPoisson = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_GammaPoisson = [#VALUE!]
   Else
      Dim i As Double, pr As Double
      comp_crit_GammaPoisson = critcompnegbinom(gamma_shape, p, q, crit_prob)
      i = comp_crit_GammaPoisson
      If p <= q Then
         pr = compbeta(p, gamma_shape, i + 1#)
      Else
         pr = beta(q, i + 1#, gamma_shape)
      End If
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         If p <= q Then
            pr = compbeta(p, gamma_shape, i + 1#)
         Else
            pr = beta(q, i + 1#, gamma_shape)
         End If
         If (pr <= crit_prob) Then
            comp_crit_GammaPoisson = i
         End If
      Else
         comp_crit_GammaPoisson = i + 1#
      End If
   End If
   comp_crit_GammaPoisson = GetRidOfMinusZeroes(comp_crit_GammaPoisson)
End Function

Private Function PBB(ByVal i As Double, ByVal ssmi As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double) As Double
    PBB = (BETA_shape1 / (i + BETA_shape1)) * (BETA_shape2 / (BETA_shape1 + BETA_shape2)) * ((i + ssmi + BETA_shape1 + BETA_shape2) / (ssmi + BETA_shape2)) * hypergeometricTerm(i, ssmi, BETA_shape1, BETA_shape2)
End Function

Public Function pmf_BetaNegativeBinomial(ByVal i As Double, ByVal r As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double) As Double
   i = AlterForIntegralChecks_Others(i)
   If (r <= 0# Or BETA_shape1 <= 0# Or BETA_shape2 <= 0#) Then
      pmf_BetaNegativeBinomial = [#VALUE!]
   ElseIf i < 0 Then
      pmf_BetaNegativeBinomial = 0#
   Else
      pmf_BetaNegativeBinomial = (BETA_shape2 / (BETA_shape1 + BETA_shape2)) * (r / (BETA_shape1 + r)) * BETA_shape1 * (i + BETA_shape1 + r + BETA_shape2) / ((i + r) * (i + BETA_shape2)) * hypergeometricTerm(i, r, BETA_shape2, BETA_shape1)
   End If
   pmf_BetaNegativeBinomial = GetRidOfMinusZeroes(pmf_BetaNegativeBinomial)
End Function

Private Function CBNB0(ByVal i As Double, ByVal r As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double, toBeAdded As Double) As Double
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   Dim mrb2 As Double, other As Double, temp As Double
   If (r < 2# Or BETA_shape2 < 2#) Then
'Assumption here that i is integral or greater than 4.
      mrb2 = max(r, BETA_shape2)
      other = Min(r, BETA_shape2)
      CBNB0 = PBB(i, other, mrb2, BETA_shape1)
      If i = 0# Then Exit Function
      CBNB0 = CBNB0 * (1# + i * (other + BETA_shape1) / (((i - 1#) + mrb2) * (other + 1#)))
      If i = 1# Then Exit Function
      i = i - 2#
      other = other + 2#
      temp = PBB(i, mrb2, other, BETA_shape1)
      If i = 0# Then
         CBNB0 = CBNB0 + temp
         Exit Function
      End If
      CBNB0 = CBNB0 + temp * (1# + i * (mrb2 + BETA_shape1) / (((i - 1#) + other) * (mrb2 + 1#)))
      If i = 1# Then Exit Function
      i = i - 2#
      mrb2 = mrb2 + 2#
      CBNB0 = CBNB0 + CBNB0(i, mrb2, BETA_shape1, other, CBNB0)
   ElseIf (BETA_shape1 < 1#) Then
      mrb2 = max(r, BETA_shape2)
      other = Min(r, BETA_shape2)
      CBNB0 = hypergeometric(i, mrb2 - 1#, other, BETA_shape1, False, ha1, hprob, hswap)
      If hswap Then
         temp = PBB(mrb2 - 1#, BETA_shape1, i + 1#, other)
         If (toBeAdded + (CBNB0 - temp)) < 0.1 * (toBeAdded + (CBNB0 + temp)) Then
            CBNB0 = CBNB2(i, mrb2, BETA_shape1, other)
         Else
            CBNB0 = CBNB0 - temp
         End If
      ElseIf ha1 < -0.9 * BETA_shape1 / (BETA_shape1 + other) Then
         CBNB0 = [#VALUE!]
      Else
         CBNB0 = hprob * (BETA_shape1 / (BETA_shape1 + other) + ha1)
      End If
   Else
      CBNB0 = hypergeometric(i, r, BETA_shape2, BETA_shape1 - 1#, False, ha1, hprob, hswap)
   End If
End Function

Private Function CBNB2(ByVal i As Double, ByVal r As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double) As Double
   Dim j As Double, ss As Double, bs2 As Double, temp As Double, D1 As Double, d2 As Double, d_count As Double, pbbval As Double
   'In general may be a good idea to take Min(i, BETA_shape1) down to just above 0 and then work on Max(i, BETA_shape1)
   ss = Min(r, BETA_shape2)
   bs2 = max(r, BETA_shape2)
   r = ss
   BETA_shape2 = bs2
   D1 = (i + 0.5) * (BETA_shape1 + 0.5) - (bs2 - 1.5) * (ss - 0.5)
   If D1 < 0# Then
      CBNB2 = CBNB0(i, ss, BETA_shape1, bs2, 0#)
      Exit Function
   End If
   D1 = Int(D1 / (bs2 + BETA_shape1)) + 1#
   If ss + D1 > bs2 Then D1 = Int(bs2 - ss)
   ss = ss + D1
   j = i - D1
   d2 = (j + 0.5) * (BETA_shape1 + 0.5) - (bs2 - 1.5) * (ss - 0.5)
   If d2 < 0# Then
      d2 = 2#
   Else
   'Could make this smaller
      d2 = Int(Sqr(d2)) + 2#
   End If
   If 2# * d2 > i Then
      d2 = Int(i / 2#)
   End If
   pbbval = PBB(i, r, BETA_shape2, BETA_shape1)
   ss = ss + d2
   bs2 = bs2 + d2
   j = j - 2# * d2
   CBNB2 = CBNB0(j, ss, BETA_shape1, bs2, 0#)
   temp = 1#
   d_count = d2 - 2#
   j = j + 1#
   Do While d_count >= 0#
      j = j + 1#
      bs2 = BETA_shape2 + d_count
      d_count = d_count - 1#
      temp = 1# + (j * (bs2 + BETA_shape1) / ((j + ss - 1#) * (bs2 + 1#))) * temp
   Loop
   j = i - d2 - D1
   temp = (ss * (j + bs2)) / (bs2 * (j + ss)) * temp
   d_count = D1 + d2 - 1#
   Do While d_count >= 0
      j = j + 1#
      ss = r + d_count
      d_count = d_count - 1#
      temp = 1# + (j * (ss + BETA_shape1) / ((j + bs2 - 1#) * (ss + 1#))) * temp
   Loop
   CBNB2 = CBNB2 + temp * pbbval
   Exit Function
End Function

Public Function cdf_BetaNegativeBinomial(ByVal i As Double, ByVal r As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double) As Double
   i = Int(i)
   If (r <= 0# Or BETA_shape1 <= 0# Or BETA_shape2 <= 0#) Then
      cdf_BetaNegativeBinomial = [#VALUE!]
   ElseIf i < 0 Then
      cdf_BetaNegativeBinomial = 0#
   Else
      cdf_BetaNegativeBinomial = CBNB0(i, r, BETA_shape1, BETA_shape2, 0#)
   End If
   cdf_BetaNegativeBinomial = GetRidOfMinusZeroes(cdf_BetaNegativeBinomial)
End Function

Public Function comp_cdf_BetaNegativeBinomial(ByVal i As Double, ByVal r As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double) As Double
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   Dim mrb2 As Double, other As Double, temp As Double, CTEMP As Double
   i = Int(i)
   mrb2 = max(r, BETA_shape2)
   other = Min(r, BETA_shape2)
   If (other <= 0# Or BETA_shape1 <= 0#) Then
      comp_cdf_BetaNegativeBinomial = [#VALUE!]
   ElseIf i < 0# Then
      comp_cdf_BetaNegativeBinomial = 1#
   ElseIf (mrb2 > 100#) Then
      comp_cdf_BetaNegativeBinomial = CBNB0(mrb2 - 1#, i + 1#, other, BETA_shape1, 0#)
   ElseIf (other < 1# Or BETA_shape1 < 1#) Then
      If (i + BETA_shape1) < 60# Then
         i = i + 1#
         temp = pmf_BetaNegativeBinomial(i, r, BETA_shape1, BETA_shape2)
         CTEMP = temp
         Do While (i + BETA_shape1) < 60#
            temp = temp * (i + r) * (i + BETA_shape2) / ((i + 1#) * (i + BETA_shape1 + r + BETA_shape2))
            CTEMP = CTEMP + temp
            i = i + 1#
         Loop
      Else
         CTEMP = 0#
      End If
      If other >= 1# Then
         comp_cdf_BetaNegativeBinomial = hypergeometric(mrb2, i, BETA_shape1, other - 1#, False, ha1, hprob, hswap)
      Else
         comp_cdf_BetaNegativeBinomial = hypergeometric(mrb2, i, BETA_shape1, -other, False, ha1, hprob, hswap)
      End If
      If hswap Then
         temp = PBB(i, mrb2, other, BETA_shape1) 'N.B. subtraction of PBB term can be done exactly from hypergeometric one if hswap false
         If temp > 0.9 * comp_cdf_BetaNegativeBinomial Then
            comp_cdf_BetaNegativeBinomial = [#VALUE!]
         Else
            comp_cdf_BetaNegativeBinomial = (comp_cdf_BetaNegativeBinomial - temp) + CTEMP
         End If
      Else
         If ha1 < -0.9 * mrb2 / (BETA_shape1 + mrb2) Then
            comp_cdf_BetaNegativeBinomial = [#VALUE!]
         Else
            comp_cdf_BetaNegativeBinomial = hprob * (mrb2 / (BETA_shape1 + mrb2) + ha1) + CTEMP
         End If
      End If
   Else
      comp_cdf_BetaNegativeBinomial = hypergeometric(i, r, BETA_shape2, BETA_shape1 - 1#, True, ha1, hprob, hswap)
   End If
   comp_cdf_BetaNegativeBinomial = GetRidOfMinusZeroes(comp_cdf_BetaNegativeBinomial)
End Function

Private Function critbetanegbinomial(ByVal A As Double, ByVal B As Double, ByVal r As Double, ByVal cprob As Double) As Double
'//i such that Pr(betanegbinomial(i,r,a,b)) >= cprob and  Pr(betanegbinomial(i-1,r,a,b)) < cprob
   If (cprob > 0.5) Then
      critbetanegbinomial = critcompbetanegbinomial(A, B, r, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   If B > r Then
      i = B
      B = r
      r = i
   End If
   If (A < 10# Or B < 10#) Then
      If r < A And A < 1# Then
         pr = cprob * A / r
      Else
         pr = cprob
      End If
      i = invcompbeta(A, B, pr, oneMinusP)
   Else
      pr = r / (r + A + B - 1#)
      i = invcompbeta(A * pr, B * pr, cprob, oneMinusP)
   End If
   If i = 0# Then
      i = max_crit / 2#
   Else
      i = r * (oneMinusP / i)
      If i >= max_crit Then
         i = max_crit - 1#
      End If
   End If
   While (True)
      If (i < 0#) Then
         i = 0#
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         critbetanegbinomial = [#VALUE!]
         Exit Function
      End If
      pr = CBNB0(i, r, A, B, 0#)
      tpr = 0#
      If (pr > cprob * (1 + cfSmall)) Then
         If (i = 0#) Then
            critbetanegbinomial = 0#
            Exit Function
         End If
         tpr = pmf_BetaNegativeBinomial(i, r, A, B)
         If (pr < (1# + 0.00001) * tpr) Then
            tpr = tpr * (((i + 1#) * (i + A + r + B)) / ((i + r) * (i + B)))
            i = i - 1#
            While (tpr > cprob)
               tpr = tpr * (((i + 1#) * (i + A + r + B)) / ((i + r) * (i + B)))
               i = i - 1#
            Wend
         Else
            pr = pr - tpr
            If (pr < cprob) Then
               critbetanegbinomial = i
               Exit Function
            End If
            i = i - 1#
            If (i = 0#) Then
               critbetanegbinomial = 0#
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               If temp > i Then
                  i = i / 10#
               Else
                  i = Int(i - temp)
                  temp2 = pmf_BetaNegativeBinomial(i, r, A, B)
                  i = i - temp * (tpr - temp2) / (2# * temp2)
               End If
            Else
               tpr = tpr * (((i + 1#) * (i + A + r + B)) / ((i + r) * (i + B)))
               pr = pr - tpr
               If (pr < cprob) Then
                  critbetanegbinomial = i
                  Exit Function
               End If
               i = i - 1#
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     tpr = tpr * (((i + 1#) * (i + A + r + B)) / ((i + r) * (i + B)))
                     pr = pr - tpr
                     i = i - 1#
                  Wend
                  critbetanegbinomial = i + 1#
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log((((i + 1#) * (i + A + r + B)) / ((i + r) * (i + B)))) + 0.5)
                  i = i - temp
                  temp2 = pmf_BetaNegativeBinomial(i, r, A, B)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log((((i + 1#) * (i + A + r + B)) / ((i + r) * (i + B))))
                     i = i - temp
                  End If
               End If
            End If
         End If
      ElseIf ((1# + cfSmall) * pr < cprob) Then
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            tpr = pmf_BetaNegativeBinomial(i, r, A, B)
            pr = pr + tpr
         Wend
         temp = (cprob - pr) / tpr
         If temp <= 0# Then
            critbetanegbinomial = i
            Exit Function
         ElseIf temp < 10# Then
            While (pr < cprob)
               i = i + 1#
               tpr = tpr * (((i + r - 1#) * (i + B - 1#)) / (i * (i + A + r + B - 1#)))
               pr = pr + tpr
            Wend
            critbetanegbinomial = i
            Exit Function
         ElseIf i + temp > max_crit Then
            critbetanegbinomial = [#VALUE!]
            Exit Function
         Else
            i = Int(i + temp)
            temp2 = pmf_BetaNegativeBinomial(i, r, A, B)
            If temp2 > 0# Then i = i + temp * (tpr - temp2) / (2# * temp2)
         End If
      Else
         critbetanegbinomial = i
         Exit Function
      End If
   Wend
End Function

Private Function critcompbetanegbinomial(ByVal A As Double, ByVal B As Double, ByVal r As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(betanegbinomial(i,r,a,b)) > cprob and  1-Pr(betanegbinomial(i-1,r,a,b)) <= cprob
   If (cprob > 0.5) Then
      critcompbetanegbinomial = critbetanegbinomial(A, B, r, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   If B > r Then
      i = B
      B = r
      r = i
   End If
   If (A < 10# Or B < 10#) Then
      If r < A And A < 1# Then
         pr = cprob * A / r
      Else
         pr = cprob
      End If
      i = invbeta(A, B, pr, oneMinusP)
   Else
      pr = r / (r + A + B - 1#)
      i = invbeta(A * pr, B * pr, cprob, oneMinusP)
   End If
   If i = 0# Then
      i = max_crit / 2#
   Else
      i = r * (oneMinusP / i)
      If i >= max_crit Then
         i = max_crit - 1#
      End If
   End If
   While (True)
      If (i < 0#) Then
         i = 0#
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         critcompbetanegbinomial = [#VALUE!]
         Exit Function
      End If
      pr = comp_cdf_BetaNegativeBinomial(i, r, A, B)
      tpr = 0#
      If (pr > cprob * (1 + cfSmall)) Then
         i = i + 1#
         tpr = pmf_BetaNegativeBinomial(i, r, A, B)
         If (pr < (1 + 0.00001) * tpr) Then
            While (tpr > cprob)
               i = i + 1#
               tpr = tpr * (((i + r - 1#) * (i + B - 1#)) / (i * (i + A + r + B - 1#)))
            Wend
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcompbetanegbinomial = i
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = pmf_BetaNegativeBinomial(i, r, A, B)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            Else
               i = i + 1#
               tpr = tpr * (((i + r - 1#) * (i + B - 1#)) / (i * (i + A + r + B - 1#)))
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcompbetanegbinomial = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * (((i + r - 1#) * (i + B - 1#)) / (i * (i + A + r + B - 1#)))
                     pr = pr - tpr
                  Wend
                  critcompbetanegbinomial = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log((((i + r - 1#) * (i + B - 1#)) / (i * (i + A + r + B - 1#)))) + 0.5)
                  i = i + temp
                  temp2 = pmf_BetaNegativeBinomial(i, r, A, B)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log((((i + r - 1#) * (i + B - 1#)) / (i * (i + A + r + B - 1#))))
                     i = i + temp
                  End If
               End If
            End If
         End If
      ElseIf ((1# + cfSmall) * pr < cprob) Then
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = pmf_BetaNegativeBinomial(i, r, A, B)
            pr = pr + tpr
            i = i - 1#
         Wend
         temp = (cprob - pr) / tpr
         If temp <= 0# Then
            critcompbetanegbinomial = i + 1#
            Exit Function
         ElseIf temp < 100# Or i < 1000# Then
            While (pr <= cprob)
               tpr = tpr * (((i + 1#) * (i + A + r + B)) / ((i + r) * (i + B)))
               pr = pr + tpr
               i = i - 1#
            Wend
            critcompbetanegbinomial = i + 1#
            Exit Function
         ElseIf temp > i Then
            i = i / 10#
         Else
            i = Int(i - temp)
            temp2 = pmf_BetaNegativeBinomial(i, r, A, B)
            If temp2 > 0# Then i = i - temp * (tpr - temp2) / (2# * temp2)
         End If
      Else
         critcompbetanegbinomial = i
         Exit Function
      End If
   Wend
End Function

Public Function crit_BetaNegativeBinomial(ByVal r As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double, ByVal crit_prob As Double) As Double
   If (BETA_shape1 <= 0# Or BETA_shape2 <= 0# Or r <= 0#) Then
      crit_BetaNegativeBinomial = [#VALUE!]
   ElseIf (crit_prob < 0# Or crit_prob >= 1#) Then
      crit_BetaNegativeBinomial = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_BetaNegativeBinomial = [#VALUE!]
   Else
      crit_BetaNegativeBinomial = critbetanegbinomial(BETA_shape1, BETA_shape2, r, crit_prob)
   End If
   crit_BetaNegativeBinomial = GetRidOfMinusZeroes(crit_BetaNegativeBinomial)
End Function

Public Function comp_crit_BetaNegativeBinomial(ByVal r As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double, ByVal crit_prob As Double) As Double
   If (BETA_shape1 <= 0# Or BETA_shape2 <= 0# Or r <= 0#) Then
      comp_crit_BetaNegativeBinomial = [#VALUE!]
   ElseIf (crit_prob <= 0# Or crit_prob > 1#) Then
      comp_crit_BetaNegativeBinomial = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_BetaNegativeBinomial = 0#
   Else
      comp_crit_BetaNegativeBinomial = critcompbetanegbinomial(BETA_shape1, BETA_shape2, r, crit_prob)
   End If
   comp_crit_BetaNegativeBinomial = GetRidOfMinusZeroes(comp_crit_BetaNegativeBinomial)
End Function

Public Function pmf_BetaBinomial(ByVal i As Double, ByVal SAMPLE_SIZE As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double) As Double
   i = AlterForIntegralChecks_Others(i)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (BETA_shape1 <= 0# Or BETA_shape2 <= 0# Or SAMPLE_SIZE < 0#) Then
      pmf_BetaBinomial = [#VALUE!]
   ElseIf i < 0 Or i > SAMPLE_SIZE Then
      pmf_BetaBinomial = 0#
   Else
      pmf_BetaBinomial = (BETA_shape1 / (i + BETA_shape1)) * (BETA_shape2 / (BETA_shape1 + BETA_shape2)) * ((SAMPLE_SIZE + BETA_shape1 + BETA_shape2) / (SAMPLE_SIZE - i + BETA_shape2)) * hypergeometricTerm(i, SAMPLE_SIZE - i, BETA_shape1, BETA_shape2)
   End If
   pmf_BetaBinomial = GetRidOfMinusZeroes(pmf_BetaBinomial)
End Function

Public Function cdf_BetaBinomial(ByVal i As Double, ByVal SAMPLE_SIZE As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double) As Double
   i = Int(i)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (BETA_shape1 <= 0# Or BETA_shape2 <= 0# Or SAMPLE_SIZE < 0#) Then
      cdf_BetaBinomial = [#VALUE!]
   ElseIf i < 0# Then
      cdf_BetaBinomial = 0#
   ElseIf i >= SAMPLE_SIZE Then
      cdf_BetaBinomial = 1#
   Else
      cdf_BetaBinomial = CBNB0(i, SAMPLE_SIZE - i, BETA_shape2, BETA_shape1, 0#)
   End If
   cdf_BetaBinomial = GetRidOfMinusZeroes(cdf_BetaBinomial)
End Function

Public Function comp_cdf_BetaBinomial(ByVal i As Double, ByVal SAMPLE_SIZE As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double) As Double
   i = Int(i)
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (BETA_shape1 <= 0# Or BETA_shape2 <= 0# Or SAMPLE_SIZE < 0#) Then
      comp_cdf_BetaBinomial = [#VALUE!]
   ElseIf i < 0# Then
      comp_cdf_BetaBinomial = 1#
   ElseIf i >= SAMPLE_SIZE Then
      comp_cdf_BetaBinomial = 0#
   Else
      comp_cdf_BetaBinomial = CBNB0(SAMPLE_SIZE - i - 1#, i + 1#, BETA_shape1, BETA_shape2, 0#)
   End If
   comp_cdf_BetaBinomial = GetRidOfMinusZeroes(comp_cdf_BetaBinomial)
End Function

Private Function critbetabinomial(ByVal A As Double, ByVal B As Double, ByVal ss As Double, ByVal cprob As Double) As Double
'//i such that Pr(betabinomial(i,ss,a,b)) >= cprob and  Pr(betabinomial(i-1,ss,a,b)) < cprob
   If (cprob > 0.5) Then
      critbetabinomial = critcompbetabinomial(A, B, ss, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   If (A + B < 1#) Then
      i = invbeta(A, B, cprob, oneMinusP) * ss
   Else
      pr = ss / (ss + A + B - 1#)
      i = invbeta(A * pr, B * pr, cprob, oneMinusP) * ss
   End If
   While (True)
      If (i < 0#) Then
         i = 0#
      ElseIf (i > ss) Then
         i = ss
      End If
      i = Int(i + 0.5)
      If (i >= max_discrete) Then
         critbetabinomial = i
         Exit Function
      End If
      pr = cdf_BetaBinomial(i, ss, A, B)
      tpr = 0#
      If (pr >= cprob * (1 + cfSmall)) Then
         If (i = 0#) Then
            critbetabinomial = 0#
            Exit Function
         End If
         tpr = pmf_BetaBinomial(i, ss, A, B)
         If (pr < (1# + 0.00001) * tpr) Then
            tpr = tpr * ((i + 1#) * (ss + B - i - 1#)) / ((A + i) * (ss - i))
            i = i - 1#
            While (tpr > cprob)
               tpr = tpr * ((i + 1#) * (ss + B - i - 1#)) / ((A + i) * (ss - i))
               i = i - 1#
            Wend
         Else
            pr = pr - tpr
            If (pr < cprob) Then
               critbetabinomial = i
               Exit Function
            End If
            i = i - 1#
            If (i = 0#) Then
               critbetabinomial = 0#
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10) Then
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = pmf_BetaBinomial(i, ss, A, B)
               i = i - temp * (tpr - temp2) / (2# * temp2)
            Else
               tpr = tpr * ((i + 1#) * (ss + B - i - 1#)) / ((A + i) * (ss - i))
               pr = pr - tpr
               If (pr < cprob) Then
                  critbetabinomial = i
                  Exit Function
               End If
               i = i - 1#
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     tpr = tpr * ((i + 1#) * (ss + B - i - 1#)) / ((A + i) * (ss - i))
                     pr = pr - tpr
                     i = i - 1#
                  Wend
                  critbetabinomial = i + 1#
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((i + 1#) * (ss + B - i - 1#)) / ((A + i) * (ss - i))) + 0.5)
                  i = i - temp
                  temp2 = pmf_BetaBinomial(i, ss, A, B)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((i + 1#) * (ss + B - i - 1#)) / ((A + i) * (ss - i)))
                     i = i - temp
                  End If
               End If
            End If
         End If
      ElseIf ((1# + cfSmall) * pr < cprob) Then
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            tpr = pmf_BetaBinomial(i, ss, A, B)
            pr = pr + tpr
         Wend
         temp = (cprob - pr) / tpr
         If temp <= 0# Then
            critbetabinomial = i
            Exit Function
         ElseIf temp < 10# Then
            While (pr < cprob)
               i = i + 1#
               tpr = tpr * ((A + i - 1#) * (ss - i + 1#)) / (i * (ss + B - i))
               pr = pr + tpr
            Wend
            critbetabinomial = i
            Exit Function
         ElseIf temp > 4E+15 Then
            i = 4E+15
         Else
            i = Int(i + temp)
            temp2 = pmf_BetaBinomial(i, ss, A, B)
            If temp2 > 0# Then i = i + temp * (tpr - temp2) / (2# * temp2)
         End If
      Else
         critbetabinomial = i
         Exit Function
      End If
   Wend
End Function

Private Function critcompbetabinomial(ByVal A As Double, ByVal B As Double, ByVal ss As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(betabinomial(i,ss,a,b)) > cprob and  1-Pr(betabinomial(i-1,ss,a,b)) <= cprob
   If (cprob > 0.5) Then
      critcompbetabinomial = critbetabinomial(A, B, ss, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   If (A + B < 1#) Then
      i = invcompbeta(A, B, cprob, oneMinusP) * ss
   Else
      pr = ss / (ss + A + B - 1#)
      i = invcompbeta(A * pr, B * pr, cprob, oneMinusP) * ss
   End If
   While (True)
      If (i < 0#) Then
         i = 0#
      ElseIf (i > ss) Then
         i = ss
      End If
      i = Int(i + 0.5)
      If (i >= max_discrete) Then
         critcompbetabinomial = i
         Exit Function
      End If
      pr = comp_cdf_BetaBinomial(i, ss, A, B)
      tpr = 0#
      If (pr >= cprob * (1 + cfSmall)) Then
         i = i + 1#
         tpr = pmf_BetaBinomial(i, ss, A, B)
         If (pr < (1 + 0.00001) * tpr) Then
            Do While (tpr > cprob)
               i = i + 1#
               temp = ss + B - i
               If temp = 0# Then Exit Do
               tpr = tpr * ((A + i - 1#) * (ss - i + 1#)) / (i * temp)
            Loop
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcompbetabinomial = i
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = pmf_BetaBinomial(i, ss, A, B)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            Else
               i = i + 1#
               tpr = tpr * ((A + i - 1#) * (ss - i + 1#)) / (i * (ss + B - i))
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcompbetabinomial = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * ((A + i - 1#) * (ss - i + 1#)) / (i * (ss + B - i))
                     pr = pr - tpr
                  Wend
                  critcompbetabinomial = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((A + i - 1#) * (ss - i + 1#)) / (i * (ss + B - i))) + 0.5)
                  i = i + temp
                  temp2 = pmf_BetaBinomial(i, ss, A, B)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((A + i - 1#) * (ss - i + 1#)) / (i * (ss + B - i)))
                     i = i + temp
                  End If
               End If
            End If
         End If
      ElseIf ((1# + cfSmall) * pr < cprob) Then
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = pmf_BetaBinomial(i, ss, A, B)
            pr = pr + tpr
            i = i - 1#
         Wend
         temp = (cprob - pr) / tpr
         If temp <= 0# Then
            critcompbetabinomial = i + 1#
            Exit Function
         ElseIf temp < 100# Or i < 1000# Then
            While (pr <= cprob)
               tpr = tpr * ((i + 1#) * (ss + B - i - 1#)) / ((A + i) * (ss - i))
               pr = pr + tpr
               i = i - 1#
            Wend
            critcompbetabinomial = i + 1#
            Exit Function
         ElseIf temp > i Then
            i = i / 10#
         Else
            i = Int(i - temp)
            temp2 = pmf_BetaNegativeBinomial(i, ss, A, B)
            If temp2 > 0# Then i = i - temp * (tpr - temp2) / (2# * temp2)
         End If
      Else
         critcompbetabinomial = i + 1#
         Exit Function
      End If
   Wend
End Function

Public Function crit_BetaBinomial(ByVal SAMPLE_SIZE As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double, ByVal crit_prob As Double) As Double
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (BETA_shape1 <= 0# Or BETA_shape2 <= 0# Or SAMPLE_SIZE < 0#) Then
      crit_BetaBinomial = [#VALUE!]
   ElseIf (crit_prob < 0# Or crit_prob > 1#) Then
      crit_BetaBinomial = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_BetaBinomial = [#VALUE!]
   ElseIf (SAMPLE_SIZE = 0# Or crit_prob = 1#) Then
      crit_BetaBinomial = SAMPLE_SIZE
   Else
      crit_BetaBinomial = critbetabinomial(BETA_shape1, BETA_shape2, SAMPLE_SIZE, crit_prob)
   End If
   crit_BetaBinomial = GetRidOfMinusZeroes(crit_BetaBinomial)
End Function

Public Function comp_crit_BetaBinomial(ByVal SAMPLE_SIZE As Double, ByVal BETA_shape1 As Double, ByVal BETA_shape2 As Double, ByVal crit_prob As Double) As Double
   SAMPLE_SIZE = AlterForIntegralChecks_Others(SAMPLE_SIZE)
   If (BETA_shape1 <= 0# Or BETA_shape2 <= 0# Or SAMPLE_SIZE < 0#) Then
      comp_crit_BetaBinomial = [#VALUE!]
   ElseIf (crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_BetaBinomial = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_BetaBinomial = 0#
   ElseIf (SAMPLE_SIZE = 0# Or crit_prob = 0#) Then
      comp_crit_BetaBinomial = SAMPLE_SIZE
   Else
      comp_crit_BetaBinomial = critcompbetabinomial(BETA_shape1, BETA_shape2, SAMPLE_SIZE, crit_prob)
   End If
   comp_crit_BetaBinomial = GetRidOfMinusZeroes(comp_crit_BetaBinomial)
End Function

Public Function pdf_normal_os(ByVal x As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1) As Double
 ' pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid N(0,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then pdf_normal_os = [#VALUE!]: Exit Function
    Dim N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    If x <= 0 Then
        pdf_normal_os = pdf_BETA(CNormal(x), N1 + r, -r) * pdf_normal(x)
    Else
        pdf_normal_os = pdf_BETA(CNormal(-x), -r, N1 + r) * pdf_normal(-x)
    End If
End Function
 
Public Function cdf_normal_os(ByVal x As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1) As Double
 ' cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid N(0,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then cdf_normal_os = [#VALUE!]: Exit Function
    Dim N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    If x <= 0 Then
        cdf_normal_os = cdf_BETA(CNormal(x), N1 + r, -r)
    Else
        cdf_normal_os = comp_cdf_BETA(CNormal(-x), -r, N1 + r)
    End If
End Function
 
Public Function comp_cdf_normal_os(ByVal x As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1) As Double
 ' 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid N(0,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_cdf_normal_os = [#VALUE!]: Exit Function
    Dim N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    If x <= 0 Then
        comp_cdf_normal_os = comp_cdf_BETA(CNormal(x), N1 + r, -r)
    Else
        comp_cdf_normal_os = cdf_BETA(CNormal(-x), -r, N1 + r)
    End If
End Function
 
Public Function inv_normal_os(ByVal p As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1) As Double
 ' inverse of cdf_normal_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 ' accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then inv_normal_os = [#VALUE!]: Exit Function
    Dim xp As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    xp = invbeta(N1 + r, -r, p, oneMinusxp)
    If Abs(xp - 0.5) < 0.00000000000001 And xp <> 0.5 Then If cdf_BETA(0.5, N1 + r, -r) = p Then inv_normal_os = 0: Exit Function
    If xp <= 0.5 Then
        inv_normal_os = inv_normal(xp)
    Else
        inv_normal_os = -inv_normal(oneMinusxp)
    End If
End Function
 
Public Function comp_inv_normal_os(ByVal p As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1) As Double
 ' inverse of comp_cdf_normal_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_inv_normal_os = [#VALUE!]: Exit Function
    Dim xp As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    xp = invcompbeta(N1 + r, -r, p, oneMinusxp)
    If Abs(xp - 0.5) < 0.00000000000001 And xp <> 0.5 Then If comp_cdf_BETA(0.5, N1 + r, -r) = p Then comp_inv_normal_os = 0: Exit Function
    If xp <= 0.5 Then
        comp_inv_normal_os = inv_normal(xp)
    Else
        comp_inv_normal_os = -inv_normal(oneMinusxp)
    End If
End Function

Public Function pdf_gamma_os(ByVal x As Double, ByVal shape_param As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal scale_param As Double = 1, Optional ByVal nc_param As Double = 0) As Double
 ' pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then pdf_gamma_os = [#VALUE!]: Exit Function
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_gamma_nc(x / scale_param, shape_param, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        pdf_gamma_os = pdf_BETA(p, N1 + r, -r) * pdf_gamma_nc(x / scale_param, shape_param, nc_param) / scale_param
    Else
        pdf_gamma_os = pdf_BETA(comp_cdf_gamma_nc(x / scale_param, shape_param, nc_param), -r, N1 + r) * pdf_gamma_nc(x / scale_param, shape_param, nc_param) / scale_param
    End If
End Function

Public Function cdf_gamma_os(ByVal x As Double, ByVal shape_param As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal scale_param As Double = 1, Optional ByVal nc_param As Double = 0) As Double
 ' cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then cdf_gamma_os = [#VALUE!]: Exit Function
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_gamma_nc(x / scale_param, shape_param, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        cdf_gamma_os = cdf_BETA(p, N1 + r, -r)
    Else
        cdf_gamma_os = comp_cdf_BETA(comp_cdf_gamma_nc(x / scale_param, shape_param, nc_param), -r, N1 + r)
    End If
End Function

Public Function comp_cdf_gamma_os(ByVal x As Double, ByVal shape_param As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal scale_param As Double = 1, Optional ByVal nc_param As Double = 0) As Double
 ' 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_cdf_gamma_os = [#VALUE!]: Exit Function
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_gamma_nc(x / scale_param, shape_param, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        comp_cdf_gamma_os = comp_cdf_BETA(p, N1 + r, -r)
    Else
        comp_cdf_gamma_os = cdf_BETA(comp_cdf_gamma_nc(x / scale_param, shape_param, nc_param), -r, N1 + r)
    End If
End Function

Public Function inv_gamma_os(ByVal p As Double, ByVal shape_param As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal scale_param As Double = 1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of cdf_gamma_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 ' accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then inv_gamma_os = [#VALUE!]: Exit Function
    Dim xp As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    xp = invbeta(N1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        inv_gamma_os = inv_gamma_nc(xp, shape_param, nc_param) * scale_param
    Else
        inv_gamma_os = comp_inv_gamma_nc(oneMinusxp, shape_param, nc_param) * scale_param
    End If
End Function

Public Function comp_inv_gamma_os(ByVal p As Double, ByVal shape_param As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal scale_param As Double = 1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of comp_cdf_gamma_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_inv_gamma_os = [#VALUE!]: Exit Function
    Dim xp As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    xp = invcompbeta(N1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        comp_inv_gamma_os = inv_gamma_nc(xp, shape_param, nc_param) * scale_param
    Else
        comp_inv_gamma_os = comp_inv_gamma_nc(oneMinusxp, shape_param, nc_param) * scale_param
    End If
End Function

Public Function pdf_chi2_os(ByVal x As Double, ByVal df As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid chi2(df) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then pdf_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_Chi2_nc(x, df, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        pdf_chi2_os = pdf_BETA(p, N1 + r, -r) * pdf_Chi2_nc(x, df, nc_param)
    Else
        pdf_chi2_os = pdf_BETA(comp_cdf_Chi2_nc(x, df, nc_param), -r, N1 + r) * pdf_Chi2_nc(x, df, nc_param)
    End If
End Function

Public Function cdf_chi2_os(ByVal x As Double, ByVal df As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid chi2(df) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then cdf_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_Chi2_nc(x, df, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        cdf_chi2_os = cdf_BETA(p, N1 + r, -r)
    Else
        cdf_chi2_os = comp_cdf_BETA(comp_cdf_Chi2_nc(x, df, nc_param), -r, N1 + r)
    End If
End Function

Public Function comp_cdf_chi2_os(ByVal x As Double, ByVal df As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid chi2(df) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_cdf_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_Chi2_nc(x, df, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        comp_cdf_chi2_os = comp_cdf_BETA(p, N1 + r, -r)
    Else
        comp_cdf_chi2_os = cdf_BETA(comp_cdf_Chi2_nc(x, df, nc_param), -r, N1 + r)
    End If
End Function

Public Function inv_chi2_os(ByVal p As Double, ByVal df As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of cdf_chi2_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 ' accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then inv_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim xp As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    xp = invbeta(N1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        inv_chi2_os = inv_Chi2_nc(xp, df, nc_param)
    Else
        inv_chi2_os = comp_inv_Chi2_nc(oneMinusxp, df, nc_param)
    End If
End Function

Public Function comp_inv_chi2_os(ByVal p As Double, ByVal df As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of comp_cdf_chi2_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_inv_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim xp As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    xp = invcompbeta(N1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        comp_inv_chi2_os = inv_Chi2_nc(xp, df, nc_param)
    Else
        comp_inv_chi2_os = comp_inv_Chi2_nc(oneMinusxp, df, nc_param)
    End If
End Function

Public Function pdf_F_os(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then pdf_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_fdist_nc(x, df1, df2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        pdf_F_os = pdf_BETA(p, N1 + r, -r) * pdf_fdist_nc(x, df1, df2, nc_param)
    Else
        pdf_F_os = pdf_BETA(comp_cdf_fdist_nc(x, df1, df2, nc_param), -r, N1 + r) * pdf_fdist_nc(x, df1, df2, nc_param)
    End If
End Function

Public Function cdf_F_os(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then cdf_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_fdist_nc(x, df1, df2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        cdf_F_os = cdf_BETA(p, N1 + r, -r)
    Else
        cdf_F_os = comp_cdf_BETA(comp_cdf_fdist_nc(x, df1, df2, nc_param), -r, N1 + r)
    End If
End Function

Public Function comp_cdf_F_os(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_cdf_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_fdist_nc(x, df1, df2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        comp_cdf_F_os = comp_cdf_BETA(p, N1 + r, -r)
    Else
        comp_cdf_F_os = cdf_BETA(comp_cdf_fdist_nc(x, df1, df2, nc_param), -r, N1 + r)
    End If
End Function

Public Function inv_F_os(ByVal p As Double, ByVal df1 As Double, ByVal df2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of cdf_F_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 ' accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then inv_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim xp As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    xp = invbeta(N1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        inv_F_os = inv_fdist_nc(xp, df1, df2, nc_param)
    Else
        inv_F_os = comp_inv_fdist_nc(oneMinusxp, df1, df2, nc_param)
    End If
End Function

Public Function comp_inv_F_os(ByVal p As Double, ByVal df1 As Double, ByVal df2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of comp_cdf_F_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_inv_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim xp As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    xp = invcompbeta(N1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        comp_inv_F_os = inv_fdist_nc(xp, df1, df2, nc_param)
    Else
        comp_inv_F_os = comp_inv_fdist_nc(oneMinusxp, df1, df2, nc_param)
    End If
End Function

Public Function pdf_BETA_os(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then pdf_BETA_os = [#VALUE!]: Exit Function
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_BETA_nc(x, shape_param1, shape_param2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        pdf_BETA_os = pdf_BETA(p, N1 + r, -r) * pdf_BETA_nc(x, shape_param1, shape_param2, nc_param)
    Else
        pdf_BETA_os = pdf_BETA(comp_cdf_BETA_nc(x, shape_param1, shape_param2, nc_param), -r, N1 + r) * pdf_BETA_nc(x, shape_param1, shape_param2, nc_param)
    End If
End Function

Public Function cdf_BETA_os(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then cdf_BETA_os = [#VALUE!]: Exit Function
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_BETA_nc(x, shape_param1, shape_param2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        cdf_BETA_os = cdf_BETA(p, N1 + r, -r)
    Else
        cdf_BETA_os = comp_cdf_BETA(comp_cdf_BETA_nc(x, shape_param1, shape_param2, nc_param), -r, N1 + r)
    End If
End Function

Public Function comp_cdf_BETA_os(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_cdf_BETA_os = [#VALUE!]: Exit Function
    Dim p As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    p = cdf_BETA_nc(x, shape_param1, shape_param2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        comp_cdf_BETA_os = comp_cdf_BETA(p, N1 + r, -r)
    Else
        comp_cdf_BETA_os = cdf_BETA(comp_cdf_BETA_nc(x, shape_param1, shape_param2, nc_param), -r, N1 + r)
    End If
End Function

Public Function inv_BETA_os(ByVal p As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of cdf_BETA_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 ' accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then inv_BETA_os = [#VALUE!]: Exit Function
    Dim xp As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    xp = invbeta(N1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        inv_BETA_os = inv_BETA_nc(xp, shape_param1, shape_param2, nc_param)
    Else
        inv_BETA_os = comp_inv_BETA_nc(oneMinusxp, shape_param1, shape_param2, nc_param)
    End If
End Function

Public Function comp_inv_BETA_os(ByVal p As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of comp_cdf_BETA_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_inv_BETA_os = [#VALUE!]: Exit Function
    Dim xp As Double, N1 As Double: N1 = n + 1
    If r > 0 Then r = r - N1
    xp = invcompbeta(N1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        comp_inv_BETA_os = inv_BETA_nc(xp, shape_param1, shape_param2, nc_param)
    Else
        comp_inv_BETA_os = comp_inv_BETA_nc(oneMinusxp, shape_param1, shape_param2, nc_param)
    End If
End Function


Attribute VB_Name = "STAT_DIST_CONTINUOUS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : NORMAL_MASS_DIST_FUNC
'DESCRIPTION   : Normal mass distribution function (PDF)
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function NORMAL_MASS_DIST_FUNC(ByVal X_VAL As Double)

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979 'Atn(1) * 4

'------------------------------------------------------------
'NORMAL_MASS_DIST_FUNC = exp(-x^2/2)/sqrt(2*Pi)
'NORMAL_MASS_DIST_FUNC = 0.398942280401433 * Exp(-0.5 * x * x)
'NORMAL_MASS_DIST_FUNC = Exp(-0.5 * x * x - 0.918938533204673)
'------------------------------------------------------------

NORMAL_MASS_DIST_FUNC = 1 / (2 * PI_VAL) ^ 0.5 * Exp(-X_VAL ^ 2 / 2)
'1 / Sqr(2 * PI_VAL) * Exp(-X_VAL ^ 2 / 2)
    
Exit Function
ERROR_LABEL:
NORMAL_MASS_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NORMDIST_FUNC
'DESCRIPTION   : Returns the probability mass function of a normal distribution.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function NORMDIST_FUNC(ByVal X_VAL As Double, _
Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal VOLATILITY_VAL As Double = 1, _
Optional ByVal VERSION As Integer = 0)

Dim PI_VAL As Double
Dim Z_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
'----------------------------------------------------------------------
Select Case VERSION
'----------------------------------------------------------------------
Case 0
'----------------------------------------------------------------------
    NORMDIST_FUNC = _
        1 / (Sqr(2 * PI_VAL) * VOLATILITY_VAL) * Exp(-(((X_VAL - _
        MEAN_VAL) ^ 2) / (2 * VOLATILITY_VAL ^ 2)))
'----------------------------------------------------------------------
Case 1
'----------------------------------------------------------------------
    Z_VAL = (X_VAL - MEAN_VAL) / VOLATILITY_VAL
    NORMDIST_FUNC = pdf_normal(Z_VAL)
'----------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------
    NORMDIST_FUNC = WorksheetFunction.NormDist(X_VAL, _
                                MEAN_VAL, VOLATILITY_VAL, 0)
'----------------------------------------------------------------------
End Select
'----------------------------------------------------------------------

Exit Function
ERROR_LABEL:
NORMDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CUMUL_NORM_DIST
'DESCRIPTION   : Returns the cumulative normal distribution function
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function NORMSDIST_FUNC(ByVal X_VAL As Double, _
Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal VOLATILITY_VAL As Double = 1, _
Optional ByVal VERSION As Integer = 0)

'    Normal_Cumulative = Excel.Application.ChiDist(z ^ 2, 1) / 2
'    If z > 0 Then Normal_Cumulative = 1 - Normal_Cumulative

Dim Z_VAL As Double

On Error GoTo ERROR_LABEL

Z_VAL = (X_VAL - MEAN_VAL) / VOLATILITY_VAL

Select Case VERSION
Case 0
    NORMSDIST_FUNC = cdf_normal(Z_VAL)
Case Else
    NORMSDIST_FUNC = CDbl(MAPLE_CUMUL_NORMDIST_FUNC(Z_VAL))
    'WorksheetFunction.NormSDist(Z_VAL)
End Select

Exit Function
ERROR_LABEL:
NORMSDIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MAPLE_CUMUL_NORMDIST_FUNC
'DESCRIPTION   : Returns the best approximation of the cumulative normal
'distribution function
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MAPLE_CUMUL_NORMDIST_FUNC(ByVal X_VAL As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim A_VAL As Variant
Dim B_VAL As Variant
Dim C_VAL As Variant
Dim D_VAL As Variant
Dim E_VAL As Variant

Dim TEMP_INT As Variant
Dim TEMP_FRAC As Variant
Dim TEMP_BIN As Variant
Dim TEMP_EXP As Variant
Dim TEMP_LOCAL As Variant

Dim ATEMP_ABS As Variant
Dim BTEMP_ABS As Variant

Dim ATEMP_SUM As Variant
Dim BTEMP_SUM As Variant

Dim TEMP_MATRIX As Variant
Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant

Dim tolerance As Variant

On Error GoTo ERROR_LABEL

tolerance = 0.000000000000001
A_VAL = CDec(0.918938533204672) + CDec(0.74178) * CDec(tolerance)

ReDim ATEMP_ARR(0 To 13)

ATEMP_ARR(0) = CDec(1.2533141373155) + CDec(0.2512078826424 * tolerance)
ATEMP_ARR(1) = CDec(0.655679542418798) + CDec(0.4715438712307 * tolerance)
ATEMP_ARR(2) = CDec(0.421369229288054) + CDec(0.4732249343335 * tolerance)
ATEMP_ARR(3) = CDec(0.304590298710103) + CDec(0.2957336125465 * tolerance)
ATEMP_ARR(4) = CDec(0.23665238291356) + CDec(0.6706239859364 * tolerance)
ATEMP_ARR(5) = CDec(0.192808104715315) + CDec(0.7648774657279 * tolerance)
ATEMP_ARR(6) = CDec(0.162377660896867) + CDec(0.4618156821028 * tolerance)
ATEMP_ARR(7) = CDec(0.14010418345305) + CDec(0.2415995345218 * tolerance)
ATEMP_ARR(8) = CDec(0.123131963257932) + CDec(0.2962821807435 * tolerance)
ATEMP_ARR(9) = CDec(0.109787282578308) + CDec(0.2912306378331 * tolerance)
ATEMP_ARR(10) = CDec(0.099028596471731) + CDec(0.9213953371886 * tolerance)
ATEMP_ARR(11) = CDec(0.090175675501064) + CDec(0.6822797803562 * tolerance)
ATEMP_ARR(12) = CDec(0.082766286501369) + CDec(0.17725226505 * tolerance)
ATEMP_ARR(13) = CDec(0.076475761016248) + CDec(0.5029934951942 * tolerance)

ATEMP_ABS = CDec(X_VAL)

If (X_VAL + X_VAL < 0) Then: ATEMP_ABS = -CDec(ATEMP_ABS)

If (11 < X_VAL) Then
   MAPLE_CUMUL_NORMDIST_FUNC = 1
  Exit Function
End If

If (X_VAL < -11) Then
  MAPLE_CUMUL_NORMDIST_FUNC = 0
  Exit Function
End If

j = Int(ATEMP_ABS)
C_VAL = CDec(j)
B_VAL = CDec(ATEMP_ABS - C_VAL)

ReDim TEMP_MATRIX(0 To 32, 1 To 1)

TEMP_MATRIX(0, 1) = ATEMP_ARR(j)
TEMP_MATRIX(1, 1) = CDec(C_VAL * TEMP_MATRIX(0, 1) - CDec(1))
For i = 2 To 32
  TEMP_MATRIX(i, 1) = CDec((TEMP_MATRIX(i - 2, 1) + C_VAL * _
                TEMP_MATRIX(i - 1, 1)) / CDec(i))
Next i

ATEMP_SUM = CDec(0)
For i = 0 To 32
  ATEMP_SUM = CDec(TEMP_MATRIX(32 - i, 1) + B_VAL * ATEMP_SUM)
Next i

D_VAL = (-CDec(0.5) * ATEMP_ABS * ATEMP_ABS - A_VAL)

If (0 < D_VAL) Then
  TEMP_LOCAL = Exp(D_VAL)
  GoTo 1983
End If
If (D_VAL <= -64) Then
  TEMP_LOCAL = Exp(D_VAL)
  GoTo 1983
End If

ReDim BTEMP_ARR(0 To 5)

BTEMP_ARR(0) = CDec(0.367879441171442) + CDec(0.3215955237702 * tolerance)
BTEMP_ARR(1) = CDec(0.135335283236612) + CDec(0.691893999495 * tolerance)
BTEMP_ARR(2) = CDec(0.018315638888734) + CDec(0.1802937180213 * tolerance)
BTEMP_ARR(3) = CDec(0.000335462627902) + CDec(0.5118388213891 * tolerance)
BTEMP_ARR(4) = CDec(0.000000112535174) + CDec(0.7192591145138 * tolerance)
BTEMP_ARR(5) = CDec(0.000000000000012) + CDec(0.6641655490942 * tolerance)

BTEMP_ABS = -CDec(D_VAL)
TEMP_INT = Int(BTEMP_ABS)
TEMP_FRAC = CDec(BTEMP_ABS - TEMP_INT)

E_VAL = TEMP_INT
k = 32
TEMP_EXP = CDec(1)
For i = 5 To 0 Step -1
  TEMP_BIN = Int(E_VAL / k)
  If (TEMP_BIN = 1) Then
    TEMP_EXP = TEMP_EXP * CDec(TEMP_BIN) * BTEMP_ARR(i)
  End If
  E_VAL = E_VAL - TEMP_BIN * k
  k = Int(k / 2)
Next i

ReDim TEMP_MATRIX(0 To 32, 1 To 1)

TEMP_MATRIX(0, 1) = CDec(1)
For i = 1 To 32
  TEMP_MATRIX(i, 1) = -TEMP_FRAC * CDec(TEMP_MATRIX(i - 1, 1)) / CDec(i)
Next i

BTEMP_SUM = CDec(0)
For i = 0 To 32
  BTEMP_SUM = CDec(TEMP_MATRIX(32 - i, 1) + BTEMP_SUM)
Next i

TEMP_LOCAL = CStr(TEMP_EXP * BTEMP_SUM)

1983:

ATEMP_SUM = 1 - ATEMP_SUM * CDec(TEMP_LOCAL)
If (X_VAL < 0#) Then
  ATEMP_SUM = CDec(CDec(1#) - ATEMP_SUM)
End If

MAPLE_CUMUL_NORMDIST_FUNC = CStr(ATEMP_SUM)

Exit Function
ERROR_LABEL:
MAPLE_CUMUL_NORMDIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NORMSINV_FUNC
'DESCRIPTION   : Returns the inverse of the normal cumulative distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function NORMSINV_FUNC(ByVal PROBABILITY_VAL As Double, _
Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal VOLATILITY_VAL As Double = 1, _
Optional ByVal VERSION As Integer = 1)
    
On Error GoTo ERROR_LABEL

Select Case VERSION
'---------------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------------
    NORMSINV_FUNC = inv_normal(PROBABILITY_VAL) * VOLATILITY_VAL + MEAN_VAL
               'inv_normal --> FROM THE ALGORITHMS_DISTS SUB
'---------------------------------------------------------------------------------------
Case 1 'Calculates the Normal Standard numbers given x, the associated
'uniform number (0, 1); VB version of the Moro's (1995) code in C
'---------------------------------------------------------------------------------------
    NORMSINV_FUNC = CNDEV_FUNC(PROBABILITY_VAL) * VOLATILITY_VAL + MEAN_VAL
'---------------------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------------------
    NORMSINV_FUNC = WorksheetFunction.NormSInv(PROBABILITY_VAL) * _
                    VOLATILITY_VAL + MEAN_VAL
    'Also EQUAL = WorksheetFunction.NormInv(PROBABILITY_VAL, MEAN_VAL,
    'VOLATILITY_VAL, 0)
End Select


Exit Function
ERROR_LABEL:
NORMSINV_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CNDEV_FUNC

'DESCRIPTION   : Calculates the Normal Standard numbers given x, the associated
'uniform number (0, 1); VB version of the Moro's (1995) code in C

'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CNDEV_FUNC(ByVal U_VAL As Double) 'As Double
    
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim X_VAL As Double
Dim R_VAL As Double
Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim CTEMP_ARR As Variant

On Error GoTo ERROR_LABEL

ATEMP_ARR = Array(2.50662823884, -18.61500062529, _
    41.39119773534, -25.44106049637)
ii = LBound(ATEMP_ARR)
    
BTEMP_ARR = Array(-8.4735109309, 23.08336743743, _
    -21.06224101826, 3.13082909833)
jj = LBound(BTEMP_ARR)
    
CTEMP_ARR = Array(0.337475482272615, 0.976169019091719, _
    0.160797971491821, 2.76438810333863E-02, _
    3.8405729373609E-03, 3.951896511919E-04, _
    3.21767881767818E-05, 2.888167364E-07, 3.960315187E-07)
kk = LBound(CTEMP_ARR)

X_VAL = U_VAL - 0.5
If Abs(X_VAL) < 0.42 Then
    R_VAL = X_VAL * X_VAL
    R_VAL = X_VAL * (((ATEMP_ARR(ii + 3) * R_VAL + ATEMP_ARR(ii + 2)) * _
    R_VAL + ATEMP_ARR(ii + 1)) * R_VAL + ATEMP_ARR(ii + 0)) _
    / ((((BTEMP_ARR(jj + 3) * R_VAL + BTEMP_ARR(jj + 2)) * R_VAL + _
    BTEMP_ARR(jj + 1)) * R_VAL + BTEMP_ARR(jj + 0)) * R_VAL + 1)
    CNDEV_FUNC = R_VAL
    Exit Function
End If
R_VAL = U_VAL
If X_VAL >= 0 Then R_VAL = 1 - U_VAL
R_VAL = Log(-Log(R_VAL))

R_VAL = CTEMP_ARR(kk + 0) + R_VAL * (CTEMP_ARR(kk + 1) + R_VAL * _
    (CTEMP_ARR(kk + 2) + R_VAL * (CTEMP_ARR(kk + 3) + R_VAL * (CTEMP_ARR(kk + 4) + _
    R_VAL * (CTEMP_ARR(kk + 5) + R_VAL * (CTEMP_ARR(kk + 6) + R_VAL * _
    (CTEMP_ARR(kk + 7) + R_VAL * CTEMP_ARR(kk + 8))))))))
If X_VAL < 0 Then R_VAL = -R_VAL
CNDEV_FUNC = R_VAL
    
Exit Function
ERROR_LABEL:
CNDEV_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CND_FUNC
'DESCRIPTION   : Cumulative normal distribution approximations
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function CND_FUNC(ByVal X_VAL As Double, _
Optional ByVal VERSION As Integer = 0)

On Error GoTo ERROR_LABEL

Select Case VERSION
    Case 0
        
' Marsaglia solutions are given by Marsaglia and series are stopped if
' adding more terms does no longer change the summation, while for arguements
' beyond 7 an asymptotics from Abramowitz & Stegun is used (their series can
' be shown to have a closed form). In cdfN_Marsaglia functions are 'Taylored'
' around 0, ... , 7 and should finish after 20 steps with Excel's exactness.
' In cdfN_Marsaglia_0 the series is given in Zero, so it might need up to
' 140 steps, the tail error may be estimated through the Gamma distribution
' as 1-GAMMA(-1/2+n,1/2*x^2)/GAMMA(-1/2+n) for n steps.
        
        CND_FUNC = UNIVAR_CUMUL_NORM_FUNC("Marsaglia", X_VAL)
    Case 1
        
' Taylor series for the cumulative normal around various integers cut
' off at 7.1 < abs(x) and should give 15 digits, due to George Marsaglia
        
        CND_FUNC = UNIVAR_CUMUL_NORM_FUNC("Marsaglia_0", X_VAL)
    Case 2
'////////////Numerical Computation of Rectangular Bivariate\\\\\\\\\\\\
        CND_FUNC = UNIVAR_CUMUL_NORM_FUNC("hart", X_VAL)
    
    Case 3
        CND_FUNC = UNIVAR_CUMUL_NORM_FUNC("ab & steg", X_VAL)
    Case 4
        CND_FUNC = UNIVAR_CUMUL_NORM_FUNC("ab & steg fix", X_VAL)
    Case 5
        CND_FUNC = UNIVAR_CUMUL_NORM_FUNC("asymptotic", X_VAL)
    Case Else
    'If Abs(X_VAL) < 4 Then
        '  Use Hart
    'ElseIf Abs(X_VAL) < 7.4 Then
    '  Use George Marsaglia
    'Else
    '  Use Asymptotic
    'End If
        CND_FUNC = UNIVAR_CUMUL_NORM_FUNC("", X_VAL)
End Select

Exit Function
ERROR_LABEL:
CND_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CBND_FUNC
'DESCRIPTION   : Cumulative bivariate normal distribution approximations
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CBND_FUNC(ByVal A_VAL As Double, _
ByVal B_VAL As Double, _
ByVal RHO As Double, _
Optional ByVal UNI_VERSION As Integer = 0, _
Optional ByVal BI_VERSION As Integer = 0)

On Error GoTo ERROR_LABEL

Select Case BI_VERSION
'-------------------------------------------------------------------------------
    Case 0
        
' Alan Genz, A function for computing bivariate normal probabilities.
'
' This function is based on the method described by
'   Drezner, Z and G.O. Wesolowsky, (1989),
'   On the computation of the bivariate normal integral,
'   Journal of Statist. Comput. Simul. 35, pp. 101-107
'
' with major modifications for double precision, and for |R| close to 1
' done by Greame West.
        
        If UNI_VERSION = 0 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("genz", "Marsaglia", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 1 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("genz", "Marsaglia_0", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 2 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("genz", "hart", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 3 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("genz", "ab & steg", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 4 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("genz", "ab & steg fix", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 5 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("genz", "asymptotic", A_VAL, B_VAL, RHO)
        Else
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("genz", "", A_VAL, B_VAL, RHO)
        End If
'-------------------------------------------------------------------------------
    Case 1

' The method of Drezner & Wesolowsky given by Genz results in good speed
' and exactness up to 14 or 15 digits - based on a good choice for cdfN1:
' may be in Fortran it is not neccessary to modify the orignal CND, but at
' least Excel needs that.

        If UNI_VERSION = 0 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezner", "Marsaglia", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 1 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezner", "Marsaglia_0", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 2 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezner", "hart", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 3 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezner", "ab & steg", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 4 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezner", "ab & steg fix", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 5 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezner", "asymptotic", A_VAL, B_VAL, RHO)
        Else
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezner", "", A_VAL, B_VAL, RHO)
        End If
'-------------------------------------------------------------------------------
    Case 2
        If UNI_VERSION = 0 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezner", "Marsaglia", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 1 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezner", "Marsaglia_0", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 2 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezner", "hart", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 3 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezner", "ab & steg", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 4 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezner", "ab & steg fix", A_VAL, B_VAL, _
                RHO)
        ElseIf UNI_VERSION = 5 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezner", "asymptotic", A_VAL, B_VAL, RHO)
        Else
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezner", "", A_VAL, B_VAL, RHO)
        End If
'-------------------------------------------------------------------------------
    
    Case 3
        If UNI_VERSION = 0 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes first", "Marsaglia", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 1 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes first", "Marsaglia_0", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 2 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes first", "hart", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 3 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes first", "ab & steg", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 4 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes first", "ab & steg fix", A_VAL, B_VAL, _
                RHO)
        ElseIf UNI_VERSION = 5 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes first", "asymptotic", A_VAL, B_VAL, RHO)
        Else
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes first", "", A_VAL, B_VAL, RHO)
        End If
    
'-------------------------------------------------------------------------------
    
    Case 4
        If UNI_VERSION = 0 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes second", "Marsaglia", A_VAL, B_VAL, _
                RHO)
        ElseIf UNI_VERSION = 1 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes second", "Marsaglia_0", A_VAL, B_VAL, _
                RHO)
        ElseIf UNI_VERSION = 2 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes second", "hart", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 3 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes second", "ab & steg", A_VAL, B_VAL, RHO)
        ElseIf UNI_VERSION = 4 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes second", "ab & steg fix", A_VAL, B_VAL, _
                RHO)
        ElseIf UNI_VERSION = 5 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes second", "asymptotic", A_VAL, B_VAL, _
                RHO)
        Else
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("drezwes second", "", A_VAL, B_VAL, _
                RHO)
        End If

'-------------------------------------------------------------------------------
    
    Case 5
        If UNI_VERSION = 0 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezwes second", "Marsaglia", A_VAL, _
                B_VAL, RHO)
        ElseIf UNI_VERSION = 1 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezwes second", "Marsaglia_0", A_VAL, _
                B_VAL, RHO)
        ElseIf UNI_VERSION = 2 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezwes second", "hart", A_VAL, _
                B_VAL, RHO)
        ElseIf UNI_VERSION = 3 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezwes second", "ab & steg", A_VAL, _
                B_VAL, RHO)
        ElseIf UNI_VERSION = 4 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezwes second", "ab & steg fix", A_VAL, _
                B_VAL, RHO)
        ElseIf UNI_VERSION = 5 Then
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezwes second", "asymptotic", A_VAL, _
                B_VAL, RHO)
        Else
            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("fixed drezwes second", "", A_VAL, _
                B_VAL, RHO)
        End If
'-------------------------------------------------------------------------------
    
    Case Else '6

' For lower correlation Marsaglia's method can be used to write
' down recursions for the Taylor series while for higher correlations it is
' better to use a method of Vasicek (which i modified a bit), both more or
' less are recursions for the incomplete Gamma function involved. For both
' the expansion is stopped by machine precision (and the cdfN1 used).

' Note that the series approach is limited by ~ 1e-15 for exactness: a
' decomposition in two (very exact) summands is
' done, but a system exactness 1 + eps <> 1 is involved.

            CBND_FUNC = BIVAR_CUMUL_NORM_FUNC("Marsaglia", "Marsaglia", A_VAL, B_VAL, RHO)
'    Case Else
'            CBND_FUNC = CDFN_2_INTEGRAL_FUNC(A_VAL, B_VAL, RHO)
End Select

Exit Function
ERROR_LABEL:
    CBND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CND_FUNC2
'DESCRIPTION   : Cummulative double precision algorithm based on Hart 1968
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CND_FUNC2(ByVal X_VAL As Variant) 'As Double
    
Dim Y_VAL As Double
Dim EXP_VAL As Double
Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double
    
On Error GoTo ERROR_LABEL

Y_VAL = Abs(X_VAL)
If Y_VAL > 37 Then
    CND_FUNC2 = 0
Else
    EXP_VAL = Exp(-Y_VAL ^ 2 / 2)
    If Y_VAL < 7.07106781186547 Then
        ATEMP_SUM = 3.52624965998911E-02 * Y_VAL + 0.700383064443688
        ATEMP_SUM = ATEMP_SUM * Y_VAL + 6.37396220353165
        ATEMP_SUM = ATEMP_SUM * Y_VAL + 33.912866078383
        ATEMP_SUM = ATEMP_SUM * Y_VAL + 112.079291497871
        ATEMP_SUM = ATEMP_SUM * Y_VAL + 221.213596169931
        ATEMP_SUM = ATEMP_SUM * Y_VAL + 220.206867912376
        BTEMP_SUM = 8.83883476483184E-02 * Y_VAL + 1.75566716318264
        BTEMP_SUM = BTEMP_SUM * Y_VAL + 16.064177579207
        BTEMP_SUM = BTEMP_SUM * Y_VAL + 86.7807322029461
        BTEMP_SUM = BTEMP_SUM * Y_VAL + 296.564248779674
        BTEMP_SUM = BTEMP_SUM * Y_VAL + 637.333633378831
        BTEMP_SUM = BTEMP_SUM * Y_VAL + 793.826512519948
        BTEMP_SUM = BTEMP_SUM * Y_VAL + 440.413735824752
        
        CND_FUNC2 = EXP_VAL * ATEMP_SUM / BTEMP_SUM
    Else
        ATEMP_SUM = Y_VAL + 0.65
        ATEMP_SUM = Y_VAL + 4 / ATEMP_SUM
        ATEMP_SUM = Y_VAL + 3 / ATEMP_SUM
        ATEMP_SUM = Y_VAL + 2 / ATEMP_SUM
        ATEMP_SUM = Y_VAL + 1 / ATEMP_SUM
        
        CND_FUNC2 = EXP_VAL / (ATEMP_SUM * 2.506628274631)
    End If
End If
  
If X_VAL > 0 Then CND_FUNC2 = 1 - CND_FUNC2

Exit Function
ERROR_LABEL:
CND_FUNC2 = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CBND_FUNC2
'DESCRIPTION   : The cumulative bivariate normal distribution function
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CBND_FUNC2(ByVal X_VAL As Double, _
ByVal Y_VAL As Double, _
ByVal RHO_VAL As Double) 'As Double

'A_VAL function for computing bivariate normal probabilities.
'Alan Genz
'Department of Mathematics
'Washington State University
'Pullman, WA 99164-3113
'Email : alangenz@wsu.edu

'This function is based on the method described by
'Drezner, Z and G.O. Wesolowsky, (1990),
'On the computation of the bivariate normal integral,
'Journal of Statist. Comput. Simul. 35, pp. 101-107,
'with major modifications for double precision, and for |R| close to 1.
'This code was originally transelated into VBA by Graeme West

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer

Dim H_VAL As Double
Dim K_VAL As Double
Dim HK_VAL As Double
Dim HS_VAL As Double
Dim BVN_VAL As Double
Dim ASS_VAL As Double
Dim ASR_VAL As Double
Dim SN_VAL As Double
Dim A_VAL As Double
Dim B_VAL As Double
Dim BS_VAL As Double
Dim C_VAL As Double
Dim D_VAL As Double
Dim XS_VAL As Double
Dim RS_VAL As Double

Dim PI_VAL As Double

Dim XX_VECTOR(1 To 10, 1 To 3) As Double
Dim W_VECTOR(1 To 10, 1 To 3) As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

W_VECTOR(1, 1) = 0.17132449237917
XX_VECTOR(1, 1) = -0.932469514203152
W_VECTOR(2, 1) = 0.360761573048138
XX_VECTOR(2, 1) = -0.661209386466265
W_VECTOR(3, 1) = 0.46791393457269
XX_VECTOR(3, 1) = -0.238619186083197

W_VECTOR(1, 2) = 4.71753363865118E-02
XX_VECTOR(1, 2) = -0.981560634246719
W_VECTOR(2, 2) = 0.106939325995318
XX_VECTOR(2, 2) = -0.904117256370475
W_VECTOR(3, 2) = 0.160078328543346
XX_VECTOR(3, 2) = -0.769902674194305
W_VECTOR(4, 2) = 0.203167426723066
XX_VECTOR(4, 2) = -0.587317954286617
W_VECTOR(5, 2) = 0.233492536538355
XX_VECTOR(5, 2) = -0.36783149899818
W_VECTOR(6, 2) = 0.249147045813403
XX_VECTOR(6, 2) = -0.125233408511469

W_VECTOR(1, 3) = 1.76140071391521E-02
XX_VECTOR(1, 3) = -0.993128599185095
W_VECTOR(2, 3) = 4.06014298003869E-02
XX_VECTOR(2, 3) = -0.963971927277914
W_VECTOR(3, 3) = 6.26720483341091E-02
XX_VECTOR(3, 3) = -0.912234428251326
W_VECTOR(4, 3) = 8.32767415767048E-02
XX_VECTOR(4, 3) = -0.839116971822219
W_VECTOR(5, 3) = 0.10193011981724
XX_VECTOR(5, 3) = -0.746331906460151
W_VECTOR(6, 3) = 0.118194531961518
XX_VECTOR(6, 3) = -0.636053680726515
W_VECTOR(7, 3) = 0.131688638449177
XX_VECTOR(7, 3) = -0.510867001950827
W_VECTOR(8, 3) = 0.142096109318382
XX_VECTOR(8, 3) = -0.37370608871542
W_VECTOR(9, 3) = 0.149172986472604
XX_VECTOR(9, 3) = -0.227785851141645
W_VECTOR(10, 3) = 0.152753387130726
XX_VECTOR(10, 3) = -7.65265211334973E-02
      
If Abs(RHO_VAL) < 0.3 Then
  l = 1
  k = 3
ElseIf Abs(RHO_VAL) < 0.75 Then
  l = 2
  k = 6
Else
  l = 3
  k = 10
End If
      
H_VAL = -X_VAL
K_VAL = -Y_VAL
HK_VAL = H_VAL * K_VAL
BVN_VAL = 0
      
If Abs(RHO_VAL) < 0.925 Then
  If Abs(RHO_VAL) > 0 Then
    HS_VAL = (H_VAL * H_VAL + K_VAL * K_VAL) / 2
    ASR_VAL = ATN_FUNC(RHO_VAL)
    For i = 1 To k
      For j = -1 To 1 Step 2
        SN_VAL = Sin(ASR_VAL * (j * XX_VECTOR(i, l) + 1) / 2)
        BVN_VAL = BVN_VAL + W_VECTOR(i, l) * Exp((SN_VAL * HK_VAL - HS_VAL) / (1 - SN_VAL * SN_VAL))
      Next j
    Next i
    BVN_VAL = BVN_VAL * ASR_VAL / (4 * PI_VAL)
  End If
  BVN_VAL = BVN_VAL + CND_FUNC2(-H_VAL) * CND_FUNC2(-K_VAL)
Else
  If RHO_VAL < 0 Then
    K_VAL = -K_VAL
    HK_VAL = -HK_VAL
  End If
  If Abs(RHO_VAL) < 1 Then
    ASS_VAL = (1 - RHO_VAL) * (1 + RHO_VAL)
    A_VAL = Sqr(ASS_VAL)
    BS_VAL = (H_VAL - K_VAL) ^ 2
    C_VAL = (4 - HK_VAL) / 8
    D_VAL = (12 - HK_VAL) / 16
    ASR_VAL = -(BS_VAL / ASS_VAL + HK_VAL) / 2
    If ASR_VAL > -100 Then BVN_VAL = A_VAL * Exp(ASR_VAL) * (1 - C_VAL * _
    (BS_VAL - ASS_VAL) * (1 - D_VAL * BS_VAL / 5) / 3 + C_VAL * _
    D_VAL * ASS_VAL * ASS_VAL / 5)
    If -HK_VAL < 100 Then
      B_VAL = Sqr(BS_VAL)
      BVN_VAL = BVN_VAL - Exp(-HK_VAL / 2) * Sqr(2 * PI_VAL) * CND_FUNC2(-B_VAL / _
      A_VAL) * B_VAL * (1 - C_VAL * BS_VAL * (1 - D_VAL * BS_VAL / 5) / 3)
    End If
    A_VAL = A_VAL / 2
    For i = 1 To k
      For j = -1 To 1 Step 2
        XS_VAL = (A_VAL * (j * XX_VECTOR(i, l) + 1)) ^ 2
        RS_VAL = Sqr(1 - XS_VAL)
        ASR_VAL = -(BS_VAL / XS_VAL + HK_VAL) / 2
        If ASR_VAL > -100 Then
           BVN_VAL = BVN_VAL + A_VAL * W_VECTOR(i, l) * Exp(ASR_VAL) * (Exp(-HK_VAL _
           * (1 - RS_VAL) / (2 * (1 + RS_VAL))) / RS_VAL - (1 + C_VAL * XS_VAL * _
           (1 + D_VAL * XS_VAL)))
        End If
      Next j
    Next i
    BVN_VAL = -BVN_VAL / (2 * PI_VAL)
  End If
  If RHO_VAL > 0 Then
    BVN_VAL = BVN_VAL + CND_FUNC2(-MAXIMUM_FUNC(H_VAL, K_VAL))
  Else
    BVN_VAL = -BVN_VAL
    If K_VAL > H_VAL Then BVN_VAL = BVN_VAL + CND_FUNC2(K_VAL) - CND_FUNC2(H_VAL)
  End If
End If
CBND_FUNC2 = BVN_VAL

Exit Function
ERROR_LABEL:
CBND_FUNC2 = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CTND_FUNC
'DESCRIPTION   : Cumulative trivariate normal distribution approximations
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CTND_FUNC(ByVal A_VAL As Double, _
ByVal B_VAL As Double, _
ByVal C_VAL As Double, _
ByVal RHO_AB As Double, _
ByVal RHO_CA As Double, _
ByVal RHO_BC As Double, _
Optional ByVal VERSION As Integer = 0)

' For trivariate only the result through integration is upported. Note that this
' does not say that is gives 'exact' results:

' Tests using Maple do show errors if the correlation matrix becomes almost
' numerical singular. But one can merely have excellent and fast result results
' with such limited exactness as Excel and a usual 15 digit environment gives
' for such a problem

On Error GoTo ERROR_LABEL
'    Select Case VERSION
'        Case 0
            CTND_FUNC = STECK_TRIVAR_CUMUL_NORM_FUNC(A_VAL, B_VAL, C_VAL, RHO_AB, RHO_CA, RHO_BC)
'        Case Else '1
'            CTND_FUNC = CDFN_3_INTEGRAL_FUNC(A_VAL, B_VAL, C_VAL, RHO_AB, RHO_CA, RHO_BC, 0)
 '       Case Else
'            CTND_FUNC = CDFN_3_INTEGRAL_FUNC(A_VAL, B_VAL, C_VAL, RHO_AB, RHO_CA, RHO_BC, 1)
'    End Select

Exit Function
ERROR_LABEL:
CTND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LOGNORMDIST_FUNC
'DESCRIPTION   : Returns the standard log normal cumulative distribution function
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function LOGNORMDIST_FUNC(ByVal X_VAL As Double, _
Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal VOLATILITY_VAL As Double = 1, _
Optional ByVal MU_SD_TYPE As Boolean = False, _
Optional ByVal VERSION As Integer = 0)

On Error GoTo ERROR_LABEL

Select Case VERSION
Case 0
    If MU_SD_TYPE = False Then
    
    'In this case mean is the mean of ln(x), and Standard_dev is the standard
    'deviation of ln(x).
    
         LOGNORMDIST_FUNC = cdf_normal((Log(X_VAL + 1) - MEAN_VAL) / VOLATILITY_VAL)
        'Same as =NORMSDIST((LN(TEMP_VALUE+1)-MEAN_VAL)/VOLATILITY_VAL)
    
    Else
    
    'Suppose we have a Log-normal set {y} and we know its
    'Mean and Standard Deviation. Hence, we need to
    'find values so that we can identify the associated Normal set {x}.
    
        VOLATILITY_VAL = (Log(VOLATILITY_VAL ^ 2 / MEAN_VAL ^ 2 + 1) / Log(10#)) ^ 0.5
        MEAN_VAL = (Log(MEAN_VAL) / Log(10#)) - VOLATILITY_VAL ^ 2 / 2
    
        LOGNORMDIST_FUNC = cdf_normal((Log(X_VAL) - MEAN_VAL) _
         / VOLATILITY_VAL)
    
    End If

Case Else
    If MU_SD_TYPE = False Then
    
        LOGNORMDIST_FUNC = WorksheetFunction.LogNormDist(X_VAL + 1, _
        MEAN_VAL, VOLATILITY_VAL)
    
    Else

        VOLATILITY_VAL = (Log(VOLATILITY_VAL ^ 2 / MEAN_VAL ^ 2 + 1) / Log(10#)) ^ 0.5
        MEAN_VAL = (Log(MEAN_VAL) / Log(10#)) - VOLATILITY_VAL ^ 2 / 2
    
        LOGNORMDIST_FUNC = WorksheetFunction.LogNormDist(X_VAL, MEAN_VAL, VOLATILITY_VAL)
    End If
End Select

Exit Function
ERROR_LABEL:
LOGNORMDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_LOGNORMDIST_FUNC
'DESCRIPTION   : Returns the inverse of the lognormal cumulative distribution
'function of x, where ln(x) is normally distributed with parameters mean
'and standard_dev.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function INVERSE_LOGNORMDIST_FUNC(ByVal PROBABILITY_VAL As Double, _
Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal VOLATILITY_VAL As Double = 1, _
Optional ByVal MU_SD_TYPE As Boolean = False, _
Optional ByVal VERSION As Integer = 0)

On Error GoTo ERROR_LABEL

Select Case VERSION
Case 0
    If MU_SD_TYPE = False Then
    
    'In this case mean is the mean of ln(x), and Standard_dev is the standard
    'deviation of ln(x).
    
         INVERSE_LOGNORMDIST_FUNC = Exp(inv_normal(PROBABILITY_VAL) * VOLATILITY_VAL + MEAN_VAL) - 1
            'Same as =Exp(NORMINV(PROBABILITY_VAL,MEAN_VAL,VOLATILITY_VAL)) -1
    
    Else
    
    'Suppose we have a Log-normal set {y} and we know its
    'Mean and Standard Deviation. Hence, we need to
    'find values so that we can identify the associated Normal set {x}.
    
        VOLATILITY_VAL = (Log(VOLATILITY_VAL ^ 2 / MEAN_VAL ^ 2 + 1) / Log(10#)) ^ 0.5
        MEAN_VAL = (Log(MEAN_VAL) / Log(10#)) - VOLATILITY_VAL ^ 2 / 2
    
        INVERSE_LOGNORMDIST_FUNC = Exp(inv_normal(PROBABILITY_VAL) * VOLATILITY_VAL + MEAN_VAL)
    
    End If
    
Case Else
    If MU_SD_TYPE = False Then
        INVERSE_LOGNORMDIST_FUNC = WorksheetFunction.LogInv(PROBABILITY_VAL, MEAN_VAL, VOLATILITY_VAL) - 1
    Else
        VOLATILITY_VAL = (Log(VOLATILITY_VAL ^ 2 / MEAN_VAL ^ 2 + 1) / Log(10#)) ^ 0.5
        MEAN_VAL = (Log(MEAN_VAL) / Log(10#)) - VOLATILITY_VAL ^ 2 / 2
        INVERSE_LOGNORMDIST_FUNC = WorksheetFunction.LogInv(PROBABILITY_VAL, MEAN_VAL, VOLATILITY_VAL)
    End If
End Select

Exit Function
ERROR_LABEL:
INVERSE_LOGNORMDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GAMMA_DIST_FUNC
'DESCRIPTION   : Returns the gamma distribution. You can use this function
'to study variables that may have a skewed distribution. The gamma
'distribution is commonly used in asynchronous analysis.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function GAMMA_DIST_FUNC(ByVal X_VAL As Double, _
ByVal ALPHA As Double, _
ByVal beta As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'X_VAL: is the value at which you want to evaluate the distribution.

'ALPHA: is a shape parameter to the distribution.

'BETA: is a scale parameter to the distribution.
'If BETA = 1, GAMMADIST returns the standard gamma distribution.

'CUMUL_FLAG: is a logical value that determines the form of the function. If
'cumulative is TRUE, GAMMA_DIST returns the cumulative distribution function;
'if FALSE, it returns the probability density function.

'COMP_FLAG: 1-cumulative distribution function


'The Gamma distribution is most often used to describe the distribution of
'the amount of time until the nth occurrence of an event in a Poisson process.
'For example, customer service or machine repair.  The Gamma distribution is
'related to many other distributions.  For example, when a Gamma distribution
'has an alpha of 1, Gamma(1, b), it becomes an Exponential distribution with
'scale parameter of b, Expo(b).  And a Chi-Square distribution with k df is the
'same as the Gamma(k/2, 2) distribution. The skew levels decreases as the scale
'parameter, b, increases.  All three means approximate the product of a and b.

On Error GoTo ERROR_LABEL

    Select Case CUMUL_FLAG
        
        Case True
        
            If COMP_FLAG = True Then
                GAMMA_DIST_FUNC = cdf_gamma(X_VAL, ALPHA, beta)
                'SAME as GAMMADIST(X_VAL, ALPHA, BETA, TRUE)
            ElseIf COMP_FLAG = False Then
                GAMMA_DIST_FUNC = comp_cdf_gamma(X_VAL, ALPHA, beta)
                'SAME as 1 - GAMMADIST(X_VAL, ALPHA, BETA, TRUE)
            End If

        Case False 'probability density function
        
        GAMMA_DIST_FUNC = pdf_gamma(X_VAL, ALPHA, beta)
        'SAME as GAMMADIST(X_VAL, ALPHA, BETA, FALSE)
    
    End Select

Exit Function
ERROR_LABEL:
GAMMA_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_GAMMA_DIST_FUNC
'DESCRIPTION   : Returns the inverse of the gamma cumulative distribution. If
'p = GAMMADIST(x,...), then GAMMAINV(p,...) = x. You can use this
'function to study a variable whose distribution may be skewed.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function INVERSE_GAMMA_DIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal ALPHA As Double, _
ByVal beta As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the gamma distribution.

'ALPHA: is a shape parameter to the distribution.

'BETA: is a scale parameter to the distribution.
'If BETA = 1, GAMMADIST returns the standard gamma distribution.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
Case True
    INVERSE_GAMMA_DIST_FUNC = inv_gamma(PROBABILITY_VAL, ALPHA, beta)
    'SAME as GAMMAINV(PROBABILITY_VAL, ALPHA, BETA)
Case False
    INVERSE_GAMMA_DIST_FUNC = comp_inv_gamma(PROBABILITY_VAL, ALPHA, beta)
    'SAME as GAMMAINV(1 - PROBABILITY_VAL, ALPHA, BETA, FALSE)
End Select

Exit Function
ERROR_LABEL:
INVERSE_GAMMA_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BETA_DIST
'DESCRIPTION   : Returns the BETA distribution. The BETA distribution is
'commonly used to study variation in the percentage of something across
'samples, such as the fraction of the day people spend watching television.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function BETA_DIST_FUNC(ByVal X_VAL As Double, _
ByVal ALPHA As Double, _
ByVal beta As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'X_VAL: is the value at which you want to evaluate the distribution.

'Alpha: is a shape parameter to the distribution.

'Beta: is a scale parameter to the distribution.
'If BETA = 1, BETADIST returns the standard BETA distribution.

'CUMUL_FLAG: is a logical value that determines the form of the function. If
'cumulative is TRUE, BETA_DIST returns the cumulative distribution function;
'if FALSE, it returns the probability density function.

'COMP_FLAG: 1-cumulative distribution function


'The Beta distribution can be used in the absence of data.  Possible applications
'are estimate the proportion of defective items in a shipment or time to complete
'a task.  The Beta distribution has two shape parameters, Alpha and Beta.  When
'the two parameters are equal, the distribution is symmetrical.  For example,
'when both ALPHA and Beta are equal to one, the distribution becomes uniform.
'If ALPHA is less than Beta, the distribution is skewed to the left.  And if ALPHA
'is greater than Beta, the distribution is skewed to the right.

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        
        If COMP_FLAG = True Then
            BETA_DIST_FUNC = cdf_BETA(X_VAL, ALPHA, beta)
        'SAME as BETADIST(X_VAL, ALPHA, BETA, TRUE)
        ElseIf COMP_FLAG = False Then
            BETA_DIST_FUNC = comp_cdf_BETA(X_VAL, ALPHA, beta)
        'SAME as 1 - BETADIST(X_VAL, ALPHA, BETA, TRUE)
        End If
    Case False 'probability density function
        BETA_DIST_FUNC = pdf_BETA(X_VAL, ALPHA, beta)
        'SAME as BETADIST(X_VAL, ALPHA, BETA, FALSE)
End Select

Exit Function
ERROR_LABEL:
BETA_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_BETA_DIST_FUNC
'DESCRIPTION   : Returns the inverse of the cumulative distribution function for
'a specified BETA distribution. That is, if probability = BETADIST(x,...),
'then BETAINV(probability,...) = x. The BETA distribution can be used
'in project planning to model probable completion times given an expected
'completion time and variability.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function INVERSE_BETA_DIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal ALPHA As Double, _
ByVal beta As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the BETA distribution.

'Alpha: is a shape parameter to the distribution.

'Beta: is a scale parameter to the distribution.
'If BETA = 1, BETADIST returns the standard BETA distribution.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        INVERSE_BETA_DIST_FUNC = inv_BETA(PROBABILITY_VAL, ALPHA, beta)
        'SAME as BETAINV(PROBABILITY_VAL, ALPHA, BETA)
    Case False
        INVERSE_BETA_DIST_FUNC = comp_inv_BETA(PROBABILITY_VAL, ALPHA, beta)
        'SAME as BETAINV(1 - PROBABILITY_VAL, ALPHA, BETA, FALSE)
End Select

Exit Function
ERROR_LABEL:
INVERSE_BETA_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LIKELIHOOD_BETA_DIST_FUNC
'DESCRIPTION   : Likelihood Beta Dist
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function LIKELIHOOD_BETA_DIST_FUNC(ByVal X_VAL As Double, _
ByVal MEAN_VAL As Double, _
ByVal VOLATILITY_VAL As Double)

'If you create a senstivity table using MEAN this function will looks
'like a PDF, but it it not

'X_VAL: is the value between A and B at which to evaluate the function.

Dim LOWER_BOUND As Double 'is an lower bound to the interval of temp_value.
Dim UPPER_BOUND As Double 'is an upper bound to the interval of temp_value.
Dim TEMP_BETA As Double

On Error GoTo ERROR_LABEL

LOWER_BOUND = (MEAN_VAL ^ 2 - MEAN_VAL ^ 3) / VOLATILITY_VAL ^ 2 - MEAN_VAL
UPPER_BOUND = MEAN_VAL * (MEAN_VAL - 1) ^ 2 / VOLATILITY_VAL ^ 2 + MEAN_VAL - 1
TEMP_BETA = Exp(GAMMA_LN_FUNC(LOWER_BOUND) + GAMMA_LN_FUNC(UPPER_BOUND) - _
        GAMMA_LN_FUNC(LOWER_BOUND + UPPER_BOUND))

'MISSING THING: Derive the GammaLn from Stirling's series for
'ln(Gamma(x)), A046968/A046969

LIKELIHOOD_BETA_DIST_FUNC = X_VAL ^ (LOWER_BOUND - 1) * _
(1 - X_VAL) ^ (UPPER_BOUND - 1) / TEMP_BETA

Exit Function
ERROR_LABEL:
LIKELIHOOD_BETA_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LAPLACE_DIST_FUNC
'DESCRIPTION   : Bilateral exponential distribution: The Laplace distribution
'is often called the double-exponential distribution. This function is
'the signed analogue of the Exponential distribution function.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 019
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function LAPLACE_DIST_FUNC(ByVal X_VAL As Double, _
ByVal MEAN_VAL As Double, _
ByVal VOLATILITY_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 1, _
Optional ByVal CUMUL_FLAG As Boolean = True)

On Error GoTo ERROR_LABEL
    
VOLATILITY_VAL = VOLATILITY_VAL / Sqr(2)
    
Select Case CUMUL_FLAG
Case False 'Density
    LAPLACE_DIST_FUNC = Exp(-Abs(X_VAL - MEAN_VAL) / VOLATILITY_VAL) / (2 * VOLATILITY_VAL) * FACTOR_VAL
Case True 'Cumulative
    If X_VAL < MEAN_VAL Then
        LAPLACE_DIST_FUNC = 1 / 2 * Exp((X_VAL - MEAN_VAL) / VOLATILITY_VAL)
    Else
        LAPLACE_DIST_FUNC = 1 - 1 / 2 * Exp(-(X_VAL - MEAN_VAL) / VOLATILITY_VAL)
    End If
End Select
    
Exit Function
ERROR_LABEL:
LAPLACE_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVY_DIST_FUNC
'DESCRIPTION   : the Levy distribution (http://www.gummy-stuff.org/Levy.htm)
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 020
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function LEVY_DIST_FUNC(ByVal X_VAL As Double, _
ByVal VOLATILITY_VAL As Double, _
ByVal K_VAL As Double, _
Optional ByVal A_VAL As Double = 0)

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL
    
PI_VAL = 3.14159265358979
Select Case A_VAL
Case Is <> 0 'Levy PDF
    LEVY_DIST_FUNC = (K_VAL / (2 * PI_VAL)) ^ 0.5 * ((Exp(-0.5 * (K_VAL / X_VAL))) _
                    / (X_VAL) ^ (3 / 2))
Case Else 'Adj Levy PDF
    LEVY_DIST_FUNC = (1 / ((2 * PI_VAL) ^ 0.5 * K_VAL)) * ((Exp((1 / (2 * K_VAL ^ 2)) * _
                    (1 / (X_VAL ^ (2 * A_VAL - 2))))) / (X_VAL ^ A_VAL))
    '=(1/((2*PI_VAL)^0.5*K_VAL))*EXP(-0.5*(X_VAL/K_VAL)^2) --> NORMAL PDF
End Select

Exit Function
ERROR_LABEL:
LEVY_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXPONENTIAL_DIST_FUNC
'DESCRIPTION   : Returns the exponential distribution. Use EXPONDIST to model
'the time between events, such as how long an automated bank teller
'takes to deliver cash. For example, you can use EXPONDIST to determine
'the probability that the process takes at most 1 minute.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 021
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function EXPONENTIAL_DIST_FUNC(ByVal X_VAL As Double, _
ByVal LAMBDA As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'X_VAL: is the value at which you want to evaluate the distribution.

'LAMBDA: is a shape parameter to the distribution.

'CUMUL_FLAG: is a logical value that determines the form of the function. If
'cumulative is TRUE, EXPONENTIAL_DIST_FUNC returns the cumulative distribution function;
'if FALSE, it returns the probability density function.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
Case True
    If COMP_FLAG = True Then
        EXPONENTIAL_DIST_FUNC = cdf_exponential(X_VAL, LAMBDA)
    'SAME as EXPONDIST(X_VAL, LAMBDA, TRUE)
    ElseIf COMP_FLAG = False Then
        EXPONENTIAL_DIST_FUNC = comp_cdf_exponential(X_VAL, LAMBDA)
    'SAME as 1 - EXPONDIST(X_VAL, LAMBDA, TRUE)
    End If
Case False 'probability density function
    EXPONENTIAL_DIST_FUNC = pdf_exponential(X_VAL, LAMBDA)
    'SAME as EXPONDIST(X_VAL, LAMBDA, FALSE)
End Select

Exit Function
ERROR_LABEL:
EXPONENTIAL_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_EXPONENTIAL_DIST_FUNC
'DESCRIPTION   : Returns the inverse of the cumulative distribution function for
'a specified exponential distribution.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 022
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function INVERSE_EXPONENTIAL_DIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal LAMBDA As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the EXPONENTIAL distribution.

'LAMBDA: is a shape parameter to the distribution.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        INVERSE_EXPONENTIAL_DIST_FUNC = inv_exponential(PROBABILITY_VAL, LAMBDA)
    Case False
        INVERSE_EXPONENTIAL_DIST_FUNC = comp_inv_exponential(PROBABILITY_VAL, LAMBDA)
End Select

Exit Function
ERROR_LABEL:
INVERSE_EXPONENTIAL_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CHI_SQUARED_DIST_FUNC
'DESCRIPTION   : Returns the one-tailed probability of the chi-squared
'distribution. The c2 distribution is associated with a c2 test.
'Use the c2 test to compare observed and expected values. For example,
'a genetic experiment might hypothesize that the next generation of
'plants will exhibit a certain set of colors. By comparing the observed
'results with the expected ones, you can decide whether your original
'hypothesis is valid.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 023
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CHI_SQUARED_DIST_FUNC(ByVal X_VAL As Double, _
ByVal DEG_FREEDOM As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'X_VAL: is the value at which you want to evaluate the distribution.

'DEG_FREEDOM: is the number of degrees of freedom.

'CUMUL_FLAG: is a logical value that determines the form of the function. If
'cumulative is TRUE, CHI_SQUARED_DIST_FUNC returns the cumulative distribution function;
'if FALSE, it returns the probability density function.

'COMP_FLAG: 1-cumulative distribution function


'The most common use of the chi-square distribution is to test the difference
'between proportions.  It has a positive skew.  The skew decreases when degree
'of freedom increases as the distribution approaches normal.  The mean of a
'chi-square distribution is its degree of freedom.

'The output shows the estimate of skewness, mean, stand deviation, maximum value,
'minimum value, lower confidence interval, and upper confidence interval from
'each of the 3 simulations .  Each of the 3 means are very closed to its
'corresponding degree of freedom.

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        If COMP_FLAG = True Then
            CHI_SQUARED_DIST_FUNC = cdf_chi_sq(X_VAL, DEG_FREEDOM)
        ElseIf COMP_FLAG = False Then
            CHI_SQUARED_DIST_FUNC = comp_cdf_chi_sq(X_VAL, DEG_FREEDOM)
        End If
    Case False 'probability density function
        CHI_SQUARED_DIST_FUNC = pdf_chi_sq(X_VAL, DEG_FREEDOM)
End Select

Exit Function
ERROR_LABEL:
CHI_SQUARED_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_CHI_SQUARED_DIST_FUNC
'DESCRIPTION   : Returns the inverse of the one-tailed probability of the
'chi-squared distribution. If probability = CHIDIST(x,...), then
'CHIINV(probability,...) = x. Use this function to compare observed
'results with expected ones in order to decide whether your original
'hypothesis is valid.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 024
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function INVERSE_CHI_SQUARED_DIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal DEG_FREEDOM As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the CHI_SQ distribution.

'DEG_FREEDOM: is the number of degrees of freedom.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        INVERSE_CHI_SQUARED_DIST_FUNC = inv_chi_sq(PROBABILITY_VAL, DEG_FREEDOM)
    Case False
        INVERSE_CHI_SQUARED_DIST_FUNC = comp_inv_chi_sq(PROBABILITY_VAL, DEG_FREEDOM)
End Select

Exit Function
ERROR_LABEL:
INVERSE_CHI_SQUARED_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TDIST_FUNC
'DESCRIPTION   : Returns the Percentage Points (probability) for the Student
't-distribution where a numeric value (x) is a calculated value of t
'for which the Percentage Points are to be computed. The t-distribution
'is used in the hypothesis testing of small sample data sets. Use this
'function in place of a table of critical values for the t-distribution.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 025
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function TDIST_FUNC(ByVal X_VAL As Double, _
ByVal DEG_FREEDOM As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True)

'X_VAL: is the value at which you want to evaluate the distribution.

'DEG_FREEDOM: is the number of degrees of freedom.

'CUMUL_FLAG: is a logical value that determines the form of the function. If
'cumulative is TRUE, TDIST_FUNC returns the cumulative distribution function;
'if FALSE, it returns the probability density function.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        TDIST_FUNC = cdf_tdist(X_VAL, DEG_FREEDOM)
        'same AS 1 - TDist(X_VAL,DEG_FREEDOM,1)
    Case False 'probability density function
        TDIST_FUNC = pdf_tdist(X_VAL, DEG_FREEDOM)
End Select

'If XTEMP_VAL > 0 Then
'    TDIST_FUNC = 1 - Excel.Application.tdist(XTEMP_VAL, DEGREES, 1)
'Else
'    TDIST_FUNC = Excel.Application.tdist(-XTEMP_VAL, DEGREES, 1)
'End If

Exit Function
ERROR_LABEL:
TDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_TDIST_FUNC
'DESCRIPTION   : Returns the t-value of the Student's t-distribution as a
'function of the probability and the degrees of freedom.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 026
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

'Student 's t distribution is used commonly for small sample size -usually a
'sample size less than 30.  A t distribution shares some common characteristics
'with the standard normal distribution.  Both distributions are symmetrical,
'both range in value from negative infinity to positive infinity, and both have
'a mean of zero and standard derivation of one.  However, a t distribution has
'a greater dispersion than the standard normal distribution.

'As the degree of freedom (sample size - 1) decreases from simulation 1 to simulation
'3, the standard deviation approaches to 1.  The mean remains close to zero through
'all 3 simulations.

Function INVERSE_TDIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal DEG_FREEDOM As Double)

On Error GoTo ERROR_LABEL

'DEG_FREEDOM: is the number of degrees of freedom.

INVERSE_TDIST_FUNC = inv_tdist(PROBABILITY_VAL, DEG_FREEDOM)
'SAME AS = TINV(PROBABILITY_VAL*2,DEG_FREEDOM) 'USING EXCEL

'If XTEMP_VAL < 0.5 Then
'    INVERSE_TDIST_FUNC = -Excel.Application.TInv(2 * XTEMP_VAL, DEGREES)
'Else
'    INVERSE_TDIST_FUNC = Excel.Application.TInv(2 * (1 - XTEMP_VAL), DEGREES)
'End If

Exit Function
ERROR_LABEL:
INVERSE_TDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FDIST_FUNC
'DESCRIPTION   : Returns the F probability distribution. You can use this
'function to determine whether two data sets have different degrees
'of diversity. For example, you can examine the test scores of men
'and women entering high school and determine if the variability in
'the females is different from that found in the males
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 027
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function FDIST_FUNC(ByVal X_VAL As Double, _
ByVal DEG_FREEDOM_1 As Double, _
ByVal DEG_FREEDOM_2 As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'X_VAL: is the value at which you want to evaluate the distribution.

'DEG_FREEDOM_1: First degree of freedom

'DEG_FREEDOM_2: Second degree of freedom

'CUMUL_FLAG: is a logical value that determines the form of the function. If
'cumulative is TRUE, FDIST_FUNC returns the cumulative distribution function;
'if FALSE, it returns the probability density function.

'COMP_FLAG: 1-cumulative distribution function


'The F distribution is commonly used for ANOVA (analysis of variance), to test
'whether the variances of two or more populations are equal.  For every F deviate,
'there are two degrees of freedom, one in the numerator and one in the denominator.
'It is the ratio of the dispersions of the two Chi-Square distributions.  As both
'of the degree of freedom increase, the percentile value is approaching to one.
'F is also used in tests of "explained variance" and is referred to as the variance
'ration - Explained variance/Unexplained variance.

''The output shows the estimate of skewness, mean, stand deviation, maximum value,
'minimum value, lower confidence interval, and upper confidence interval from each
'of the 3 simulations .  Many things happened as the degree of freedom becoming
'larger from simulation 1 to 3:  the percentile value also approaching to 1; skew
'level decreases (the distribution approaches to normal); mean is approaching to
'1 (mean(F) = df2/(df2-2)); the standard deviation decreases.

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        If COMP_FLAG = True Then
            FDIST_FUNC = cdf_fdist(X_VAL, DEG_FREEDOM_1, DEG_FREEDOM_2)
        ElseIf COMP_FLAG = False Then
            FDIST_FUNC = comp_cdf_fdist(X_VAL, DEG_FREEDOM_1, DEG_FREEDOM_2)
        End If
    Case False 'probability density function
        FDIST_FUNC = pdf_fdist(X_VAL, DEG_FREEDOM_1, DEG_FREEDOM_2)
End Select

Exit Function
ERROR_LABEL:
    FDIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_FDIST_FUNC
'DESCRIPTION   : Returns the inverse of the cumulative distribution function for
'a specified F distribution. That is, if probability = FDIST(x,...),
'then FINV(probability,...) = x. The F distribution can be used
'in project planning to model probable completion times given an expected
'completion time and variability.
'LIBRARY       : STATISTICS
'GROUP         : DIST_CONTINUOUS
'ID            : 028
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function INVERSE_FDIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal DEG_FREEDOM_1 As Double, _
ByVal DEG_FREEDOM_2 As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the F distribution.

'DEG_FREEDOM_1: First degree of freedom

'DEG_FREEDOM_2: Second degree of freedom

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        INVERSE_FDIST_FUNC = inv_fdist(PROBABILITY_VAL, DEG_FREEDOM_1, DEG_FREEDOM_2)
    Case False
        INVERSE_FDIST_FUNC = comp_inv_fdist(PROBABILITY_VAL, DEG_FREEDOM_1, DEG_FREEDOM_2)
End Select

Exit Function
ERROR_LABEL:
INVERSE_FDIST_FUNC = Err.number
End Function

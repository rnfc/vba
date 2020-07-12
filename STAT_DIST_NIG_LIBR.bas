Attribute VB_Name = "STAT_DIST_NIG_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'--------------------------------------------------------------------------------
'Theoretical moments of the NIG (Normal Inverse Gaussian) Distribution.
'--------------------------------------------------------------------------------

Private PUB_ALPHA_VAL As Double
Private PUB_BETA_VAL As Double
Private PUB_MU_VAL As Double
Private PUB_DELTA_VAL As Double
Private Const PUB_EPSILON As Double = 2 ^ 52


'************************************************************************************
'************************************************************************************
'FUNCTION      : NIG_MLE_SOLVER_FUNC
'DESCRIPTION   : Maximum Likelihood Estimation of Univariate NIG parameters
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function NIG_MLE_SOLVER_FUNC(ByRef DATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

'PARAM_RNG --> Moment Matching --> Mean, Var, Skew, Kurt

On Error GoTo ERROR_LABEL

NIG_MLE_SOLVER_FUNC = NELDER_MEAD_OPTIMIZATION2_FUNC("NIG_MLE_OBJ_FUNC", DATA_RNG, PARAM_RNG)

Exit Function
ERROR_LABEL:
NIG_MLE_SOLVER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NIG_MLE_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function NIG_MLE_OBJ_FUNC(ByRef DATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim TEMP_SUM As Double
Dim DATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

If (Abs(PARAM_VECTOR(2, 1)) > PARAM_VECTOR(1, 1)) Or (PARAM_VECTOR(4, 1) < 0) Then
    NIG_MLE_OBJ_FUNC = PUB_EPSILON
Else
    TEMP_SUM = 0
    For i = 1 To UBound(DATA_VECTOR, 1)
        TEMP_SUM = TEMP_SUM + Log(NIG_PDF_FUNC(DATA_VECTOR(i, 1) + 0, PARAM_VECTOR(1, 1) + 0, PARAM_VECTOR(2, 1) + 0, PARAM_VECTOR(3, 1) + 0, PARAM_VECTOR(4, 1) + 0))
    Next i
    NIG_MLE_OBJ_FUNC = -1 * TEMP_SUM
End If

Exit Function
ERROR_LABEL:
NIG_MLE_OBJ_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NIG_PDF_FUNC
'DESCRIPTION   : NIG Prob distribution function. Uses a custom Bessel function
' in order to minimize compatibility issues
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function NIG_PDF_FUNC(ByVal X_VAL As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal MU_VAL As Double, _
ByVal DELTA_VAL As Double)

Dim PI_VAL As Double
Dim Z_VAL As Double
Dim K_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
Z_VAL = Sqr(DELTA_VAL ^ 2 + (X_VAL - MU_VAL) ^ 2)
K_VAL = BESSEL_K1_FUNC(ALPHA_VAL * Z_VAL)
NIG_PDF_FUNC = (DELTA_VAL * ALPHA_VAL / PI_VAL) * Exp(DELTA_VAL * Sqr(ALPHA_VAL ^ 2 - BETA_VAL ^ 2)) / Z_VAL * K_VAL * Exp(BETA_VAL * (X_VAL - MU_VAL))

Exit Function
ERROR_LABEL:
NIG_PDF_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NIG_LLF_FUNC
'DESCRIPTION   : NIG Prob distribution function. Uses a custom Bessel
' function in order to minimize compatibility issues
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function NIG_LLF_FUNC(ByVal DATA_RNG As Variant, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal MU_VAL As Double, _
ByVal DELTA_VAL As Double)

Dim i As Long
Dim NROWS As Long
Dim TEMP_SUM As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

NROWS = UBound(DATA_VECTOR, 1)

TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + Log(NIG_PDF_FUNC(DATA_VECTOR(i, 1), ALPHA_VAL, BETA_VAL, MU_VAL, DELTA_VAL))
Next i

NIG_LLF_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
NIG_LLF_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NIG_CDF_FUNC
'DESCRIPTION   : NIG cumulative probability function
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function NIG_CDF_FUNC(ByVal X_VAL As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal MU_VAL As Double, _
ByVal DELTA_VAL As Double, _
Optional ByVal NO_POINTS As Long = 500, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim START_VAL As Double
Dim END_VAL As Double

Dim DX_VAL As Double
Dim TEMP_SUM As Double

Dim TEMP_VAL As Double
Dim FACTOR_VAL As Double

Dim XTEMP_ARR() As Double
Dim YTEMP_ARR() As Double

Dim epsilon As Double

On Error GoTo ERROR_LABEL

'------------------------------------------------------------------------
Select Case VERSION
'------------------------------------------------------------------------
Case 0 'calculated numerically as described in Beus/Bressers/de Graaf's
'"Alternative Investments and Risk Measurement" (2003)
'------------------------------------------------------------------------
    epsilon = (ALPHA_VAL * DELTA_VAL) ^ 2 - (BETA_VAL * DELTA_VAL) ^ 2
    FACTOR_VAL = Sqr(epsilon)
    For i = 1 To NO_POINTS - 1
        TEMP_VAL = TEMP_VAL + (i ^ -2) * NORMSDIST_FUNC((X_VAL - MU_VAL) / DELTA_VAL, BETA_VAL * DELTA_VAL * (NO_POINTS / i - 1), Sqr(NO_POINTS / i - 1), 0) * IG_PDF_FUNC(NO_POINTS / i - 1, FACTOR_VAL, epsilon)
    Next i
    NIG_CDF_FUNC = NO_POINTS * TEMP_VAL
'------------------------------------------------------------------------
Case 1 'Calculated using Numerical Integration
'------------------------------------------------------------------------
    START_VAL = MU_VAL - 5 * DELTA_VAL
    END_VAL = X_VAL
    ReDim XTEMP_ARR(1 To NO_POINTS + 1)
    ReDim YTEMP_ARR(1 To NO_POINTS + 1)
    DX_VAL = (END_VAL - START_VAL) / NO_POINTS
    For i = 1 To NO_POINTS + 1
        XTEMP_ARR(i) = START_VAL + (i - 1) * DX_VAL
        YTEMP_ARR(i) = NIG_PDF_FUNC(XTEMP_ARR(i), ALPHA_VAL, BETA_VAL, MU_VAL, DELTA_VAL)
    Next i
    TEMP_SUM = 0
    For i = 1 To (NO_POINTS) / 2
        j = (i * 2) - 1
        TEMP_SUM = TEMP_SUM + (1 / 3) * (YTEMP_ARR(j) + 4 * YTEMP_ARR(j + 1) + YTEMP_ARR(j + 2)) * DX_VAL
    Next i
    NIG_CDF_FUNC = TEMP_SUM
'------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------
    PUB_ALPHA_VAL = ALPHA_VAL
    PUB_BETA_VAL = BETA_VAL
    PUB_MU_VAL = MU_VAL
    PUB_DELTA_VAL = DELTA_VAL
    START_VAL = MU_VAL - 5 * DELTA_VAL
    END_VAL = X_VAL
    epsilon = 10 ^ -14 '0.000000000000001
    NIG_CDF_FUNC = GAUSS_KRONROD_INTEGRATION_FUNC("NIG_CDF_OBJ_FUNC", START_VAL, END_VAL, NO_POINTS, epsilon)
 '   NIG_CDF_FUNC = TANH_SINH_FUNC("NIG_CDF_OBJ_FUNC", START_VAL, END_VAL, epsilon, 10)(1)
  '  NIG_CDF_FUNC = ROMBERG_FUNC("NIG_CDF_OBJ_FUNC", START_VAL, END_VAL, epsilon, 14)(1)
   ' NIG_CDF_FUNC = GAUSS_KRONROD_FUNC("NIG_CDF_OBJ_FUNC", START_VAL, END_VAL, epsilon, NO_POINTS)(1)

'------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
NIG_CDF_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NIG_CDF_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function NIG_CDF_OBJ_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
NIG_CDF_OBJ_FUNC = NIG_PDF_FUNC(X_VAL, PUB_ALPHA_VAL, PUB_BETA_VAL, PUB_MU_VAL, PUB_DELTA_VAL)
Exit Function
ERROR_LABEL:
NIG_CDF_OBJ_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IG_PDF_FUNC
'DESCRIPTION   : Probability distribution function of the Inverse Gaussian
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************
'IGpdf
Private Function IG_PDF_FUNC(ByVal X_VAL As Double, _
ByVal A_VAL As Double, _
ByVal B_VAL As Double)

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
IG_PDF_FUNC = (A_VAL / Sqr(2 * PI_VAL * B_VAL)) * (X_VAL ^ (-3 / 2)) * Exp(-0.5 * ((A_VAL - B_VAL * X_VAL) ^ 2) / (B_VAL * X_VAL))

Exit Function
ERROR_LABEL:
IG_PDF_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NIG_INV_CDF_FUNC
'DESCRIPTION   : Inverse NIG cdf
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************
'NIGinvcdf

Function NIG_INV_CDF_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal MU_VAL As Double, _
ByVal DELTA_VAL As Double, _
Optional ByVal CDF_TYPE As Integer = 0, _
Optional ByVal ICDF_TYPE As Integer = 1, _
Optional ByVal NO_POINTS As Long = 500, _
Optional ByVal nLOOPS As Long = 10000, _
Optional ByVal epsilon As Double = 0.0000001)

'Const epsilon As Double = 0.0000001
'Const nLOOPS As Integer = 10000 '20000

Dim i As Long
Dim k As Long

Dim S_VAL As Double
Dim M_VAL As Double

Dim X0_VAL As Double
Dim X1_VAL As Double
Dim X2_VAL As Double
Dim PX_VAL As Double

Dim GAMMA_VAL As Double

On Error GoTo ERROR_LABEL

GAMMA_VAL = Sqr(ALPHA_VAL ^ 2 - BETA_VAL ^ 2)
' set upper and lower bounds with the help of Chebyshev'S_VAL inequality
M_VAL = MU_VAL + DELTA_VAL * (BETA_VAL / GAMMA_VAL)
S_VAL = Sqr(DELTA_VAL * (ALPHA_VAL ^ 2 / GAMMA_VAL ^ 3))

'-----------------------------------------------------------------------------------------------------------------
Select Case ICDF_TYPE
'-----------------------------------------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------------------------------------
    k = 1
    If NIG_CDF_FUNC(M_VAL, ALPHA_VAL, BETA_VAL, MU_VAL, DELTA_VAL, NO_POINTS, CDF_TYPE) > PROBABILITY_VAL Then
        Do While NIG_CDF_FUNC(M_VAL - k * S_VAL, ALPHA_VAL, BETA_VAL, MU_VAL, DELTA_VAL, NO_POINTS, CDF_TYPE) > PROBABILITY_VAL
            k = k + 1
        Loop
        X0_VAL = M_VAL - k * S_VAL
        X2_VAL = M_VAL - (k - 1) * S_VAL
    Else
        Do While NIG_CDF_FUNC(M_VAL + k * S_VAL, ALPHA_VAL, BETA_VAL, MU_VAL, DELTA_VAL, NO_POINTS, CDF_TYPE) < PROBABILITY_VAL
            k = k + 1
        Loop
        X0_VAL = M_VAL + (k - 1) * S_VAL
        X2_VAL = M_VAL + k * S_VAL
    End If
'-----------------------------------------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------------------------------------
    X0_VAL = M_VAL - Sqr(1 / PROBABILITY_VAL - 1) * S_VAL
    X2_VAL = M_VAL + Sqr(1 / (1 - PROBABILITY_VAL) - 1) * S_VAL
'-----------------------------------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------------------------------

For i = 1 To nLOOPS ' find inv cdf value with the bisection method
    X1_VAL = (X0_VAL + X2_VAL) / 2
    PX_VAL = NIG_CDF_FUNC(X1_VAL, ALPHA_VAL, BETA_VAL, MU_VAL, DELTA_VAL, NO_POINTS, CDF_TYPE)
    If Abs(PX_VAL - PROBABILITY_VAL) < epsilon Then
        Exit For
    End If
    If PX_VAL < PROBABILITY_VAL Then
        X0_VAL = X1_VAL
    Else
        X2_VAL = X1_VAL
    End If
Next i
NIG_INV_CDF_FUNC = X1_VAL

Exit Function
ERROR_LABEL:
NIG_INV_CDF_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NIG_RANDOM_FUNC
'DESCRIPTION   : Simulate a NIG distributed random variable as a
'mean-variance mixture (Rydberg-MC method)

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************
'NIGrnd
Function NIG_RANDOM_FUNC(ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal MU_VAL As Double, _
ByVal DELTA_VAL As Double)

Dim Z_VAL As Double
Dim X_VAL As Double

On Error GoTo ERROR_LABEL

Z_VAL = IG_RANDOM_FUNC(DELTA_VAL, Sqr(ALPHA_VAL ^ 2 - BETA_VAL ^ 2))
X_VAL = NORMSINV_FUNC(Rnd(), 0, 1, 0)

NIG_RANDOM_FUNC = MU_VAL + BETA_VAL * Z_VAL + Sqr(Z_VAL) * X_VAL

Exit Function
ERROR_LABEL:
NIG_RANDOM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IG_RANDOM_FUNC
'DESCRIPTION   : Generates an Inverse Gaussian random number
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************
'IGrnd
Function IG_RANDOM_FUNC(ByVal DELTA_VAL As Double, _
ByVal GAMMA_VAL As Double)

Dim Y_VAL As Double
Dim Z_VAL As Double

Dim X1_VAL As Double
Dim X2_VAL As Double

Dim P1_VAL As Double
Dim P2_VAL As Double

On Error GoTo ERROR_LABEL

Z_VAL = INVERSE_CHI_SQUARED_DIST_FUNC(Rnd(), 1, False)
X1_VAL = DELTA_VAL / GAMMA_VAL + 1 / (2 * GAMMA_VAL ^ 2) * (Z_VAL + Sqr(4 * GAMMA_VAL * DELTA_VAL * Z_VAL + Z_VAL ^ 2))
X2_VAL = DELTA_VAL / GAMMA_VAL + 1 / (2 * GAMMA_VAL ^ 2) * (Z_VAL - Sqr(4 * GAMMA_VAL * DELTA_VAL * Z_VAL + Z_VAL ^ 2))
Y_VAL = Rnd()
P1_VAL = DELTA_VAL / (DELTA_VAL + GAMMA_VAL * X1_VAL)
P2_VAL = 1 - P1_VAL
If Y_VAL < P1_VAL Then
    IG_RANDOM_FUNC = X1_VAL
Else
    IG_RANDOM_FUNC = X2_VAL
End If

Exit Function
ERROR_LABEL:
IG_RANDOM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NIG_MLE_MOMENTS_FUNC
'DESCRIPTION   : Maximum Likelihood Moments Estimation (Converts NIG
'parameters to moments)
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************
'NIGmomts
Function NIG_MLE_MOMENTS_FUNC(ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal MU_VAL As Double, _
ByVal DELTA_VAL As Double)

Dim TEMP_VECTOR As Variant
Dim GAMMA_VAL As Double

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To 4, 1 To 1)
'TEMP_VECTOR(1, 1) = "MEAN"
'TEMP_VECTOR(2, 1) = "VARIANCE"
'TEMP_VECTOR(3, 1) = "SKEWNESS"
'TEMP_VECTOR(4, 1) = "KURTOSIS"

GAMMA_VAL = Sqr(ALPHA_VAL ^ 2 - BETA_VAL ^ 2)

TEMP_VECTOR(1, 1) = MU_VAL + DELTA_VAL * (BETA_VAL / GAMMA_VAL)
TEMP_VECTOR(2, 1) = DELTA_VAL * (ALPHA_VAL ^ 2 / GAMMA_VAL ^ 3)
TEMP_VECTOR(3, 1) = 3 * BETA_VAL / ALPHA_VAL / Sqr(DELTA_VAL * GAMMA_VAL)
TEMP_VECTOR(4, 1) = 3 + 3 * (1 + 4 * (BETA_VAL / ALPHA_VAL) ^ 2) / (DELTA_VAL * GAMMA_VAL)

NIG_MLE_MOMENTS_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
NIG_MLE_MOMENTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NIG_MLE_PARAMETERS_FUNC
'DESCRIPTION   : Maximum Likelihood Parameters Estimation (Converts moments to
'NIG parameters)
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INVERSE
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************
'NIGparams
Function NIG_MLE_PARAMETERS_FUNC( _
ByVal MEAN_VAL As Double, _
ByVal VARIANCE_VAL As Double, _
ByVal SKEW_VAL As Double, _
ByVal KURT_VAL As Double)

Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To 4, 1 To 1)
'TEMP_VECTOR(1, 1) = "ALPHA"
'TEMP_VECTOR(2, 1) = "BETA"
'TEMP_VECTOR(3, 1) = "MU"
'TEMP_VECTOR(4, 1) = "DELTA"
TEMP_VECTOR(1, 1) = Sqr((3 * KURT_VAL - 4 * SKEW_VAL ^ 2 - 9) / (VARIANCE_VAL * (KURT_VAL - 5 / 3 * SKEW_VAL ^ 2 - 3) ^ 2))
TEMP_VECTOR(2, 1) = SKEW_VAL / (Sqr(VARIANCE_VAL) * (KURT_VAL - 5 / 3 * SKEW_VAL ^ 2 - 3))
TEMP_VECTOR(3, 1) = MEAN_VAL - 3 * SKEW_VAL * Sqr(VARIANCE_VAL) / (3 * KURT_VAL - 4 * SKEW_VAL ^ 2 - 9)
TEMP_VECTOR(4, 1) = 3 ^ (3 / 2) * Sqr(VARIANCE_VAL * (KURT_VAL - 5 / 3 * SKEW_VAL ^ 2 - 3)) / (3 * KURT_VAL - 4 * SKEW_VAL ^ 2 - 9)

NIG_MLE_PARAMETERS_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
NIG_MLE_PARAMETERS_FUNC = Err.number
End Function


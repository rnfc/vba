Attribute VB_Name = "STAT_REGRESSION_PROBIT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PROBIT_NLLS_FUNC
'DESCRIPTION   : Non Linear Least Square (Probit)
'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_PROBIT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function PROBIT_NLLS_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NO_VAR As Long
Dim NCOLUMNS As Long

Dim TSS_VAL As Double
Dim YSQ_VAL As Double
Dim SSR_VAL As Double
Dim RSQ_VAL As Double
Dim YMEAN_VAL As Double

Dim YFIT_VAL As Double

Dim TEMP_VAL As Double
Dim DELTA_VAL As Double

Dim FACTOR_VAL As Variant

Dim X_MATRIX As Variant
Dim Y_VECTOR As Variant

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim XT_MATRIX As Variant
Dim XTX_MATRIX As Variant
Dim XTXY_MATRIX As Variant

Dim XTXI_MATRIX As Variant
Dim XTXIX_MATRIX As Variant
Dim XTXIXT_MATRIX As Variant

Dim HAT_VECTOR As Variant
Dim COEF_VECTOR As Variant
Dim COEFD_VECTOR As Variant
Dim RESID1_VECTOR As Variant
Dim RESID2_VECTOR As Variant
Dim RSE_VECTOR As Variant

Dim PREDICT_VECTOR As Variant
Dim PHI_VECTOR As Variant

Dim G_VECTOR As Variant ' Gradient Vector
Dim GT_VECTOR As Variant

Dim TEMP_MATRIX As Variant

Dim TEMP_VECTOR As Variant
Dim FACTOR_VECTOR As Variant
Dim ANOVA_MATRIX As Variant

Dim PHI_MATRIX As Variant

Dim nLOOPS As Long
Dim tolerance As Double

On Error GoTo ERROR_LABEL

nLOOPS = 25: tolerance = 10 ^ (-12)
l = 0: DELTA_VAL = 1

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then
    XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)
End If

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(XDATA_MATRIX, 1)
NO_VAR = UBound(XDATA_MATRIX, 2)

'----------------------------------------------------------------------------------------
Select Case INTERCEPT_FLAG
'----------------------------------------------------------------------------------------
Case True
'----------------------------------------------------------------------------------------
    NCOLUMNS = NO_VAR + 1
    ReDim X_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    ReDim XT_MATRIX(1 To NCOLUMNS, 1 To NROWS)
    For i = 1 To NROWS
        X_MATRIX(i, 1) = 1
        XT_MATRIX(1, i) = 1
        For j = 2 To NCOLUMNS
            X_MATRIX(i, j) = XDATA_MATRIX(i, j - 1)
            XT_MATRIX(j, i) = XDATA_MATRIX(i, j - 1)
        Next j
    Next i
    Y_VECTOR = YDATA_VECTOR
    XTX_MATRIX = MMULT_FUNC(XT_MATRIX, X_MATRIX) 'X'X
    XTXI_MATRIX = MATRIX_INVERSE_FUNC(XTX_MATRIX, 0) 'X'X -1
    XTXIX_MATRIX = MMULT_FUNC(XTXI_MATRIX, XT_MATRIX) 'ESTIMATES
    COEF_VECTOR = MMULT_FUNC(XTXIX_MATRIX, Y_VECTOR)
'----------------------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------------------
    NCOLUMNS = NO_VAR
    XT_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)
    XTX_MATRIX = MMULT_FUNC(XT_MATRIX, XDATA_MATRIX) 'X'X
    XTXI_MATRIX = MATRIX_INVERSE_FUNC(XTX_MATRIX, 0) 'X'X -1
    XTXY_MATRIX = MMULT_FUNC(XT_MATRIX, YDATA_VECTOR)
    COEF_VECTOR = MMULT_FUNC(XTXI_MATRIX, XTXY_MATRIX)
    Y_VECTOR = YDATA_VECTOR
    X_MATRIX = XDATA_MATRIX
'----------------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------------

YMEAN_VAL = 0: For i = 1 To NROWS: YMEAN_VAL = YMEAN_VAL + Y_VECTOR(i, 1): Next i
YMEAN_VAL = YMEAN_VAL / NROWS
GoSub COEFFICIENT_LINE

Do While (DELTA_VAL > tolerance) And l < nLOOPS

'-----------------------------------------------------------------------------
'Regress the residuals on the gradient vector; do not include an
'additional intercept term.

'Check for convergence (small values for changes in the coefficients or small changes in SSR).

'Add the computed coefficients to the previous coefficients and compute
'residuals
'-----------------------------------------------------------------------------

    XT_MATRIX = MATRIX_TRANSPOSE_FUNC(G_VECTOR)
    XTX_MATRIX = MMULT_FUNC(XT_MATRIX, G_VECTOR) 'X'X
    XTXI_MATRIX = MATRIX_INVERSE_FUNC(XTX_MATRIX, 0) 'X'X -1
    XTXY_MATRIX = MMULT_FUNC(XT_MATRIX, RESID1_VECTOR)
    COEFD_VECTOR = MMULT_FUNC(XTXI_MATRIX, XTXY_MATRIX)
    
    DELTA_VAL = 0
    For k = 1 To NCOLUMNS
        DELTA_VAL = DELTA_VAL + COEFD_VECTOR(k, 1) ^ 2
        COEF_VECTOR(k, 1) = COEF_VECTOR(k, 1) + COEFD_VECTOR(k, 1)
    Next k
    GoSub COEFFICIENT_LINE
    l = l + 1
Loop
GoSub COEFFICIENT_LINE

'--------------------------------------------------------------------
If (INTERCEPT_FLAG = False) Then: TSS_VAL = YSQ_VAL
'--------------------------------------------------------------------

RSQ_VAL = 1 - (SSR_VAL / TSS_VAL)
PHI_MATRIX = MMULT_FUNC(GT_VECTOR, G_VECTOR)
PHI_MATRIX = MATRIX_INVERSE_FUNC(PHI_MATRIX, 0)

ReDim HAT_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
ReDim FACTOR_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    For k = 1 To NCOLUMNS: TEMP_VECTOR(k, 1) = G_VECTOR(i, k): Next k
    If NCOLUMNS = 1 Then
        FACTOR_VAL = TEMP_VECTOR(1, 1) * PHI_MATRIX(1, 1) * TEMP_VECTOR(1, 1)
    Else
        TEMP_MATRIX = MMULT_FUNC(PHI_MATRIX, TEMP_VECTOR)
        FACTOR_VAL = 0: For k = 1 To NCOLUMNS: FACTOR_VAL = FACTOR_VAL + TEMP_MATRIX(k, 1) * TEMP_VECTOR(k, 1): Next k
    End If
    FACTOR_VECTOR(i, 1) = FACTOR_VAL
    HAT_VECTOR(i, 1) = RESID2_VECTOR(i, 1) / (1 - FACTOR_VECTOR(i, 1))
Next i

'--------Robust SE using Davidson and MacKinnon's "better approach" (p. 553);
'--------implemented here.
ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    For k = 1 To NCOLUMNS
        TEMP_VAL = 0
        For i = 1 To NROWS
            TEMP_MATRIX(j, k) = G_VECTOR(i, j) * G_VECTOR(i, k) * HAT_VECTOR(i, 1) + TEMP_VAL
            TEMP_VAL = TEMP_MATRIX(j, k)
        Next i
    Next k
Next j

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------

If NCOLUMNS = 1 Then
    ReDim XTXIXT_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
    XTXIXT_MATRIX(1, 1) = PHI_MATRIX(1, 1) * TEMP_MATRIX(1, 1)
Else
    XTXIXT_MATRIX = MMULT_FUNC(PHI_MATRIX, TEMP_MATRIX)
End If

If NCOLUMNS = 1 Then
    ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
    TEMP_MATRIX(1, 1) = XTXIXT_MATRIX(1, 1) * PHI_MATRIX(1, 1)
Else
    TEMP_MATRIX = MMULT_FUNC(XTXIXT_MATRIX, PHI_MATRIX)
End If

'  Now get the robust standard errors, square roots of diagonal entries
'  Davidson and MacKinnon textbook (1993) p. 553 recommends using a correction
'  factor in which one divides the estimate of the standard error
'  by (1- ht) where ht is the square root of the t'th diagonal entry in
'  the "hat matrix".

'  This hat matrix is sometimes called P because it projects orthogonally
'  onto the space spanned by the columns of X.

'  Stata uses a much simpler correction namely sqrt(N/(N-K)).
'  Davidson and MacKinnon (553-54) say that Stata's correction is
'  inferior to dividing by (1-ht).

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
ReDim ANOVA_MATRIX(1 To 3 + NCOLUMNS, 1 To 3)

ANOVA_MATRIX(1, 1) = "Pseudo R-Sqr" 'SS Residuals
ANOVA_MATRIX(1, 2) = RSQ_VAL
ANOVA_MATRIX(1, 3) = ""
 
ANOVA_MATRIX(2, 1) = "SS[res]" 'SS Residuals
ANOVA_MATRIX(2, 2) = SSR_VAL 'SS Regression = TSS_VAL-SSR
ANOVA_MATRIX(2, 3) = ""
    
ANOVA_MATRIX(3, 1) = "VAR"
ANOVA_MATRIX(3, 2) = "COEF"
ANOVA_MATRIX(3, 3) = "EST. SE"

If INTERCEPT_FLAG = True Then
    ANOVA_MATRIX(4, 1) = "Alpha"
Else
    ANOVA_MATRIX(4, 1) = "Beta: " & 1
End If

ANOVA_MATRIX(4, 2) = COEF_VECTOR(1, 1)
ANOVA_MATRIX(4, 3) = ""

If (NCOLUMNS <> 1) Then
    For i = 2 To NCOLUMNS
        If NO_VAR < NCOLUMNS Then
            ANOVA_MATRIX(3 + i, 1) = "Beta: " & i - 1
        Else
            ANOVA_MATRIX(3 + i, 1) = "Beta: " & i
        End If
        ANOVA_MATRIX(3 + i, 2) = COEF_VECTOR(i, 1)
        ANOVA_MATRIX(3 + i, 3) = ""
    Next i
End If

'If DELTA_VAL > tolerance Then: GoTo ERROR_LABEL 'This is probably a case of perfect classification."

ReDim RSE_VECTOR(1 To NCOLUMNS, 1 To 1)
For i = 1 To NCOLUMNS
    If TEMP_MATRIX(i, i) < 0 Then: TEMP_MATRIX(i, i) = 0
    RSE_VECTOR(i, 1) = Sqr(TEMP_MATRIX(i, i))
    ANOVA_MATRIX(3 + i, 3) = RSE_VECTOR(i, 1)
Next i
    
PROBIT_NLLS_FUNC = ANOVA_MATRIX
 
Exit Function
'---------------------------------------------------------------------------------------------
COEFFICIENT_LINE:
'---------------------------------------------------------------------------------------------
    SSR_VAL = 0: TSS_VAL = 0: YSQ_VAL = 0
    ReDim RESID1_VECTOR(1 To NROWS, 1 To 1)
    ReDim RESID2_VECTOR(1 To NROWS, 1 To 1)
    
    ReDim PHI_VECTOR(1 To NROWS, 1 To 1)
    ReDim PREDICT_VECTOR(1 To NROWS, 1 To 1)
    
    ReDim G_VECTOR(1 To NROWS, 1 To NCOLUMNS)
    ReDim GT_VECTOR(1 To NCOLUMNS, 1 To NROWS)
    
    For i = 1 To NROWS
        YFIT_VAL = 0
        For j = 1 To NCOLUMNS
            YFIT_VAL = YFIT_VAL + COEF_VECTOR(j, 1) * X_MATRIX(i, j)
        Next j
        TEMP_VAL = NORMSDIST_FUNC(-1 * YFIT_VAL, 0, 1, 0)
        PREDICT_VECTOR(i, 1) = 1 - TEMP_VAL
        If Y_VECTOR(i, 1) = 1 Then
            RESID1_VECTOR(i, 1) = TEMP_VAL
        Else
            RESID1_VECTOR(i, 1) = TEMP_VAL - 1
        End If
'    RESID1_VECTOR(i, 1) = Y_VECTOR(i, 1) - PREDICT_VECTOR(i, 1)
        PHI_VECTOR(i, 1) = NORMDIST_FUNC(YFIT_VAL, 0, 1, 0)
        RESID2_VECTOR(i, 1) = RESID1_VECTOR(i, 1) ^ 2
        TSS_VAL = TSS_VAL + (Y_VECTOR(i, 1) - YMEAN_VAL) ^ 2
        YSQ_VAL = YSQ_VAL + Y_VECTOR(i, 1) ^ 2
        SSR_VAL = SSR_VAL + RESID2_VECTOR(i, 1)
    
        For j = 1 To NCOLUMNS
            G_VECTOR(i, j) = PHI_VECTOR(i, 1) * X_MATRIX(i, j) 'gradient vector
            GT_VECTOR(j, i) = G_VECTOR(i, j)
        Next j
    Next i
'---------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------
ERROR_LABEL:
PROBIT_NLLS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : PROBIT_ML_FUNC
'DESCRIPTION   : Probit Maximum Likelihood Results For Dependent Variable
'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_PROBIT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function PROBIT_ML_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NO_VAR As Long
Dim NCOLUMNS As Long
    
Dim TSS_VAL As Double
Dim YSQ_VAL As Double
Dim RSQ_VAL As Double
Dim YMEAN_VAL As Double

Dim LN1_VAL As Double
Dim LN2_VAL As Double

Dim BASE_VAL As Double
Dim FIT_VAL As Double
Dim DELTA_VAL As Double
Dim TEMP_VAL As Double
Dim TEMP_SUM As Double

Dim FACTOR_VAL As Variant

Dim X_MATRIX As Variant
Dim Y_VECTOR As Variant

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim XT_MATRIX As Variant
Dim XTX_MATRIX As Variant
Dim XTXY_MATRIX As Variant

Dim XTXI_MATRIX As Variant
Dim XTXIX_MATRIX As Variant

Dim YFIT_VECTOR As Variant
Dim COEF_VECTOR As Variant
Dim HAT_VECTOR As Variant
Dim PHI_VECTOR As Variant
Dim RSE_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim ANOVA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

Dim nLOOPS As Long
Dim epsilon As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL

nLOOPS = 200: tolerance = 0.00000001: epsilon = 0.01
l = 0: DELTA_VAL = 1

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then
    XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)
End If

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(XDATA_MATRIX, 1)
NO_VAR = UBound(XDATA_MATRIX, 2)

'----------------------------------------------------------------------------------------
Select Case INTERCEPT_FLAG
'----------------------------------------------------------------------------------------
Case True
'----------------------------------------------------------------------------------------
    NCOLUMNS = NO_VAR + 1
    ReDim X_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    ReDim XT_MATRIX(1 To NCOLUMNS, 1 To NROWS)
    For i = 1 To NROWS
        X_MATRIX(i, 1) = 1
        XT_MATRIX(1, i) = 1
        For j = 2 To NCOLUMNS
            X_MATRIX(i, j) = XDATA_MATRIX(i, j - 1)
            XT_MATRIX(j, i) = XDATA_MATRIX(i, j - 1)
        Next j
    Next i
    Y_VECTOR = YDATA_VECTOR
    XTX_MATRIX = MMULT_FUNC(XT_MATRIX, X_MATRIX) 'X'X
    XTXI_MATRIX = MATRIX_INVERSE_FUNC(XTX_MATRIX, 0) 'X'X -1
    XTXIX_MATRIX = MMULT_FUNC(XTXI_MATRIX, XT_MATRIX) 'ESTIMATES
    COEF_VECTOR = MMULT_FUNC(XTXIX_MATRIX, Y_VECTOR)
'----------------------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------------------
    NCOLUMNS = NO_VAR
    XT_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)
    XTX_MATRIX = MMULT_FUNC(XT_MATRIX, XDATA_MATRIX) 'X'X
    XTXI_MATRIX = MATRIX_INVERSE_FUNC(XTX_MATRIX, 0) 'X'X -1
    XTXY_MATRIX = MMULT_FUNC(XT_MATRIX, YDATA_VECTOR)
    COEF_VECTOR = MMULT_FUNC(XTXI_MATRIX, XTXY_MATRIX)
    Y_VECTOR = YDATA_VECTOR
    X_MATRIX = XDATA_MATRIX
'----------------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------------

YMEAN_VAL = 0: For i = 1 To NROWS: YMEAN_VAL = YMEAN_VAL + Y_VECTOR(i, 1): Next i
YMEAN_VAL = YMEAN_VAL / NROWS

TSS_VAL = 0: YSQ_VAL = 0
ReDim YFIT_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    YFIT_VECTOR(i, 1) = 0
    For j = 1 To NCOLUMNS
        YFIT_VECTOR(i, 1) = YFIT_VECTOR(i, 1) + COEF_VECTOR(j, 1) * X_MATRIX(i, j)
    Next j
Next i

'-----------------Computing LnBeta vector and Hessian matrix-------------
ReDim PHI_VECTOR(1 To NROWS, 1 To 1)
ReDim HAT_VECTOR(1 To NROWS, 1 To 1)
ReDim LAMBDA_VECTOR(1 To NROWS, 1 To 1)
ReDim SCALAR_VECTOR(1 To NROWS, 1 To 1)

Do While (Abs(DELTA_VAL) > tolerance) And l < nLOOPS
    LN1_VAL = 0
    l = l + 1
    ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
    ReDim TEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NROWS 'Avoid problems when HAT_VECTOR is too close to 0 or 1
        TEMP_VAL = NORMSDIST_FUNC(-YFIT_VECTOR(i, 1), 0, 1, 0)
        HAT_VECTOR(i, 1) = 1 - TEMP_VAL
        If HAT_VECTOR(i, 1) = 0 Then HAT_VECTOR(i, 1) = NORMSDIST_FUNC(YFIT_VECTOR(i, 1), 0, 1, 0)
        PHI_VECTOR(i, 1) = NORMDIST_FUNC(YFIT_VECTOR(i, 1), 0, 1, 0)
        LAMBDA_VECTOR(i, 1) = ((Y_VECTOR(i, 1) * PHI_VECTOR(i, 1) / HAT_VECTOR(i, 1)) + (1 - Y_VECTOR(i, 1)) * (-PHI_VECTOR(i, 1) / (TEMP_VAL)))
        SCALAR_VECTOR(i, 1) = -LAMBDA_VECTOR(i, 1) * (LAMBDA_VECTOR(i, 1) + YFIT_VECTOR(i, 1))
        For k = 1 To NCOLUMNS
            ' TEMP_VECTOR(k,1) = TEMP_VECTOR(k,1) + ((YTEMP(i, 1) - HAT_VECTOR(i,1)) / (HAT_VECTOR(i,1) * (1 - HAT_VECTOR(i,1)))) * PHI_VECTOR(i,1) * XTEMP(i, k)
            TEMP_VECTOR(k, 1) = TEMP_VECTOR(k, 1) + ((Y_VECTOR(i, 1) - HAT_VECTOR(i, 1)) / (HAT_VECTOR(i, 1) * TEMP_VAL)) * PHI_VECTOR(i, 1) * X_MATRIX(i, k)
            For j = 1 To NCOLUMNS
                TEMP_MATRIX(j, k) = SCALAR_VECTOR(i, 1) * X_MATRIX(i, j) * X_MATRIX(i, k) + TEMP_MATRIX(j, k)
            Next j
        Next k
        LN1_VAL = LN1_VAL + Y_VECTOR(i, 1) * Log(HAT_VECTOR(i, 1)) + (1 - Y_VECTOR(i, 1)) * Log(TEMP_VAL)
    Next i
'-------------------------------------------------------------------------
'---------------------------Invert Hessian--------------------------------
    If NCOLUMNS = 1 Then
        FACTOR_VAL = (1 / TEMP_MATRIX(1, 1)) * TEMP_VECTOR(1, 1)
        For k = 1 To NCOLUMNS
            COEF_VECTOR(k, 1) = COEF_VECTOR(k, 1) - FACTOR_VAL
        Next k
    Else
        TEMP_MATRIX = MATRIX_INVERSE_FUNC(TEMP_MATRIX, 0)
        FACTOR_VAL = MMULT_FUNC(TEMP_MATRIX, TEMP_VECTOR)
        For k = 1 To NCOLUMNS
            COEF_VECTOR(k, 1) = COEF_VECTOR(k, 1) - FACTOR_VAL(k, 1)
        Next k
    End If

    DELTA_VAL = LN1_VAL - LN2_VAL
    LN2_VAL = LN1_VAL

    For i = 1 To NROWS
        YFIT_VECTOR(i, 1) = 0
        For k = 1 To NCOLUMNS
            YFIT_VECTOR(i, 1) = YFIT_VECTOR(i, 1) + COEF_VECTOR(k, 1) * X_MATRIX(i, k)
        Next k
    Next i

Loop

'If DELTA_VAL > tolerance Then GoTo ERROR_LABEL
'Convergence not achieved.

BASE_VAL = 0
FIT_VAL = Log(YMEAN_VAL)

For i = 1 To NROWS
    BASE_VAL = BASE_VAL + Y_VECTOR(i, 1) * FIT_VAL + (1 - Y_VECTOR(i, 1)) * Log(1 - YMEAN_VAL)
Next i
RSQ_VAL = 1 - LN1_VAL / BASE_VAL ' Compute the pseudo R^2

If NCOLUMNS = 1 Then
    TEMP_SUM = 0
    For k = 1 To NCOLUMNS
        TEMP_SUM = TEMP_SUM + Abs(FACTOR_VAL)
    Next k
Else
    TEMP_SUM = 0
    For k = 1 To NCOLUMNS
        TEMP_SUM = TEMP_SUM + Abs(FACTOR_VAL(k, 1))
    Next k
End If


'If TEMP_SUM > epsilon Then: GoTo ERROR_LABEL 'WARNING: This is a case of perfect classification
    
ReDim ANOVA_MATRIX(1 To 3 + NCOLUMNS, 1 To 3)

ANOVA_MATRIX(1, 1) = "Pseudo R-Sqr" 'SS Residuals
ANOVA_MATRIX(1, 2) = RSQ_VAL
ANOVA_MATRIX(1, 3) = ""
 
ANOVA_MATRIX(2, 1) = "Log likelihood (lnL)"
ANOVA_MATRIX(2, 2) = LN1_VAL
ANOVA_MATRIX(2, 3) = ""
    
ANOVA_MATRIX(3, 1) = "VAR"
ANOVA_MATRIX(3, 2) = "COEF"
ANOVA_MATRIX(3, 3) = "EST. SE"

If INTERCEPT_FLAG = True Then
    ANOVA_MATRIX(4, 1) = "Alpha"
Else
    ANOVA_MATRIX(4, 1) = "Beta: " & 1
End If

ANOVA_MATRIX(4, 2) = COEF_VECTOR(1, 1)
ANOVA_MATRIX(4, 3) = ""

If (NCOLUMNS <> 1) Then
    For i = 2 To NCOLUMNS
        If NO_VAR < NCOLUMNS Then
            ANOVA_MATRIX(3 + i, 1) = "Beta: " & i - 1
        Else
            ANOVA_MATRIX(3 + i, 1) = "Beta: " & i
        End If
        ANOVA_MATRIX(3 + i, 2) = COEF_VECTOR(i, 1)
        ANOVA_MATRIX(3 + i, 3) = ""
    Next i
End If

'If DELTA_VAL > tolerance Then: GoTo ERROR_LABEL 'This is probably a case of perfect classification."

ReDim RSE_VECTOR(1 To NCOLUMNS, 1 To 1)
For i = 1 To NCOLUMNS
    RSE_VECTOR(i, 1) = Sqr(-TEMP_MATRIX(i, i))
    ANOVA_MATRIX(3 + i, 3) = RSE_VECTOR(i, 1)
Next i

PROBIT_ML_FUNC = ANOVA_MATRIX
 
Exit Function
ERROR_LABEL:
PROBIT_ML_FUNC = Err.number
End Function

'////////////////////////////////////PERFECT\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'Dummy Dependent variable models

'In addition to coefficient estimates, these functions report estimated SEs
'These functions do not use Excel's Solver and the SEs are computed in the
'estimation step. The nonlinear least squares (NLLS) estimates
'use robust standard errors to correct for heteroscedasticity.
'Finally, they compute pseudo R2 values for the maximum likelihood
'routines.   The pseudo R2 value is the correlation between the observed
'values of the dependent variable and the fitted values.


'Logistic regression

'In particular there is the "logit model" of which the simplest sort is
'Log(yi) = a + xi

'where xi is some quantity on which success or failure in the i-th in a
'sequence of Bernoulli trials may depend, and pi is the probability of
'success in the i-th case. For example, x may be the age of a patient
'admitted to a hospital with a heart attack, and "success" may be the
'event that the patient dies before leaving the hospital. Having observed
'the values of x in a sequence of cases and whether there was a "success"
'or a "failure" in each such case.

'The result can then be used to assess the probability of "success" in a
'subsequent case in which the value of x is known. Estimation and prediction
'by this method are called logistic regression.

'-------------------------------------------------------------------------------
'-----------------Documentation for Probit and Logit Routines-------------------
'-------------------------------------------------------------------------------

'Davidson, R. and J. G. MacKinnon (1993). Estimation and Inference in
'Econometrics. New York, Oxford University Press.

'Goldberger, A. S. (1991). A Course in Econometrics. Cambridge, Mass.,
'Harvard University.

'Ruud, P. A. (2000). An Introduction to Classical Econometric Theory.
'New York, Oxford Universtity Press.

'Wooldridge, J. M. (2002). Econometric Analysis of Cross Section and
'Panel Data, The MIT Press.

'Wooldridge, J. M. (2000). Introductory Econometrics: A Modern Approach,
'Southwestern College Publishing.

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

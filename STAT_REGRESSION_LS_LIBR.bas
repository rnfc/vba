Attribute VB_Name = "STAT_REGRESSION_LS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : REGRESSION_LS1_FUNC
'DESCRIPTION   : Multiple regression Frame: We use the adjustment to robust
'standard errors suggested by Davidson and MacKinnon (1993).
'LIBRARY       : REGRESSION
'GROUP         : MULTIPLE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function REGRESSION_LS1_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = True, _
Optional ByVal SE_VERSION As Long = 2, _
Optional ByVal OUTPUT As Integer = 0)

'---------------------------------------------------------------------------

' Uses LU Factorization for the Inverse of a Matrix

'  Davidson and MacKinnon textbook (1993) p. 553 recommends
'  using a correction factor in which one divides the estimate
'  of the standard error by a factor where the factor is the
'  square root of the t'th diagonal entry in the "hat matrix".

'  This hat matrix is sometimes called P because it projects
'  orthogonally onto the space spanned by the columns of
'  the independent variables.

'  Stata uses a much simpler correction namely sqrt(N/(N-K)).
'  Davidson and MacKinnon (553-54) say that Stata's correction
'  is inferior to dividing by (1-factor).

'REFERENCES:

'Davidson, R. and J. G. MacKinnon (1993). Estimation and Inference in
'Econometrics. New York, Oxford University Press.

'Goldberger, A. S. (1991). A Course in Econometrics.
'Cambridge, Mass., Harvard University.

'Ruud, P. A. (2000). An Introduction to Classical Econometric Theory.
'New York, Oxford Universtity Press.

'Wooldridge, J. M. (2002). Econometric Analysis of Cross Section and
'Panel Data, The MIT Press.

'Wooldridge, J. M. (2000). Introductory Econometrics: A Modern Approach,
'Southwestern College Publishing.

'---------------------------------------------------------------------------
Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NO_VAR As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim FSTAT_VAL As Double
Dim MULT_VAL As Double

Dim TSS_VAL As Double
Dim YSQ_VAL As Double
Dim SSR_VAL As Double
Dim RSQ_VAL As Double
Dim RMSE_VAL As Double
Dim YFIT_VAL As Double
Dim YMEAN_VAL As Double

Dim SE_VECTOR As Variant
Dim RSE_VECTOR As Variant
Dim HT_VECTOR As Variant
Dim HAT_VECTOR As Variant
Dim COEF_VECTOR As Variant
Dim RESID_VECTOR As Variant
Dim RESID_SQR_VECTOR As Variant

Dim Y_VECTOR As Variant
Dim X_MATRIX As Variant
Dim XT_MATRIX As Variant
Dim XTX_MATRIX As Variant
Dim XTY_MATRIX As Variant

Dim XTXI_MATRIX As Variant
Dim XTXIXT_MATRIX As Variant

Dim S_MATRIX As Variant
Dim XTXIS_MATRIX As Variant
Dim RSE_MATRIX As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then
    XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then
    GoTo ERROR_LABEL
End If

NROWS = UBound(XDATA_MATRIX, 1)
NO_VAR = UBound(XDATA_MATRIX, 2)
    
'--------------------------------------------------------------------------------------------------------------
Select Case INTERCEPT_FLAG
'--------------------------------------------------------------------------------------------------------------
Case True
'--------------------------------------------------------------------------------------------------------------
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
    XTXIXT_MATRIX = MMULT_FUNC(XTXI_MATRIX, XT_MATRIX) 'ESTIMATES
    COEF_VECTOR = MMULT_FUNC(XTXIXT_MATRIX, Y_VECTOR)
'--------------------------------------------------------------------------------------------------------------
Case False
'--------------------------------------------------------------------------------------------------------------
    NCOLUMNS = NO_VAR
    XT_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)
    XTX_MATRIX = MMULT_FUNC(XT_MATRIX, XDATA_MATRIX) 'X'X
    XTXI_MATRIX = MATRIX_INVERSE_FUNC(XTX_MATRIX, 0) 'X'X -1
    XTY_MATRIX = MMULT_FUNC(XT_MATRIX, YDATA_VECTOR)
    COEF_VECTOR = MMULT_FUNC(XTXI_MATRIX, XTY_MATRIX)
        
    Y_VECTOR = YDATA_VECTOR
    X_MATRIX = XDATA_MATRIX
'--------------------------------------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------------------------------------

ReDim RESID_VECTOR(1 To NROWS, 1 To 1)
ReDim RESID_SQR_VECTOR(1 To NROWS, 1 To 1)
ReDim RSE_VECTOR(1 To NCOLUMNS, 1 To 1)

YMEAN_VAL = 0
For i = 1 To NROWS
    YMEAN_VAL = YMEAN_VAL + Y_VECTOR(i, 1)
Next i
YMEAN_VAL = YMEAN_VAL / NROWS

RMSE_VAL = 0: TSS_VAL = 0: YSQ_VAL = 0
For i = 1 To NROWS
     YFIT_VAL = 0
     TSS_VAL = TSS_VAL + (Y_VECTOR(i, 1) - YMEAN_VAL) ^ 2
     For j = 1 To NCOLUMNS
         YFIT_VAL = YFIT_VAL + COEF_VECTOR(j, 1) * X_MATRIX(i, j)
     Next j
     RESID_VECTOR(i, 1) = Y_VECTOR(i, 1) - YFIT_VAL
     RESID_SQR_VECTOR(i, 1) = RESID_VECTOR(i, 1) ^ 2
     RMSE_VAL = RMSE_VAL + RESID_SQR_VECTOR(i, 1)
     YSQ_VAL = YSQ_VAL + Y_VECTOR(i, 1) ^ 2
Next i

'--------------------------------------------------------------------
If (INTERCEPT_FLAG = False) Then: TSS_VAL = YSQ_VAL
'--------------------------------------------------------------------
SSR_VAL = RMSE_VAL
RMSE_VAL = (RMSE_VAL / (NROWS - NCOLUMNS)) ^ 0.5
RSQ_VAL = 1 - (SSR_VAL / TSS_VAL)
If (INTERCEPT_FLAG = True) Then
    FSTAT_VAL = ((TSS_VAL - SSR_VAL) / (NCOLUMNS - 1)) / (SSR_VAL / (NROWS - NCOLUMNS))
Else
    FSTAT_VAL = ((TSS_VAL - SSR_VAL) / (NCOLUMNS)) / (SSR_VAL / (NROWS - NCOLUMNS))
End If

If NCOLUMNS = 1 Then
    ReDim SE_VECTOR(1 To NCOLUMNS, 1 To NCOLUMNS)
    SE_VECTOR(1, 1) = XTXI_MATRIX(1, 1) ^ 0.5 * RMSE_VAL
Else
    ReDim SE_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        SE_VECTOR(i, 1) = XTXI_MATRIX(i, i) ^ 0.5 * RMSE_VAL
    Next i
End If

ReDim HT_VECTOR(1 To NROWS, 1 To 1)
ReDim HAT_VECTOR(1 To NROWS, 1 To 1)

'-------------------------------------------------------------------------
Select Case SE_VERSION
'-------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------
    For h = 1 To NROWS
        HAT_VECTOR(h, 1) = RESID_SQR_VECTOR(h, 1)
    Next h
'-------------------------------------------------------------------------
Case 1
'-------------------------------------------------------------------------
    For h = 1 To NROWS
        HAT_VECTOR(h, 1) = RESID_SQR_VECTOR(h, 1) * (NROWS / (NROWS - NCOLUMNS))
    Next h
'-------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------
    ReDim XT_VECTOR(1 To NCOLUMNS, 1 To 1)
    For h = 1 To NROWS
        For k = 1 To NCOLUMNS
            XT_VECTOR(k, 1) = X_MATRIX(h, k)
        Next k
        If NCOLUMNS = 1 Then
            MULT_VAL = XT_VECTOR(1, 1) * XTXI_MATRIX(1, 1) * XT_VECTOR(1, 1)
        Else
            ReDim YTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
            ReDim XTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
            For i = 1 To NCOLUMNS: XTEMP_VECTOR(i, 1) = XT_VECTOR(i, 1): Next i
            For i = 1 To NCOLUMNS
                For j = 1 To 1
                    YTEMP_VECTOR(i, j) = 0
                    For k = 1 To NCOLUMNS: YTEMP_VECTOR(i, j) = XTXI_MATRIX(i, k) * XTEMP_VECTOR(k, j) + YTEMP_VECTOR(i, j): Next k
                Next j
            Next i
            MULT_VAL = 0
            For i = 1 To NCOLUMNS
                MULT_VAL = MULT_VAL + YTEMP_VECTOR(i, 1) * XT_VECTOR(i, 1)
            Next i
        End If
        HT_VECTOR(h, 1) = MULT_VAL
        If HT_VECTOR(h, 1) = 1 Then
            HAT_VECTOR(h, 1) = 0
        Else
            If SE_VERSION = 2 Then
                HAT_VECTOR(h, 1) = RESID_SQR_VECTOR(h, 1) / (1 - HT_VECTOR(h, 1))
            Else
                HAT_VECTOR(h, 1) = RESID_SQR_VECTOR(h, 1) / ((1 - HT_VECTOR(h, 1)) ^ 2) 'Here is the Difference between Case 2
            End If
        End If
    Next h
'-------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------

ReDim S_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    For k = 1 To NCOLUMNS
        TEMP_SUM = 0
        For i = 1 To NROWS
            S_MATRIX(j, k) = X_MATRIX(i, j) * X_MATRIX(i, k) * HAT_VECTOR(i, 1) + TEMP_SUM
            TEMP_SUM = S_MATRIX(j, k)
        Next i
    Next k
Next j

ReDim XTXIS_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
If NCOLUMNS = 1 Then
    XTXIS_MATRIX(1, 1) = XTXI_MATRIX(1, 1) * S_MATRIX(1, 1)
Else
    For i = 1 To NCOLUMNS
        For j = 1 To NCOLUMNS
            XTXIS_MATRIX(i, j) = 0
            For k = 1 To NCOLUMNS
                XTXIS_MATRIX(i, j) = XTXI_MATRIX(i, k) * S_MATRIX(k, j) + XTXIS_MATRIX(i, j)
            Next k
        Next j
    Next i
End If

ReDim RSE_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
If NCOLUMNS = 1 Then
    RSE_MATRIX(1, 1) = XTXIS_MATRIX(1, 1) * XTXI_MATRIX(1)
Else
    For i = 1 To NCOLUMNS
        For j = 1 To NCOLUMNS
            RSE_MATRIX(i, j) = 0
            For k = 1 To NCOLUMNS
                RSE_MATRIX(i, j) = XTXIS_MATRIX(i, k) * XTXI_MATRIX(k, j) + RSE_MATRIX(i, j)
            Next k
        Next j
    Next i
End If

For i = 1 To NCOLUMNS
    If RSE_MATRIX(i, i) < 0 Then: RSE_MATRIX(i, i) = 0
    RSE_VECTOR(i, 1) = RSE_MATRIX(i, i) ^ 0.5
Next i

ReDim TEMP_MATRIX(1 To 5 + NCOLUMNS, 1 To 4)

TEMP_MATRIX(1, 1) = "OBS"
TEMP_MATRIX(1, 2) = NROWS
 
TEMP_MATRIX(2, 1) = "YMEAN"
TEMP_MATRIX(2, 2) = YMEAN_VAL

TEMP_MATRIX(3, 1) = "RMSE"
TEMP_MATRIX(3, 2) = RMSE_VAL

TEMP_MATRIX(4, 1) = "F-STAT"
TEMP_MATRIX(4, 2) = FSTAT_VAL

TEMP_MATRIX(1, 3) = "SSR" 'SS Residuals
TEMP_MATRIX(1, 4) = SSR_VAL

TEMP_MATRIX(3, 3) = "SST" 'SS Total
TEMP_MATRIX(3, 4) = TSS_VAL

TEMP_MATRIX(2, 3) = "SSREG" 'SS Regression
TEMP_MATRIX(2, 4) = TSS_VAL - SSR_VAL

TEMP_MATRIX(4, 3) = "R^2"
TEMP_MATRIX(4, 4) = RSQ_VAL
'RSQ = 1 - SSR_VAL / (NROWS * YSIGMA_VAL ^ 2)
'RSQ = 1 - NROWS / (NROWS - 1) * RSQ

TEMP_MATRIX(5, 1) = "VAR"
TEMP_MATRIX(5, 2) = "COEF"
TEMP_MATRIX(5, 3) = "SE"

If INTERCEPT_FLAG = True Then
    TEMP_MATRIX(6, 1) = "Alpha"
Else
    TEMP_MATRIX(6, 1) = "Beta: " & 1
End If

TEMP_MATRIX(6, 2) = COEF_VECTOR(1, 1)
TEMP_MATRIX(6, 3) = SE_VECTOR(1, 1)

If (NCOLUMNS <> 1) Then
    For i = 2 To NCOLUMNS
        If NO_VAR < NCOLUMNS Then
            TEMP_MATRIX(5 + i, 1) = "Beta: " & i - 1
        Else
            TEMP_MATRIX(5 + i, 1) = "Beta: " & i
        End If
        TEMP_MATRIX(5 + i, 2) = COEF_VECTOR(i, 1)
        TEMP_MATRIX(5 + i, 3) = SE_VECTOR(i, 1)
    Next i
End If

Select Case SE_VERSION
Case 0
    TEMP_MATRIX(5, 4) = "HC0"
Case 1
    TEMP_MATRIX(5, 4) = "HC1"
Case 2
    TEMP_MATRIX(5, 4) = "HC2"
Case Else
    TEMP_MATRIX(5, 4) = "HC3"
End Select
For i = 1 To NCOLUMNS: TEMP_MATRIX(5 + i, 4) = RSE_VECTOR(i, 1): Next i
Select Case OUTPUT
    Case 0
        REGRESSION_LS1_FUNC = TEMP_MATRIX
    Case 1
        REGRESSION_LS1_FUNC = COEF_VECTOR
    Case 2
        REGRESSION_LS1_FUNC = RESID_VECTOR
    Case 3
        REGRESSION_LS1_FUNC = SE_VECTOR
    Case 4
        REGRESSION_LS1_FUNC = HAT_VECTOR
    Case 5
        REGRESSION_LS1_FUNC = RSE_VECTOR
    Case Else
        REGRESSION_LS1_FUNC = Array(TEMP_MATRIX, COEF_VECTOR, RESID_VECTOR, SE_VECTOR, HAT_VECTOR, RSE_VECTOR)
End Select

Exit Function
ERROR_LABEL:
REGRESSION_LS1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : REGRESSION_LS2_FUNC

'DESCRIPTION   : Multiple regression Frame: We use the cholesky adjustment for the
'inverse of the matrix and the correction factor suggested by Davidson
'and MacKinnon (1993).

'Davidson and MacKinnon textbook (1993) p. 553 recommends using a correction factor
'in which one divides the estimate of the standard error by (1- HT) where HT is the
'square root of the t'th diagonal entry in the "hat matrix". This hat matrix is
'sometimes called P because it projects orthogonally onto the space spanned by the
'columns of X.

'Stata uses a much simpler correction namely sqrt(N/(N-j)).
'Davidson and MacKinnon (553-54) say that Stata's correction is inferior to dividing
'by (1-HT).

'LIBRARY       : STATISTICS
'GROUP         : REGRESSION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function REGRESSION_LS2_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = True, _
Optional ByVal SE_VERSION As Long = 2, _
Optional ByVal CI_VAL As Double = 0.95, _
Optional ByVal OUTPUT As Integer = 1)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim NO_VAR As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim MULT_VAL As Double

Dim PXT_VAL As Double
Dim PYT_VAL As Double
Dim PYX_VAL As Double

Dim YMEAN_VAL As Double
Dim YSTDEV_VAL As Double

Dim RMEAN_VAL As Double
Dim RSTDEVP_VAL As Double

Dim YFIT_VAL As Double

Dim SSR_VAL As Double
Dim DW_VAL As Double
Dim RMSE_VAL As Double
Dim TSS_VAL As Double
Dim YSQ_VAL As Double
Dim RSQ_VAL As Double
Dim FSTAT_VAL As Double
Dim MAPE_VAL As Double
Dim FACTOR_VAL As Double
Dim ERROR_STR As String

Dim P_ARR() As Double
Dim HT_ARR() As Double
Dim XT_ARR() As Double
Dim SE_ARR() As Double

Dim HAT_ARR() As Double
Dim RSE_ARR() As Double

Dim RESID1_ARR() As Double
Dim RESID2_ARR() As Double

Dim S_MATRIX() As Double
Dim T_MATRIX() As Double

Dim RSE_MATRIX() As Double
Dim COEF_VECTOR() As Double

Dim X_MATRIX() As Double
Dim XT_MATRIX() As Double
Dim XTX_MATRIX() As Double
Dim XTXI_MATRIX() As Double
Dim XTXIS_MATRIX() As Double
Dim XTXIXT_MATRIX() As Double
Dim XTXIXTX_MATRIX() As Double

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_MATRIX As Variant

Const PI_VAL As Double = 3.14159265358979

On Error GoTo ERROR_LABEL

ERROR_STR = ""
XDATA_MATRIX = XDATA_RNG
NO_VAR = UBound(XDATA_MATRIX, 2)
YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
'----------------------------------------------------------------------------------
NROWS = UBound(YDATA_VECTOR, 1)
If NROWS > UBound(XDATA_MATRIX, 1) Then: NROWS = UBound(XDATA_MATRIX, 1)
'----------------------------------------------------------------------------------
'tells us the number of parameters to estimate
If INTERCEPT_FLAG = True Then NCOLUMNS = NO_VAR + 1 Else NCOLUMNS = NO_VAR
GoSub INPUTS_LINE: GoSub COEF_LINE: GoSub OUTPUT_LINE

'---------------------------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To 7, 1 To 4)
    If INTERCEPT_FLAG = True Then
        j = NROWS - NO_VAR - 1 'Residual DF
        k = NROWS - j - 1 'Regression DF
    Else
        j = NROWS - NO_VAR 'Residual DF
        k = NROWS - j 'Regression DF
    End If
'--------------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(1, 1) = "R^2"
    TEMP_MATRIX(1, 2) = RSQ_VAL '1 - SSR_VAL / TSS_VAL
    
    TEMP_MATRIX(2, 1) = "RBar^2"
    'TEMP_MATRIX(2, 2) = 1 - (((1 - RSQ_VAL) * (NROWS - 1)) / (NROWS - k - 1))
    TEMP_MATRIX(2, 2) = 1 - ((SSR_VAL / (NROWS - k - 1)) / (TSS_VAL / (NROWS - 1)))
    'The Adjusted R-Squared is similar to the R-Squared, however, the Adjusted RSquared
    'takes into account the number of independent variables in the regression. The
    'Adjusted R-Squared is useful when comparing the fit of two equations with the same
    'dependent variable but a different number of explanatory variables
    
    'Johnston, Jack and John DiNardo (1997). Econometric Methods. New York The
    'McGraw-Hill Companies, Incorporated, pg 74.
'--------------------------------------------------------------------------------------------------------------------------------------
    
    TEMP_MATRIX(3, 1) = "RMSE" 'S.E. of regression
    TEMP_MATRIX(3, 2) = RMSE_VAL
    
    TEMP_MATRIX(4, 1) = "SSR" 'Sum squared resid
    TEMP_MATRIX(4, 2) = SSR_VAL
    
'--------------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(5, 1) = "F-test"
    TEMP_MATRIX(5, 2) = FSTAT_VAL

    TEMP_MATRIX(6, 1) = "Prob(F)"
    TEMP_MATRIX(6, 2) = FDIST_FUNC(FSTAT_VAL, k, j, True, False)
'--------------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(7, 1) = "S.D. Resids"
    TEMP_MATRIX(7, 2) = RSTDEVP_VAL
'--------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(1, 3) = "MAPE"
    TEMP_MATRIX(1, 4) = MAPE_VAL
'--------------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(2, 3) = "CV Regr"
    TEMP_MATRIX(2, 4) = RMSE_VAL / YMEAN_VAL * 100
    'The coefficient of variation for the regression is a measure of the average error relative to the actual mean of the
    'dependent variable
'--------------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(3, 3) = "Durbin-Watson"
    TEMP_MATRIX(3, 4) = DW_VAL
    'The Durbin-Watson test statistic is a measure of first-order autocorrelation in the model.
    'http://en.wikipedia.org/wiki/Durbin%E2%80%93Watson_statistic
'--------------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(4, 3) = "Rho"
    'The most common procedure for modeling a system with autocorrelation is a first-order autoregressive process or an AR(1).
    'In an AR(1) process the error in time t is lagged on the error in t-1 which yields the equation: et = p * et-1 + rt
    'et = Error term in time t from a regression model: y =Xb + e
    'p = Parameter rho that determines the properties of et
    'rt = Independent disturbances for the AR(1) process
    XTEMP_VECTOR = XDATA_MATRIX
    YTEMP_VECTOR = YDATA_VECTOR
    jj = NCOLUMNS: ii = NROWS
    NROWS = NROWS - 1: NCOLUMNS = 1: INTERCEPT_FLAG = False
    ReDim YDATA_VECTOR(1 To NROWS, 1 To 1)
    ReDim XDATA_MATRIX(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        YDATA_VECTOR(i, 1) = RESID1_ARR(i + 1)
        XDATA_MATRIX(i, 1) = RESID1_ARR(i + 0)
    Next i
    
    'parameter rho can be calculated from the regression equation y=Xb+e as:
    GoSub INPUTS_LINE: GoSub COEF_LINE
    TEMP_MATRIX(4, 4) = COEF_VECTOR(1, 1)
    XDATA_MATRIX = XTEMP_VECTOR
    YDATA_VECTOR = YTEMP_VECTOR
    NCOLUMNS = jj: NROWS = ii: INTERCEPT_FLAG = True
'--------------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(5, 3) = "Akaike Information Criterion"
    TEMP_MATRIX(5, 4) = Log(SSR_VAL / NROWS) + (2 * (NO_VAR) / NROWS)
    'The Akaike Information Criterion is used in the selection of regressors. A penalty
    'for increasing the number of regressors is added to a transformation of the minimum
    'residual sum of squares. The Akaike Information Criterion is calculated as follows
    'http://en.wikipedia.org/wiki/Akaike_information_criterion
    'Johnston, Jack and John DiNardo (1997). Econometric Methods. New York The McGraw-Hill Companies, Incorporated, pg 74.
'--------------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(6, 3) = "Schwarz Information Criterion"
    TEMP_MATRIX(6, 4) = Log(SSR_VAL / NROWS) + (NO_VAR / NROWS) * Log(NROWS)
    'The Schwarz Criterion is used in the selection of lags for an AR(p) process. A
    'penalty for increasing the number of lags is added to a transformation of the minimum
    'residual sum of squares.
    'Johnston, Jack and John DiNardo (1997). Econometric Methods. New York The McGraw-Hill Companies, Incorporated, pg 74.
    'http://en.wikipedia.org/wiki/Bayesian_information_criterion
    'http://en.wikipedia.org/wiki/Hannan%E2%80%93Quinn_information_criterion
    'http://en.wikipedia.org/wiki/Newey%E2%80%93West_estimator
'--------------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(7, 3) = "Log Likelihood"
    TEMP_MATRIX(7, 4) = -(NROWS / 2) * Log(2 * PI_VAL) - (NROWS / 2) * Log(SSR_VAL / NROWS) - (NROWS / 2)
    'If INTERCEPT_FLAG = True Then k = NO_VAR + 1 else k = NO_VAR
    'Akaike: -2 * (LLIKE_VAL / NROWS) + ((2 * k) / NROWS)
    'Schwarz: = -2 * (LLIKE_VAL / NROWS) + ((k * Log(NROWS)) / NROWS)
    'http://en.wikipedia.org/wiki/Likelihood-ratio_test
'--------------------------------------------------------------------------------------------------------------------------------------
Case 1
'--------------------------------------------------------------------------------------------------------------------------------------
    GoSub RSE_LINE
    ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 8)
    TEMP_MATRIX(0, 1) = "Heading"
    TEMP_MATRIX(0, 2) = "Coefficient"
    TEMP_MATRIX(0, 3) = "HC" & SE_VERSION
    TEMP_MATRIX(0, 4) = "S.E."
    TEMP_MATRIX(0, 5) = "t-test"
    TEMP_MATRIX(0, 6) = "Prob(t)"
    TEMP_MATRIX(0, 7) = "Elasticity at Mean"
    TEMP_MATRIX(0, 8) = "Variance Inflation Factor"
    'In statistics, the variance inflation factor (VIF) quantifies the severity of multicollinearity
    'in an ordinary least squares regression analysis. It provides an index that measures how much the
    'variance of an estimated regression coefficient (the square of the estimate's standard deviation) is
    'increased because of collinearity.
    
    'TEMP_MATRIX(0, 9) = "Partial Correlation"
    'TEMP_MATRIX(0, 10) = "Semipartial Correlation"
    If INTERCEPT_FLAG = True Then
        k = NROWS - NO_VAR - 1
        j = 1: TEMP_MATRIX(j, 1) = "Alpha"
        TEMP_MATRIX(j, 7) = ""
        For j = 2 To NCOLUMNS
            TEMP_MATRIX(j, 1) = "Beta: " & j - 1
            TEMP_MATRIX(j, 7) = 0: For i = 1 To NROWS: TEMP_MATRIX(j, 7) = TEMP_MATRIX(j, 7) + X_MATRIX(i, j): Next i
            TEMP_MATRIX(j, 7) = COEF_VECTOR(j, 1) * (TEMP_MATRIX(j, 7) / NROWS) / YMEAN_VAL
        Next j
    Else
        k = NROWS - NO_VAR 'Residual DF
        For j = 1 To NCOLUMNS
            TEMP_MATRIX(j, 1) = "Beta: " & j
            TEMP_MATRIX(j, 7) = 0: For i = 1 To NROWS: TEMP_MATRIX(j, 7) = TEMP_MATRIX(j, 7) + X_MATRIX(i, j): Next i
            TEMP_MATRIX(j, 7) = COEF_VECTOR(j, 1) * (TEMP_MATRIX(j, 7) / NROWS) / YMEAN_VAL
        Next j
    End If
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(j, 2) = COEF_VECTOR(j, 1)
        TEMP_MATRIX(j, 3) = RSE_ARR(j)
        TEMP_MATRIX(j, 4) = SE_ARR(j)
        TEMP_MATRIX(j, 5) = TEMP_MATRIX(j, 2) / TEMP_MATRIX(j, 4)
        TEMP_MATRIX(j, 6) = 2 * (1 - TDIST_FUNC(Abs(TEMP_MATRIX(j, 5)), k, True))
    Next j
    If NO_VAR > 1 Then 'http://en.wikipedia.org/wiki/Variance_inflation_factor
        NO_VAR = NO_VAR - 1 'UBound(XDATA_MATRIX, 2) - 1
        If INTERCEPT_FLAG = True Then NCOLUMNS = NO_VAR + 1 Else NCOLUMNS = NO_VAR
        XTEMP_VECTOR = XDATA_MATRIX
        For h = 1 To NO_VAR + 1
            ReDim XDATA_MATRIX(1 To NROWS, 1 To NO_VAR)
            ReDim YDATA_VECTOR(1 To NROWS, 1 To 1)
            For i = 1 To NROWS
                YDATA_VECTOR(i, 1) = XTEMP_VECTOR(i, h)
                k = 1
                For j = 1 To NO_VAR + 1
                    If j = h Then GoTo 1985
                    XDATA_MATRIX(i, k) = XTEMP_VECTOR(i, j)
                    k = k + 1
1985:
                Next j
            Next i
            If INTERCEPT_FLAG = True Then
                If h = 1 Then: TEMP_MATRIX(h, 8) = ""
                GoSub INPUTS_LINE: GoSub COEF_LINE: GoSub OUTPUT_LINE
                TEMP_MATRIX(h + 1, 8) = 1 / (1 - RSQ_VAL)
            Else 'Revise This
                INTERCEPT_FLAG = True: GoSub INPUTS_LINE: GoSub COEF_LINE: GoSub OUTPUT_LINE: INTERCEPT_FLAG = False
                TEMP_MATRIX(h, 8) = 1 / (1 - RSQ_VAL)
            End If
        Next h
    Else
        For j = 1 To NCOLUMNS: TEMP_MATRIX(j, 8) = "": Next j
    End If
'    TEMP_MATRIX = YDATA_VECTOR
    'http://www.dss.uniud.it/utenti/rizzi/econometrics_part2_file/heteroUK2.pdf
'---------------------------------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To UBound(XDATA_MATRIX, 1), 1 To 9)
    TEMP_MATRIX(0, 1) = "Actual Y"
    TEMP_MATRIX(0, 2) = "Predicted Y"
    TEMP_MATRIX(0, 3) = "Residuals"
    TEMP_MATRIX(0, 4) = "SE Mean Predicted Y"
    TEMP_MATRIX(0, 5) = "SE Predicted Y"
    TEMP_MATRIX(0, 6) = "Lower " & Format(CI_VAL, "0%") & " Conf. Interval"
    TEMP_MATRIX(0, 7) = "Upper " & Format(CI_VAL, "0%") & " Conf. Interval"
    TEMP_MATRIX(0, 8) = "Lower " & Format(CI_VAL, "0%") & " Predict. Interval"
    TEMP_MATRIX(0, 9) = "Upper " & Format(CI_VAL, "0%") & " Predict. Interval"
    '-----------------------------------------------------------------------------------------------
    If INTERCEPT_FLAG = True Then j = NROWS - NO_VAR - 1 Else j = NROWS - NO_VAR 'Residual DF
    FACTOR_VAL = -INVERSE_TDIST_FUNC((1 - CI_VAL) / 2, j)
    'The least squares projection matrix or the hat matrix determines the predicted
    'values of a regression model. The diagonal elements of the hat matrix or Leverage can
    'be used to measure the effect that the individual observations of the dependant variable
    'have on the corresponding estimation of that observation.
    For i = 1 To NROWS
        TEMP_MATRIX(i, 4) = RMSE_VAL * XTXIXTX_MATRIX(i, i) ^ 0.5 'Diagonal of Hat Matrix
        TEMP_MATRIX(i, 5) = RMSE_VAL * (1 + XTXIXTX_MATRIX(i, i)) ^ 0.5
        
        TEMP_MATRIX(i, 1) = YDATA_VECTOR(i, 1)
        TEMP_MATRIX(i, 3) = RESID1_ARR(i)
        TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 1) - TEMP_MATRIX(i, 3)
        
        MULT_VAL = FACTOR_VAL * TEMP_MATRIX(i, 4)
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 2) - MULT_VAL
        TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 2) + MULT_VAL
        
        MULT_VAL = FACTOR_VAL * RMSE_VAL * (1 + (TEMP_MATRIX(i, 4) / RMSE_VAL) ^ 2) ^ 0.5
        TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 2) - MULT_VAL
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 2) + MULT_VAL
    Next i
    If UBound(XDATA_MATRIX, 1) > NROWS Then 'Projected Values for the Exogenous Variables
        h = UBound(YDATA_VECTOR, 1): NROWS = UBound(XDATA_MATRIX, 1): GoSub INPUTS_LINE
        For i = h + 1 To NROWS
            TEMP_MATRIX(i, 1) = CVErr(xlErrNA) 'NORMSINV_FUNC(Rnd(), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 5), 0)
            TEMP_MATRIX(i, 3) = CVErr(xlErrNA) '""
            ReDim XTEMP_VECTOR(1 To NCOLUMNS)
            For j = 1 To NCOLUMNS 'Thanks to Dr. Zaric!!!
                XTEMP_VECTOR(j) = 0: For k = 1 To NCOLUMNS: XTEMP_VECTOR(j) = XTEMP_VECTOR(j) + X_MATRIX(i, k) * XTXI_MATRIX(k, j): Next k
            Next j
            TEMP_MATRIX(i, 2) = 0: TEMP_MATRIX(i, 5) = 0
            For j = 1 To NCOLUMNS
                TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 2) + COEF_VECTOR(j, 1) * X_MATRIX(i, j)
                TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 5) + XTEMP_VECTOR(j) * XT_MATRIX(j, i)
            Next j
            'TEMP_MATRIX(i, 4) = (TEMP_MATRIX(i, 5) ^ 2 - RMSE_VAL ^ 2) ^ 0.5
            TEMP_MATRIX(i, 4) = RMSE_VAL * TEMP_MATRIX(i, 5) ^ 0.5
            TEMP_MATRIX(i, 5) = RMSE_VAL * (1 + TEMP_MATRIX(i, 5)) ^ 0.5
            
            MULT_VAL = FACTOR_VAL * TEMP_MATRIX(i, 4)
            TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 2) - MULT_VAL
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 2) + MULT_VAL
            
            MULT_VAL = FACTOR_VAL * RMSE_VAL * (1 + (TEMP_MATRIX(i, 4) / RMSE_VAL) ^ 2) ^ 0.5
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 2) - MULT_VAL
            TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 2) + MULT_VAL
        Next i
    End If
'---------------------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------------------

REGRESSION_LS2_FUNC = TEMP_MATRIX

'--------------------------------------------------------------------------------------------------------------------------------
Exit Function
'--------------------------------------------------------------------------------------------------------------------------------
INPUTS_LINE:
'--------------------------------------------------------------------------------------------------------------------------------
    ReDim X_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    ReDim XT_MATRIX(1 To NCOLUMNS, 1 To NROWS)
    If INTERCEPT_FLAG = True Then
        For i = 1 To NROWS
            j = 1: X_MATRIX(i, j) = 1: XT_MATRIX(j, i) = 1
            For j = 2 To NCOLUMNS
                X_MATRIX(i, j) = XDATA_MATRIX(i, j - 1)
                XT_MATRIX(j, i) = X_MATRIX(i, j)
            Next j
        Next i
    Else
        For i = 1 To NROWS
            For j = 1 To NCOLUMNS
                X_MATRIX(i, j) = XDATA_MATRIX(i, j)
                XT_MATRIX(j, i) = X_MATRIX(i, j)
            Next j
        Next i
    End If
'--------------------------------------------------------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------------------------------------------------------
COEF_LINE:
'--------------------------------------------------------------------------------------------------------------------------------
    ReDim T_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
    ReDim XTX_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
    For i = 1 To NCOLUMNS
        For j = 1 To NCOLUMNS
            T_MATRIX(i, j) = 0
            XTX_MATRIX(i, j) = 0
            For k = 1 To NROWS
                T_MATRIX(i, j) = XT_MATRIX(i, k) * X_MATRIX(k, j) + T_MATRIX(i, j)
                XTX_MATRIX(i, j) = XT_MATRIX(i, k) * X_MATRIX(k, j) + XTX_MATRIX(i, j)
            Next k
        Next j
    Next i
    ReDim P_ARR(1 To NCOLUMNS): For i = 1 To NCOLUMNS: P_ARR(i) = i: Next i
    If REGRESSION_LS_MC_TEST_FUNC(T_MATRIX, P_ARR, NCOLUMNS) = False Then
        ERROR_STR = "There is perfect or near-perfect multicollinearity in the independent variables. Thus the regression fails."
        GoTo ERROR_LABEL
    End If
    
    ReDim XTXIXT_MATRIX(1 To NCOLUMNS, 1 To NROWS)
    If NCOLUMNS = 1 Then
        ReDim XTXI_MATRIX(1 To 1)
        XTXI_MATRIX(1) = 1 / XTX_MATRIX(1, 1)
        For i = 1 To NROWS: XTXIXT_MATRIX(1, i) = XTXI_MATRIX(1) * XT_MATRIX(1, i): Next i
    Else
        XTXI_MATRIX = MATRIX_CHOLESKY_INVERSE_FUNC(XTX_MATRIX, 0, True)
        For i = 1 To NCOLUMNS
            For j = 1 To NROWS
                XTXIXT_MATRIX(i, j) = 0
                For k = 1 To NCOLUMNS: XTXIXT_MATRIX(i, j) = XTXI_MATRIX(i, k) * XT_MATRIX(k, j) + XTXIXT_MATRIX(i, j): Next k
            Next j
        Next i
    End If

    ReDim XTXIXTX_MATRIX(1 To NROWS, 1 To NROWS) 'Halt Matrix
    For i = 1 To NROWS
        For j = 1 To NROWS
            XTXIXTX_MATRIX(i, j) = 0
            For k = 1 To NCOLUMNS: XTXIXTX_MATRIX(i, j) = XTXIXTX_MATRIX(i, j) + X_MATRIX(i, k) * XTXIXT_MATRIX(k, j): Next k
        Next j
    Next i
    ReDim COEF_VECTOR(1 To NCOLUMNS, 1 To 1)
    For j = 1 To NCOLUMNS
        COEF_VECTOR(j, 1) = 0
        For i = 1 To NROWS: COEF_VECTOR(j, 1) = COEF_VECTOR(j, 1) + XTXIXT_MATRIX(j, i) * YDATA_VECTOR(i, 1): Next i
    Next j

'--------------------------------------------------------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------------------------------------------------------
OUTPUT_LINE:
'--------------------------------------------------------------------------------------------------------------------------------
    YMEAN_VAL = 0: For i = 1 To NROWS: YMEAN_VAL = YMEAN_VAL + YDATA_VECTOR(i, 1): Next i
    YMEAN_VAL = YMEAN_VAL / NROWS
    ReDim RESID1_ARR(1 To NROWS): ReDim RESID2_ARR(1 To NROWS)
    YSTDEV_VAL = 0
    RMEAN_VAL = 0: DW_VAL = 0: MAPE_VAL = 0
    RMSE_VAL = 0: TSS_VAL = 0: YSQ_VAL = 0
    For i = 1 To NROWS
        YSTDEV_VAL = YSTDEV_VAL + (YDATA_VECTOR(i, 1) - YMEAN_VAL) ^ 2
        YFIT_VAL = 0
        TSS_VAL = TSS_VAL + (YDATA_VECTOR(i, 1) - YMEAN_VAL) ^ 2
        For j = 1 To NCOLUMNS: YFIT_VAL = YFIT_VAL + COEF_VECTOR(j, 1) * X_MATRIX(i, j): Next j
        RESID1_ARR(i) = YDATA_VECTOR(i, 1) - YFIT_VAL
        RMEAN_VAL = RMEAN_VAL + RESID1_ARR(i)
        If YDATA_VECTOR(i, 1) <> 0 Then: MAPE_VAL = MAPE_VAL + Abs(RESID1_ARR(i) / YDATA_VECTOR(i, 1))
        RESID2_ARR(i) = RESID1_ARR(i) ^ 2
        RMSE_VAL = RMSE_VAL + RESID2_ARR(i)
        YSQ_VAL = YSQ_VAL + YDATA_VECTOR(i, 1) ^ 2
        If i > 1 Then: DW_VAL = DW_VAL + (RESID1_ARR(i) - RESID1_ARR(i - 1)) ^ 2
    Next i
    YSTDEV_VAL = (YSTDEV_VAL / (NROWS - 1)) ^ 0.5
    RMEAN_VAL = RMEAN_VAL / NROWS
    MAPE_VAL = MAPE_VAL / NROWS * 100
    RSTDEVP_VAL = 0
    For i = 1 To NROWS: RSTDEVP_VAL = RSTDEVP_VAL + (RESID1_ARR(i) - RMEAN_VAL) ^ 2: Next i
    RSTDEVP_VAL = (RSTDEVP_VAL / (NROWS - 0)) ^ 0.5
    SSR_VAL = RMSE_VAL
    If INTERCEPT_FLAG = False Then
        TSS_VAL = YSQ_VAL
        FSTAT_VAL = ((TSS_VAL - SSR_VAL) / (NCOLUMNS)) / (SSR_VAL / (NROWS - NCOLUMNS))
    Else
        FSTAT_VAL = ((TSS_VAL - SSR_VAL) / (NCOLUMNS - 1)) / (SSR_VAL / (NROWS - NCOLUMNS))
    End If
    RSQ_VAL = 1 - (SSR_VAL / TSS_VAL)
    DW_VAL = DW_VAL / SSR_VAL
    RMSE_VAL = (RMSE_VAL / (NROWS - NCOLUMNS)) ^ 0.5
    ReDim SE_ARR(1 To NCOLUMNS)
    If NCOLUMNS = 1 Then
        SE_ARR(1) = XTXI_MATRIX(1) ^ 0.5 * RMSE_VAL
    Else
        For i = 1 To NCOLUMNS
            SE_ARR(i) = XTXI_MATRIX(i, i)
            SE_ARR(i) = SE_ARR(i) ^ 0.5
            SE_ARR(i) = SE_ARR(i) * RMSE_VAL
        Next i
    End If
    
'--------------------------------------------------------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------------------------------------------------------
RSE_LINE: ' Get S matrix
' It has the same dimensions as XTX
' First , following Davidson and MacKinnnon, p. 553,  divide each ESQ by (1-HT)
' to get HT_ARR, multiply the i'th row of X_MATRIX into X_MATRIX'XInverse
' and the transpose of the i'th row of X_MATRIX
'--------------------------------------------------------------------------------------------------------------------------------
    ReDim HT_ARR(1 To NROWS)
    ReDim HAT_ARR(1 To NROWS)
    ' robust SE using Davidson and MacKinnon's various approaches (p. 553) implemented here
    Select Case SE_VERSION
    Case 0
        For i = 1 To NROWS: HAT_ARR(i) = RESID2_ARR(i): Next i
    Case 1
        For i = 1 To NROWS: HAT_ARR(i) = RESID2_ARR(i) * (NROWS / (NROWS - NCOLUMNS)): Next i
    Case Else
        ReDim XT_ARR(1 To NCOLUMNS)
        For h = 1 To NROWS
            For k = 1 To NCOLUMNS: XT_ARR(k) = X_MATRIX(h, k): Next k
            If NCOLUMNS = 1 Then
                MULT_VAL = XT_ARR(1) * XTXI_MATRIX(1) * XT_ARR(1)
            Else
                ReDim YTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
                ReDim XTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
                For i = 1 To NCOLUMNS: XTEMP_VECTOR(i, 1) = XT_ARR(i): Next i
                For i = 1 To NCOLUMNS
                    For j = 1 To 1
                        YTEMP_VECTOR(i, j) = 0
                        For k = 1 To NCOLUMNS: YTEMP_VECTOR(i, j) = XTXI_MATRIX(i, k) * XTEMP_VECTOR(k, j) + YTEMP_VECTOR(i, j): Next k
                    Next j
                Next i
                MULT_VAL = 0
                For i = 1 To NCOLUMNS: MULT_VAL = MULT_VAL + YTEMP_VECTOR(i, 1) * XT_ARR(i): Next i
            End If
            HT_ARR(h) = MULT_VAL
            If HT_ARR(h) = 1 Then
                HAT_ARR(h) = 0
            Else
                If SE_VERSION = 2 Then
                    HAT_ARR(h) = RESID2_ARR(h) / (1 - HT_ARR(h))
                Else
                    HAT_ARR(h) = RESID2_ARR(h) / ((1 - HT_ARR(h)) ^ 2)
                End If
            End If
        Next h
    End Select
    
    ReDim S_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        For k = 1 To NCOLUMNS
            TEMP_SUM = 0
            For h = 1 To NROWS
                S_MATRIX(j, k) = X_MATRIX(h, j) * X_MATRIX(h, k) * HAT_ARR(h) + TEMP_SUM
                TEMP_SUM = S_MATRIX(j, k)
            Next h
        Next k
    Next j
    
    ReDim XTXIS_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
    If NCOLUMNS = 1 Then
        XTXIS_MATRIX(1, 1) = XTXI_MATRIX(1) * S_MATRIX(1, 1)
    Else
        For i = 1 To NCOLUMNS
            For j = 1 To NCOLUMNS
                XTXIS_MATRIX(i, j) = 0
                For k = 1 To NCOLUMNS: XTXIS_MATRIX(i, j) = XTXI_MATRIX(i, k) * S_MATRIX(k, j) + XTXIS_MATRIX(i, j): Next k
            Next j
        Next i
    End If
    ReDim RSE_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
    If NCOLUMNS = 1 Then
        RSE_MATRIX(1, 1) = XTXIS_MATRIX(1, 1) * XTXI_MATRIX(1)
    Else
        For i = 1 To NCOLUMNS
            For j = 1 To NCOLUMNS
                RSE_MATRIX(i, j) = 0
                For k = 1 To NCOLUMNS: RSE_MATRIX(i, j) = XTXIS_MATRIX(i, k) * XTXI_MATRIX(k, j) + RSE_MATRIX(i, j): Next k
            Next j
        Next i
    End If
    
    ReDim RSE_ARR(1 To NCOLUMNS)
    For i = 1 To NCOLUMNS
        If RSE_MATRIX(i, i) < 0 Then: RSE_MATRIX(i, i) = 0
        RSE_ARR(i) = RSE_MATRIX(i, i) ^ 0.5
    Next i
'------------------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
'------------------------------------------------------------------------------------------------------------------------
If ERROR_STR = "" Then
    REGRESSION_LS2_FUNC = Err.number
Else
    REGRESSION_LS2_FUNC = ERROR_STR
End If
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : REGRESSION_LS_MC_TEST_FUNC

'DESCRIPTION   : Test for perfect or near-perfect multicollinearity
'in the independent variables

'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_CHOLESKY
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function REGRESSION_LS_MC_TEST_FUNC(ByRef DATA_MATRIX As Variant, _
ByRef DATA_ARR As Variant, _
Optional ByVal NSIZE As Long = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_SUM As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 0.0000001
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL
If NSIZE = 0 Then: NSIZE = UBound(DATA_MATRIX, 1)
If IsArray(DATA_ARR) = False Then: ReDim DATA_ARR(1 To NSIZE)

REGRESSION_LS_MC_TEST_FUNC = True
For i = 1 To NSIZE
    For j = 1 To NSIZE
        TEMP_SUM = DATA_MATRIX(i, j)
        k = i - 1
        Do While k > 0
            TEMP_SUM = TEMP_SUM - DATA_MATRIX(i, k) * DATA_MATRIX(j, k)
            k = k - 1
        Loop
        If i = j Then
             If TEMP_SUM < tolerance Then
                REGRESSION_LS_MC_TEST_FUNC = False
    'One reason for the test to fail is that there may be perfect or
    'near-perfect multicollinearity in the independent variables. Thus
    'the regression fails.
                Exit Function
             End If
             DATA_ARR(i) = TEMP_SUM ^ 0.5
        Else
             DATA_MATRIX(j, i) = TEMP_SUM / DATA_ARR(i)
        End If
    Next j
Next i

Exit Function
ERROR_LABEL:
REGRESSION_LS_MC_TEST_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : REGRESSION_INPUTS1_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_DATA
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function REGRESSION_INPUTS1_FUNC( _
ByRef XDATA_RNG As Excel.Range, _
ByRef YDATA_RNG As Excel.Range)

Dim h As Long  'searching through cells in an area
Dim i As Long
Dim j As Long
Dim k As Long

Dim hh As Long
Dim ii As Long ' keep track of which X areas have problems
Dim jj As Long  '
Dim kk As Long ' index for x variables
Dim ll As Long  ' count the number of valid observations

Dim NSIZE As Long
Dim NO_X_ROWS As Long

Dim X_NROWS As Long
Dim Y_NROWS As Long

Dim X_NCOLUMNS As Long
Dim Y_NCOLUMNS As Long

Dim ROWS_ARR() As Long    ' keep track of how many obs in each area
Dim COLUMNS_ARR() As Long ' keep track of how many x variables in each area

Dim ERROR_STR As String
Dim X_VAR_LABEL_ARR() As String
Dim Y_VAR_LABEL_STR As String

Dim YTEMP_VECTOR() As Double
Dim XTEMP_MATRIX() As Double 'The actual values go here

Dim ERROR_MATCH_FLAG As Boolean  'indicator for mismatch between no of obs in X cols

On Error GoTo ERROR_LABEL

ERROR_STR = ""
NSIZE = XDATA_RNG.Areas.COUNT
X_NCOLUMNS = 0
ERROR_MATCH_FLAG = False
ReDim COLUMNS_ARR(1 To NSIZE)
ReDim ROWS_ARR(1 To NSIZE)
For i = 1 To NSIZE
    COLUMNS_ARR(i) = XDATA_RNG.Areas(i).Columns.COUNT
    ROWS_ARR(i) = XDATA_RNG.Areas(i).Rows.COUNT
    X_NCOLUMNS = X_NCOLUMNS + COLUMNS_ARR(i)
    If i > 1 Then
        If ROWS_ARR(i) <> ROWS_ARR(i - 1) Then
            ERROR_MATCH_FLAG = True
            ii = i - 1
            jj = i
        End If
    End If
Next i
If X_NCOLUMNS > 51 Then
    ERROR_STR = "Unfortunately, this function cannot handle more than 51 independent variables.  You've selected " & X_NCOLUMNS & ". Sorry!"
    GoTo ERROR_LABEL
End If
' Warning if ERROR_MATCH_FLAG is true
If ERROR_MATCH_FLAG = True Then
    ERROR_STR = "The number of rows in X area " & ii & "does not equal the number of observations in X area " & jj & ". Please try again."
    GoTo ERROR_LABEL
End If
' go through cells in each area (determine how many there will be)
' labels need to be found
X_NROWS = ROWS_ARR(1) - 1
ReDim X_VAR_LABEL_ARR(1 To X_NCOLUMNS) As String
Y_NROWS = YDATA_RNG.Rows.COUNT - 1
Y_NCOLUMNS = YDATA_RNG.Columns.COUNT
If X_NROWS <> Y_NROWS Then
    ERROR_STR = "You must select the same number of rows for both the X variable(s) and the Y variable. Please try again."
    GoTo ERROR_LABEL
End If
' Check that we have just one Y column
If Y_NCOLUMNS > 1 Then
    ERROR_STR = "You must select only one column for the Y variable. Please try again."
    GoTo ERROR_LABEL
End If
' Check on labels
NO_X_ROWS = X_NROWS
ReDim XTEMP_MATRIX(1 To NO_X_ROWS, 1 To X_NCOLUMNS)
ReDim YTEMP_VECTOR(1 To NO_X_ROWS, 1 To 1)
kk = 0
For i = 1 To NSIZE
    For j = 1 To COLUMNS_ARR(i)
        kk = kk + 1
        h = j
        X_VAR_LABEL_ARR(kk) = XDATA_RNG.Areas(i).Cells(h)
        If IsNumeric(X_VAR_LABEL_ARR(kk)) = True Then
            hh = MsgBox("The X variable label in column " & kk & " you've chosen is a number.  Do you really want the variable label to be " & X_VAR_LABEL_ARR(kk) & "?", vbYesNo, Title:="Potential Label Problem")
            If hh = vbNo Then GoTo ERROR_LABEL
        End If
    Next j
Next i

Y_VAR_LABEL_STR = YDATA_RNG(1)
If IsNumeric(Y_VAR_LABEL_STR) = True Then
    hh = MsgBox("The Y variable label you've chosen is a number. Do you really want the variable label to be " & Y_VAR_LABEL_STR & "?", vbYesNo, Title:="Potential Label Problem")
    If hh = vbNo Then GoTo ERROR_LABEL
End If
    
' Start reading the data
' Must read in one SROW at a time across Y variable and X variables
' Data is assumed to be in columnar format!
ll = 0
For i = 1 To NO_X_ROWS
    On Error GoTo 1982
    ' Read y data first
    ll = ll + 1
     'remember first SROW is label so must add one
    ' We are sent to error handling if this isn't a number
    ' Now check for blanks
    YTEMP_VECTOR(ll, 1) = YDATA_RNG(i + 1, 1)
    If IsEmpty(YDATA_RNG(i + 1, 1)) = True Then
        ll = ll - 1 ' we are going to skip this obs.
        GoTo 1983
    End If
    ' If we've passed, go to the x variables
            kk = 0
    For j = 1 To NSIZE
            
        For k = 1 To COLUMNS_ARR(j)
            h = i * COLUMNS_ARR(j) + k
            kk = kk + 1
            On Error GoTo 1982
     '   XTEMP_MATRIX(ll, j) = XDATA_RNG(i + 1, j)
     '   Check for empty values
            If IsEmpty(XDATA_RNG.Areas(j).Cells(h)) = True Then
                ll = ll - 1
                GoTo 1983
            Else
                XTEMP_MATRIX(ll, kk) = XDATA_RNG.Areas(j).Cells(h)
            End If
        Next k
    Next j
    GoTo 1983
1982:
    ll = ll - 1
    'Resume 1983
1983:
Next i
' End reading in data
If ll < X_NCOLUMNS Then
    ERROR_STR = "There aren't enough observations with non-missing values to obtain parameter estimates.  Try again."
    GoTo ERROR_LABEL
End If

REGRESSION_INPUTS1_FUNC = Array(XTEMP_MATRIX(), YTEMP_VECTOR(), X_NCOLUMNS, NO_X_ROWS)

Exit Function
ERROR_LABEL:
If ERROR_STR = "" Then
    REGRESSION_INPUTS1_FUNC = Err.number
Else
    REGRESSION_INPUTS1_FUNC = ERROR_STR
End If
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REGRESSION_INPUTS2_FUNC

'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_DATA
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function REGRESSION_INPUTS2_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long  ' count the number of valid observations

Dim X_NROWS As Long
Dim Y_NROWS As Long

Dim X_NCOLUMNS As Long
Dim Y_NCOLUMNS As Long

Dim ERROR_STR As String

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant

Dim YTEMP_VECTOR() As Double
Dim XTEMP_MATRIX() As Double 'The actual values go here

On Error GoTo ERROR_LABEL

YDATA_VECTOR = YDATA_RNG
XDATA_MATRIX = XDATA_RNG

ERROR_STR = ""
'---------------------------------------------------------------------------------
X_NCOLUMNS = UBound(XDATA_MATRIX, 2)
'---------------------------------------------------------------------------------
If X_NCOLUMNS > 51 Then
    ERROR_STR = "Unfortunately, this function cannot handle more than 51 independent variables.  You've selected " & X_NCOLUMNS & ". Sorry!"
    GoTo ERROR_LABEL
End If
'---------------------------------------------------------------------------------
' go through cells in each area (determine how many there will be)
' labels need to be found
X_NROWS = UBound(XDATA_MATRIX, 1) - 1
Y_NROWS = UBound(YDATA_VECTOR, 1) - 1
Y_NCOLUMNS = UBound(YDATA_VECTOR, 2)

If X_NROWS <> Y_NROWS Then
    ERROR_STR = "You must select the same number of rows for both the X variable(s) and the Y variable. Please try again."
    GoTo ERROR_LABEL
End If
'---------------------------------------------------------------------------------
' Check that we have just one Y column
If Y_NCOLUMNS > 1 Then
    ERROR_STR = "You must select only one column for the Y variable. Please try again."
    GoTo ERROR_LABEL
End If

'---------------------------------------------------------------------------------
' Check on labels
ReDim XTEMP_MATRIX(1 To X_NROWS, 1 To X_NCOLUMNS)
ReDim YTEMP_VECTOR(1 To X_NROWS, 1 To 1)
'---------------------------------------------------------------------------------
For j = 1 To X_NCOLUMNS
    If IsNumeric(XDATA_MATRIX(1, j)) = True Then
        ERROR_STR = "The X variable label in column " & j & " you've chosen is a number. Please try again."
        'Potential Label Problem
        GoTo ERROR_LABEL
    End If
Next j

'---------------------------------------------------------------------------------
If IsNumeric(YDATA_VECTOR(1, 1)) = True Then
    ERROR_STR = "The Y variable label you've chosen is a number. Please try again." 'Potential Label Problem
    GoTo ERROR_LABEL
End If
'---------------------------------------------------------------------------------
' Start reading the data
' Must read in one SROW at a time across Y variable and X variables
' Data is assumed to be in columnar format!
'---------------------------------------------------------------------------------
k = 0
For i = 1 To X_NROWS
    On Error GoTo 1982
    ' Read y data first
    k = k + 1
     'remember first SROW is label so must add one
    ' We are sent to error handling if this isn't a number
    ' Now check for blanks
    YTEMP_VECTOR(k, 1) = YDATA_VECTOR(i + 1, 1)
    If IsEmpty(YDATA_VECTOR(i + 1, 1)) = True Then
        k = k - 1 ' we are going to skip this obs.
        GoTo 1983
    End If
    ' If we've passed, go to the x variables
    For j = 1 To X_NCOLUMNS
        On Error GoTo 1982
     '   Check for empty values
        If IsEmpty(XDATA_MATRIX(i + 1, j)) = True Then
            k = k - 1
            GoTo 1983
        Else
            XTEMP_MATRIX(k, j) = XDATA_MATRIX(i + 1, j)
        End If
    Next j
    GoTo 1983
1982:
    k = k - 1
'Resume 1983
1983:
Next i

'---------------------------------------------------------------------------------
' End reading in data
If k < X_NCOLUMNS Then
    ERROR_STR = "There aren't enough observations with non-missing values to obtain parameter estimates. Try again."
    GoTo ERROR_LABEL
End If
'---------------------------------------------------------------------------------

REGRESSION_INPUTS2_FUNC = Array(XTEMP_MATRIX(), YTEMP_VECTOR(), k, X_NCOLUMNS, X_NROWS)

Exit Function
ERROR_LABEL:
If ERROR_STR = "" Then
    REGRESSION_INPUTS2_FUNC = Err.number
Else
    REGRESSION_INPUTS2_FUNC = ERROR_STR
End If
End Function

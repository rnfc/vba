Attribute VB_Name = "FINAN_ASSET_MOMENTS_VAR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Suggested Literature: "CoVaR", Tobias Adrian & Markus K. Brunnermeier, September 2008, Federal
'Reserve Bank of New York Staff Report

Function ASSET_PREDICTED_COVAR_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 0.99, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim HIST1_VAR As Double
Dim HIST2_VAR As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim CUMUL1_SUM As Double
Dim CUMUL2_SUM As Double

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'----------------------------------------------------------------------
DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then
    DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)
End If
If DATA_TYPE <> 0 Then
    DATA1_VECTOR = MATRIX_PERCENT_FUNC(DATA1_VECTOR, 0)
End If
NROWS = UBound(DATA1_VECTOR, 1)
'----------------------------------------------------------------------
DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then
    DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
End If
If DATA_TYPE <> 0 Then
    DATA2_VECTOR = MATRIX_PERCENT_FUNC(DATA2_VECTOR, 0)
End If
'----------------------------------------------------------------------
If UBound(DATA2_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------
HIST1_VAR = HISTOGRAM_PERCENTILE_FUNC(DATA1_VECTOR, 1 - CONFIDENCE_VAL, 1)
'HIST1_VAR = WorksheetFunction.PERCENTILE(DATA1_VECTOR, 1 - CONFIDENCE)
HIST2_VAR = HISTOGRAM_PERCENTILE_FUNC(DATA2_VECTOR, 1 - CONFIDENCE_VAL, 1)
'HIST2_VAR = WorksheetFunction.PERCENTILE(DATA2_VECTOR, 1 - CONFIDENCE)

'----------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------
Case 0 'Historical VaR
'----------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = 1
        TEMP_MATRIX(i, 2) = DATA1_VECTOR(i, 1)
    Next i
    TEMP_MATRIX = QUANTILE_REGRESSION_FUNC(TEMP_MATRIX, DATA2_VECTOR, 1 - CONFIDENCE_VAL, , 1000, 10 ^ -10, 2)
    HIST1_VAR = TEMP_MATRIX(1, 1) + TEMP_MATRIX(2, 1) * HIST1_VAR
    'Absolute, % Deviation from VaR
    ASSET_PREDICTED_COVAR_FUNC = Array(HIST1_VAR, HIST1_VAR / HIST2_VAR)
'----------------------------------------------------------------------
Case Else 'Historical ETL
'----------------------------------------------------------------------

    TEMP1_SUM = 0: CUMUL1_SUM = 0
    TEMP2_SUM = 0: CUMUL2_SUM = 0
    For i = 1 To NROWS
        If DATA1_VECTOR(i, 1) <= HIST1_VAR Then
            TEMP1_SUM = TEMP1_SUM + 1
            CUMUL1_SUM = CUMUL1_SUM + DATA1_VECTOR(i, 1)
        End If
        If DATA2_VECTOR(i, 1) <= HIST2_VAR Then
            TEMP2_SUM = TEMP2_SUM + 1
            CUMUL2_SUM = CUMUL2_SUM + DATA2_VECTOR(i, 1)
        End If
    Next i
    ASSET_PREDICTED_COVAR_FUNC = Array(CUMUL1_SUM / TEMP1_SUM, CUMUL2_SUM / TEMP2_SUM)
'----------------------------------------------------------------------
End Select
'----------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_PREDICTED_COVAR_FUNC = Err.number
End Function


'Maximum Likelihood estimation of a simple Gaussian respectively normal
'mixture distribution

'Gaussian Mixture Distribution (DATA_RNG --> DATE / DATA

Function ASSET_GAUSSIAN_MIXTURE_MLE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal INITIAL_PROBABILITY As Double = 0.6, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim DATA_VECTOR As Variant
Dim DATE_VECTOR As Variant

Dim DATA_MATRIX As Variant
Dim PARAM_VECTOR As Variant
Dim SORTED_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then
    DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
End If

If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
NROWS = UBound(DATA_MATRIX, 1)

ReDim DATA_VECTOR(1 To NROWS, 1 To 1)
ReDim DATE_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    DATE_VECTOR(i, 1) = DATA_MATRIX(i, 1)
    DATA_VECTOR(i, 1) = DATA_MATRIX(i, 2)
Next i

MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
SIGMA_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)
SORTED_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
'----------------------------------------------------------------------------------
If IsArray(PARAM_RNG) = True Then
'----------------------------------------------------------------------------------
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then
        PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    End If
'----------------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------------
    ReDim PARAM_VECTOR(1 To 5, 1 To 1)
    PARAM_VECTOR(1, 1) = INITIAL_PROBABILITY
    PARAM_VECTOR(2, 1) = Abs(MEAN_VAL)
    PARAM_VECTOR(3, 1) = SIGMA_VAL
    PARAM_VECTOR(4, 1) = -PARAM_VECTOR(2, 1)
    PARAM_VECTOR(5, 1) = PARAM_VECTOR(3, 1)
    PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION2_FUNC("GAUSSIAN_MIXTURE_OBJ_FUNC", DATA_VECTOR, PARAM_VECTOR)
'----------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------------
Case 0
'----------------------------------------------------------------------------------
    ReDim DATA_MATRIX(0 To NROWS, 1 To 12)
    DATA_MATRIX(0, 1) = "DATE"
    DATA_MATRIX(0, 2) = "DATA"
    DATA_MATRIX(0, 3) = "NORMAL (PROB DENSITY)"
    DATA_MATRIX(0, 4) = "NORMAL MIXTURE (PROB DENSITY)"
    DATA_MATRIX(0, 5) = "NORMAL MIX 1 (PROB DENSITY)"
    DATA_MATRIX(0, 6) = "NORMAL MIX 2 (PROB DENSITY)"
    DATA_MATRIX(0, 7) = "lnN (LN PROB DENSITY)"
    DATA_MATRIX(0, 8) = "lnNM (LN PROB DENSITY)"
    DATA_MATRIX(0, 9) = "SORTED DATA"
    DATA_MATRIX(0, 10) = "EMPIRICAL (CUMUL)"
    DATA_MATRIX(0, 11) = "NORMAL (CUMUL)"
    DATA_MATRIX(0, 12) = "NORMAL MIXTURE (CUMUL)"
    For i = 1 To NROWS
        DATA_MATRIX(i, 1) = DATE_VECTOR(i, 1)
        DATA_MATRIX(i, 2) = DATA_VECTOR(i, 1)
        DATA_MATRIX(i, 3) = NORMDIST_FUNC(DATA_MATRIX(i, 2), MEAN_VAL, SIGMA_VAL, 0)
        DATA_MATRIX(i, 4) = PARAM_VECTOR(1, 1) * NORMDIST_FUNC(DATA_MATRIX(i, 2), PARAM_VECTOR(2, 1), PARAM_VECTOR(3, 1), 0) + (1 - PARAM_VECTOR(1, 1)) * NORMDIST_FUNC(DATA_MATRIX(i, 2), PARAM_VECTOR(4, 1), PARAM_VECTOR(5, 1), 0)
        DATA_MATRIX(i, 5) = NORMDIST_FUNC(DATA_MATRIX(i, 2), PARAM_VECTOR(2, 1), PARAM_VECTOR(3, 1), 0)
        DATA_MATRIX(i, 6) = NORMDIST_FUNC(DATA_MATRIX(i, 2), PARAM_VECTOR(4, 1), PARAM_VECTOR(5, 1), 0)
        DATA_MATRIX(i, 7) = Log(DATA_MATRIX(i, 3))
        DATA_MATRIX(i, 8) = Log(DATA_MATRIX(i, 4))
        DATA_MATRIX(i, 9) = SORTED_VECTOR(i, 1)
        DATA_MATRIX(i, 10) = i / NROWS
        DATA_MATRIX(i, 11) = NORMSDIST_FUNC(DATA_MATRIX(i, 9), MEAN_VAL, SIGMA_VAL, 0)
        DATA_MATRIX(i, 12) = PARAM_VECTOR(1, 1) * NORMSDIST_FUNC(DATA_MATRIX(i, 9), PARAM_VECTOR(2, 1), PARAM_VECTOR(3, 1), 0) + (1 - PARAM_VECTOR(1, 1)) * NORMSDIST_FUNC(DATA_MATRIX(i, 9), PARAM_VECTOR(4, 1), PARAM_VECTOR(5, 1), 0)
    Next i
    ASSET_GAUSSIAN_MIXTURE_MLE_FUNC = DATA_MATRIX
'----------------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------------
    ASSET_GAUSSIAN_MIXTURE_MLE_FUNC = PARAM_VECTOR
'----------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_GAUSSIAN_MIXTURE_MLE_FUNC = Err.number
End Function

Private Function GAUSSIAN_MIXTURE_OBJ_FUNC(ByRef DATA_RNG As Variant, _
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

If (PARAM_VECTOR(1, 1) < 0) Or (PARAM_VECTOR(1, 1) > 1) Then
    GAUSSIAN_MIXTURE_OBJ_FUNC = 2 ^ 52
Else
    TEMP_SUM = 0
    For i = 1 To UBound(DATA_VECTOR, 1)
        TEMP_SUM = TEMP_SUM + Log(PARAM_VECTOR(1, 1) * NORMDIST_FUNC(DATA_VECTOR(i, 1), PARAM_VECTOR(2, 1), PARAM_VECTOR(3, 1), 0) + (1 - PARAM_VECTOR(1, 1)) * NORMDIST_FUNC(DATA_VECTOR(i, 1), PARAM_VECTOR(4, 1), PARAM_VECTOR(5, 1), 0))
    Next i
    GAUSSIAN_MIXTURE_OBJ_FUNC = -TEMP_SUM
End If

Exit Function
ERROR_LABEL:
GAUSSIAN_MIXTURE_OBJ_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_GPD_VAR_FUNC

'DESCRIPTION   : Tail Risk Modeling: Fitting the Generalized Pareto
'Distribution to historical returns and calculating VaR,
'Expected Shortfall (expected loss beyond VaR).

'LIBRARY       : FINAN_ASSET
'GROUP         : VAR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_GPD_VAR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal THRESHOLD As Double = -1, _
Optional ByVal SCALE_VAL As Double = 0.654513425377483, _
Optional ByVal SHAPE_VAL As Double = 0.115568986717223, _
Optional ByVal CONFIDENCE_VAL As Double = 0.01, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 2)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim PI_VAL As Double
Dim TEMP_SUM As Double
Dim TEMP_VAL As Double

Dim TEMP_ARR As Variant
Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
'------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------
    NROWS = UBound(DATA_VECTOR, 1)
    PI_VAL = 3.14159265358979
    
    ReDim TEMP_VECTOR(1 To 12, 1 To 2)
    
    TEMP_VECTOR(1, 1) = "NUMBER OF OBSERVATIONS"
    TEMP_VECTOR(2, 1) = "PER PERIOD VOLATILITY"
    TEMP_VECTOR(3, 1) = "PER PERIOD AVG RETURN"
    TEMP_VECTOR(4, 1) = "CONFIDENCE LEVEL"
    TEMP_VECTOR(5, 1) = "HISTORICAL VAR"
    TEMP_VECTOR(6, 1) = "HISTORICAL EXECTED SHORTFALL"
    TEMP_VECTOR(7, 1) = "PARAMETRIC NORMAL VAR"
    TEMP_VECTOR(8, 1) = "NORMAL EXPECTED SHORTFALL"
    TEMP_VECTOR(9, 1) = "GPD VAR"
    TEMP_VECTOR(10, 1) = "GPD EXPECTED SHORTFALL"
    TEMP_VECTOR(11, 1) = "THRESHOLD RETURN"
    TEMP_VECTOR(12, 1) = "NUMBER OF EXCEEDANCES"
    
    TEMP_VECTOR(1, 2) = NROWS
    TEMP_VECTOR(2, 2) = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)
    TEMP_VECTOR(3, 2) = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
    TEMP_VECTOR(4, 2) = 1 - CONFIDENCE_VAL
    
    j = CONFIDENCE_VAL * NROWS
    TEMP_ARR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
    TEMP_VECTOR(5, 2) = TEMP_ARR(j, 1)
    TEMP_VECTOR(6, 2) = ASSET_GPD_MEAN_FUNC(DATA_VECTOR, THRESHOLD)
    TEMP_VECTOR(7, 2) = NORMSINV_FUNC(1 - TEMP_VECTOR(4, 2), TEMP_VECTOR(3, 2), TEMP_VECTOR(2, 2), 0)
    TEMP_VECTOR(8, 2) = -TEMP_VECTOR(2, 2) * Exp(-0.5 * NORMSINV_FUNC(1 - TEMP_VECTOR(4, 2), 0, 1, 0) ^ 2) / ((1 - TEMP_VECTOR(4, 2)) * Sqr(2 * PI_VAL))
    
    TEMP_VAL = ASSET_GPD_COUNT_FUNC(DATA_VECTOR, THRESHOLD)
    TEMP_VECTOR(9, 2) = THRESHOLD - SCALE_VAL / SHAPE_VAL * (((TEMP_VECTOR(1, 2) / TEMP_VAL) * (1 - TEMP_VECTOR(4, 2))) ^ (-SHAPE_VAL) - 1)
    
    TEMP_VECTOR(10, 2) = -(-TEMP_VECTOR(9, 2) + SCALE_VAL + SHAPE_VAL * THRESHOLD) / (1 - SHAPE_VAL)
    TEMP_VECTOR(11, 2) = THRESHOLD
    TEMP_VECTOR(12, 2) = TEMP_VAL
    
    ASSET_GPD_VAR_FUNC = TEMP_VECTOR
'------------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------------
    TEMP_ARR = GPD_DIST_FUNC(DATA_VECTOR, THRESHOLD)
    NROWS = UBound(TEMP_ARR, 1)
    
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 6)
    
    TEMP_MATRIX(0, 1) = "INDEX"
    TEMP_MATRIX(0, 2) = "R<=T"
    TEMP_MATRIX(0, 3) = "EMP CDF"
    TEMP_MATRIX(0, 4) = "GPD CDF"
    TEMP_MATRIX(0, 5) = "GPD INV CDF"
    TEMP_MATRIX(0, 6) = "SQR DIFF"
    
    TEMP_SUM = 0
    For i = 1 To NROWS
        
        TEMP_MATRIX(i, 1) = i
        TEMP_MATRIX(i, 2) = TEMP_ARR(i, 1)
        TEMP_MATRIX(i, 3) = TEMP_ARR(i, 2)
        
        If IsNumeric(TEMP_ARR(i, 1)) = True Then
            TEMP_VAL = GPD_CDF_FUNC(-1 * TEMP_ARR(i, 1), SHAPE_VAL, SCALE_VAL, THRESHOLD * -1)
            TEMP_MATRIX(i, 4) = 1 - TEMP_VAL
        Else
            TEMP_MATRIX(i, 4) = CVErr(xlErrNA)
        End If

        If IsNumeric(TEMP_ARR(i, 2)) = True Then
            TEMP_VAL = GPD_INV_CDF_FUNC(1 - TEMP_ARR(i, 2), SHAPE_VAL, SCALE_VAL, THRESHOLD * -1)
            TEMP_MATRIX(i, 5) = -1 * TEMP_VAL
            TEMP_MATRIX(i, 6) = (TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 2)) ^ 2
            TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 6)
        Else
            TEMP_MATRIX(i, 5) = CVErr(xlErrNA)
            TEMP_MATRIX(i, 6) = CVErr(xlErrNA)
        End If

    Next i
    
    If OUTPUT = 1 Then
        ASSET_GPD_VAR_FUNC = TEMP_MATRIX
    Else 'RMSE
    '...goal function of the distribution fitting: minimize square root
    'of average squared deviation of the empirical form the theoretical
    'distribution.
        ASSET_GPD_VAR_FUNC = Sqr(TEMP_SUM / ASSET_GPD_COUNT_FUNC(DATA_VECTOR, THRESHOLD)) '...run Min SOLVER to re-fit the data
        'Changing Cells: SCALE_VAL & SHAPE VAL
        'Constraints: SCALE_VAL >= 0
    End If
End Select

Exit Function
ERROR_LABEL:
ASSET_GPD_VAR_FUNC = Err.number
End Function

Private Function ASSET_GPD_MEAN_FUNC(ByRef DATA_RNG As Variant, _
ByVal THRESHOLD As Double)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim TEMP_VAL As Double
Dim TEMP_SUM As Double

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)

NROWS = UBound(DATA_VECTOR, 1)
j = 0
TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_VAL = DATA_VECTOR(i, 1)
    If TEMP_VAL <= THRESHOLD Then
        j = j + 1
        TEMP_SUM = TEMP_SUM + TEMP_VAL
    End If
Next i

ASSET_GPD_MEAN_FUNC = TEMP_SUM / j
Exit Function
ERROR_LABEL:
ASSET_GPD_MEAN_FUNC = Err.number
End Function

Private Function ASSET_GPD_COUNT_FUNC(ByRef DATA_RNG As Variant, _
ByVal THRESHOLD As Double)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim TEMP_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)

NROWS = UBound(DATA_VECTOR, 1)
j = 0
For i = 1 To NROWS
    TEMP_VAL = DATA_VECTOR(i, 1)
    If TEMP_VAL <= THRESHOLD Then: j = j + 1
Next i

ASSET_GPD_COUNT_FUNC = j

Exit Function
ERROR_LABEL:
ASSET_GPD_COUNT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TDIST_VAR_FUNC

'DESCRIPTION   : Tail Risk Modeling: Fitting the T-Distribution to historical
'returns and calculating VaR, Expected Shortfall (expected loss beyond VaR).

'LIBRARY       : FINAN_ASSET
'GROUP         : VAR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_TDIST_VAR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal nDEGREES As Variant = 30, _
Optional ByVal CONFIDENCE_VAL As Double = 0.025, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim TEMP_SUM As Double
Dim MEAN_VAL As Double
Dim VOLAT_VAL As Double

Dim DF_VAL As Double
Dim FLAG_VAL As Boolean 'IS FIT OF T DISTR BETTER THAN NORMAL DISTR?

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)
DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)

TEMP_VECTOR = FIT_TDIST_FUNC(DATA_VECTOR, nDEGREES)
DF_VAL = TEMP_VECTOR(LBound(TEMP_VECTOR) + 1)
FLAG_VAL = TEMP_VECTOR(LBound(TEMP_VECTOR) + 2)
Erase TEMP_VECTOR

MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
VOLAT_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)
        
ReDim TEMP_MATRIX(0 To NROWS, 1 To 8)

TEMP_MATRIX(0, 1) = "INDEX"
TEMP_MATRIX(0, 2) = "SORTED RETURN"
TEMP_MATRIX(0, 3) = "STAND RETURN"
TEMP_MATRIX(0, 4) = "EMPIR CDF"
TEMP_MATRIX(0, 5) = "T-CDF"
TEMP_MATRIX(0, 6) = "T-CDF INV"
TEMP_MATRIX(0, 7) = "T-CDF FIT"
TEMP_MATRIX(0, 8) = "NORM FIT"

TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = (TEMP_MATRIX(i, 2) - MEAN_VAL) / VOLAT_VAL
    TEMP_MATRIX(i, 4) = i / (1 + NROWS)
    
    If TEMP_MATRIX(i, 3) > 0 Then
        TEMP_MATRIX(i, 5) = TDIST_FUNC(TEMP_MATRIX(i, 3), DF_VAL, True)
    Else
        TEMP_MATRIX(i, 5) = 1 - TDIST_FUNC(-1 * TEMP_MATRIX(i, 3), DF_VAL, True)
    End If
    
    If TEMP_MATRIX(i, 4) < 0.5 Then
        TEMP_MATRIX(i, 6) = INVERSE_TDIST_FUNC(TEMP_MATRIX(i, 4), DF_VAL)
    Else
        TEMP_MATRIX(i, 6) = INVERSE_TDIST_FUNC((1 - TEMP_MATRIX(i, 4)), DF_VAL) * -1
    End If
    
    TEMP_MATRIX(i, 7) = MEAN_VAL + TEMP_MATRIX(i, 6) * VOLAT_VAL
    TEMP_MATRIX(i, 8) = MEAN_VAL + NORMSINV_FUNC(TEMP_MATRIX(i, 4), 0, 1, 0) * VOLAT_VAL
Next i

'------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------
    NROWS = UBound(DATA_VECTOR, 1)
    
    ReDim TEMP_VECTOR(1 To 10, 1 To 2)
    
    TEMP_VECTOR(1, 1) = "NUMBER OF OBSERVATIONS"
    TEMP_VECTOR(1, 2) = NROWS
    
    TEMP_VECTOR(2, 1) = "PER PERIOD VOLATILITY"
    TEMP_VECTOR(2, 2) = VOLAT_VAL
    
    TEMP_VECTOR(3, 1) = "PER PERIOD AVG RETURN"
    TEMP_VECTOR(3, 2) = MEAN_VAL
        
    TEMP_VECTOR(4, 1) = "CONFIDENCE LEVEL"
    TEMP_VECTOR(4, 2) = 1 - CONFIDENCE_VAL
    
    TEMP_VECTOR(5, 1) = "HISTORICAL VAR"
    j = CONFIDENCE_VAL * NROWS
    TEMP_VECTOR(5, 2) = DATA_VECTOR(j, 1)
    
    TEMP_VECTOR(6, 1) = "STUDENT T VAR"
    TEMP_VECTOR(6, 2) = TEMP_MATRIX(j, 7)
    
    TEMP_VECTOR(7, 1) = "NORMAL VAR"
    TEMP_VECTOR(7, 2) = TEMP_MATRIX(j, 8)
    
    TEMP_VECTOR(8, 1) = "UPPER BOUND DEGREES OF FREEDOM"
    TEMP_VECTOR(8, 2) = nDEGREES
    
    TEMP_VECTOR(9, 1) = "BEST-FIT DEGREES OF FREEDOM"
    TEMP_VECTOR(9, 2) = DF_VAL
    
    TEMP_VECTOR(10, 1) = "IS FIT OF T DISTR BETTER THAN NORMAL DISTR?"
    TEMP_VECTOR(10, 2) = FLAG_VAL
    
    
    ASSET_TDIST_VAR_FUNC = TEMP_VECTOR
'------------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------------
    ASSET_TDIST_VAR_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
ASSET_TDIST_VAR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_VARS_FUNC
'DESCRIPTION   : Calculates parametric Value-At-Risk based
'LIBRARY       : FINAN_ASSET
'GROUP         : VAR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_VARS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 0.05, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal PERIODS As Integer = 12, _
Optional ByVal COUNT_BASIS As Integer = 12)

'COUNT_BASIS = data frequency expressed as a fraction of a year
'PERIODS = horizon expressed as a fraction of a year

Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim TEMP_DEV As Double
Dim TEMP_VAR As Double

Dim MEAN_VAL As Double
Dim VOLAT_VAL As Double
Dim SKEW_VAL As Double
Dim KURT_VAL As Double

Dim FACTOR_VAL As Variant
Dim NORMSINV_VAL As Double
Dim NORMDIST_VAL As Double

Dim HVAR_VAL As Double
Dim NVAR_VAL As Double
Dim CVAR_VAL As Double
Dim MVAR_VAL As Double

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)

GoSub MOMENTS_LINE
GoSub HVAR_LINE
GoSub NVAR_LINE
GoSub MVAR_LINE
GoSub CVAR_LINE

ASSET_VARS_FUNC = Array(MEAN_VAL, VOLAT_VAL, SKEW_VAL, KURT_VAL, NORMSINV_VAL, NORMDIST_VAL, HVAR_VAL, NVAR_VAL, MVAR_VAL, CVAR_VAL)

'----------------------------------------------------------------------------------------------------------------
Exit Function
'----------------------------------------------------------------------------------------------------------------
MOMENTS_LINE:
'----------------------------------------------------------------------------------------------------------------
    TEMP_SUM = 0
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1)
    Next i
    MEAN_VAL = TEMP_SUM / NROWS
    TEMP_DEV = 0: TEMP_SUM = 0: TEMP_VAR = 0
    For i = 1 To NROWS
        TEMP_DEV = (DATA_VECTOR(i, 1) - MEAN_VAL)
        TEMP_SUM = TEMP_SUM + TEMP_DEV
        TEMP_VAR = TEMP_DEV * TEMP_DEV + TEMP_VAR
    Next i
    VOLAT_VAL = (TEMP_VAR - TEMP_SUM * TEMP_SUM / NROWS) / (NROWS - 1)
        'Variance: Corrected two-pass formula.
        'VOLAT_VAL = Sqr(NROWS / (NROWS - 1)) * VOLAT_VAL 'Population
    VOLAT_VAL = Sqr(VOLAT_VAL) 'Sample Standard Deviation
    SKEW_VAL = 0: KURT_VAL = 0     ' Calculate 3rd and 4th moments
    For i = 1 To NROWS
        SKEW_VAL = SKEW_VAL + ((DATA_VECTOR(i, 1) - MEAN_VAL) / VOLAT_VAL) ^ 3
        KURT_VAL = KURT_VAL + ((DATA_VECTOR(i, 1) - MEAN_VAL) / VOLAT_VAL) ^ 4
    Next i
    SKEW_VAL = SKEW_VAL / NROWS
'    SKEW_VAL = SKEW_VAL * (NROWS / ((NROWS - 1) * (NROWS - 2))) 'Excel Definition
    KURT_VAL = (KURT_VAL / NROWS) - 3
'    KURT_VAL = (KURT_VAL * (NROWS * (NROWS + 1) / ((NROWS - 1) * (NROWS - 2) * (NROWS - 3)))) - ((3 * (NROWS - 1) ^ 2 / ((NROWS - 2) * (NROWS - 3)))) 'Excel Definition
    NORMSINV_VAL = NORMSINV_FUNC(CONFIDENCE_VAL, 0, 1, 0)
    NORMDIST_VAL = NORMDIST_FUNC(NORMSINV_VAL, 0, 1, 1)
'----------------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------------
HVAR_LINE: 'Calculates Historical Value-At-Risk
'----------------------------------------------------------------------------------------------------------------
    HVAR_VAL = HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, CONFIDENCE_VAL, 1) * ((1 / PERIODS) / (1 / COUNT_BASIS))
'----------------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------------
NVAR_LINE: 'Calculates Value-At-Risk based on the assumption of non-normal distribution
'----------------------------------------------------------------------------------------------------------------
    NVAR_VAL = MEAN_VAL * ((1 / PERIODS) / (1 / COUNT_BASIS)) + NORMSINV_VAL * VOLAT_VAL * ((1 / PERIODS) / (1 / COUNT_BASIS)) ^ (0.5)
'----------------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------------
MVAR_LINE: 'Calculates Value-At-Risk based on the assumption of non-normal distribution, making use of the Cornish-Fisher Expansion
'----------------------------------------------------------------------------------------------------------------
'Modified VaR (calcs with first three moments)
    FACTOR_VAL = NORMSINV_VAL + (1 / 6) * (NORMSINV_VAL ^ 2 - 1) * SKEW_VAL + (1 / 24) * (NORMSINV_VAL ^ 3 - 3 * NORMSINV_VAL) * KURT_VAL - (1 / 36) * (2 * NORMSINV_VAL ^ 3 - 5 * NORMSINV_VAL) * SKEW_VAL ^ 2
    MVAR_VAL = MEAN_VAL * ((1 / PERIODS) / (1 / COUNT_BASIS)) + FACTOR_VAL * VOLAT_VAL * ((1 / PERIODS) / (1 / COUNT_BASIS)) ^ (0.5)
'----------------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------------
CVAR_LINE:
'Calculates Conditional Value-At-Risk based on the assumption of a normal distribution. Whereas VaR
'measures the maximum loss for a given confidence interval, "Conditional VaR" (CVaR) corresponds to the expected
'loss conditional on the loss being greater than or equal to the VaR. On the other hand, it's not really clear
'what 'expected' means once you get into the 1% tail. Conditional VaR answers an entirely hypothetical
'question; 'what is your expected loss in a 'normal market', once the market moves beyond normal limits? Because
'this question can never be answered empirically, conditional VaR cannot be validated or backtested.

'Conditional VaR is a useful specialized tool, primarily in optimization and in some highly stable situations. But
'VaR has far more general utility. CVar is also known as Expected Tail Loss and Expected Shortfall.
'----------------------------------------------------------------------------------------------------------------
    CVAR_VAL = MEAN_VAL * ((1 / PERIODS) / (1 / COUNT_BASIS)) - (NORMDIST_VAL / CONFIDENCE_VAL) * VOLAT_VAL * ((1 / PERIODS) / (1 / COUNT_BASIS)) ^ (0.5)
'----------------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
'----------------------------------------------------------------------------------------------------------------
ASSET_VARS_FUNC = Err.number
'----------------------------------------------------------------------------------------------------------------
'Value-At-Risk (VaR) answers the question, "How much can the value of a
'portfolio decline with a given probability in a given time period?".

'The most common assumption is that returns follow a normal distribution.
'One of the properties of the normal distribution is that 95 percent of
'all observations occur within 1.96 standard deviations from the mean. This
'means that the probability that an observation will fall 1.96 standard
'deviations below the mean is only 2.5 percent. For the purposes of
'calculating VAR we are interested only in losses, not gains, so this is
'the relevant probability.

'Example: XYZ Fund has an (arithmetic) average monthly return of 2.03 percent
'and a standard deviation of 3.27 percent. Thus, its monthly VAR at the 2.5
'percent probability level is 2.03%-1.96*3.27=-4.38%, or $43.80 for a $1,000
'investment, meaning that the probability of losing more than this is 2.5 percent.

'VAR is often said to have an advantage over other risk measures in that it
'is more forward-looking.  While it can be described as forward-looking, VAR
'relies on information derived from historical price time series many times.
'However, the strength of VAR models is that they allow us to construct a
'measure of risk for the portfolio not from its own past volatility but from the
'volatilities of risk factors affecting the portfolio as it is constituted today.

'Clearly, if the present composition of the fund's portfolio is significantly different
'than it was during the past year, then historical measures would not predict its
'future performance very accurately. However, as long as we know the fund's current
'composition and can assume that it will stay the same during the period for which
'we want to know the VAR, we can use a model based on the historical data about the
'risk factors to make statistical inferences about the probability distribution of
'the fund's future returns. In fact, for certain portfolios it is necessary to have
'a model based on risk factors even if one does not trade the portfolio at all. This
'is particularly true for portfolios consisting of bonds and/or options and futures,
'because such portfolios "age," that is, their characteristics change from the passage
'of time alone. In particular, as bonds approach maturity, their value approaches face
'value and their volatility diminishes and disappears altogether at maturity, when the
'bond can be redeemed at face value. Options, on the other hand, tend to lose value as
'they approach expiration, all other things being equal. This is one of the reasons why
'VAR analysis is used more frequently in derivatives and fixed-income investment and is
'less widespread for equities.

'Risk managers at mutual fund companies may also be interested in the value at risk as
'it applies to underperforming the fund's chosen benchmark. This measure, known as
'"relative" or "tracking" VAR, can be thought of as the VAR of a portfolio consisting
'of long positions in all the stocks the fund currently owns and a short position in
'the fund's benchmark.

'While VAR provides a view of risk based on low-probability losses, for symmetrical
'bell-shaped distributions such as those typically followed by stock returns, VAR is
'highly correlated with volatility as measured by the standard deviation. In fact,
'for normally distributed returns, value at risk is directly proportional to standard
'deviation.
 
'Comparison of VaR and Classical Portfolio Theory

'"   Risk in portfolio theory = standard deviation of returns,
'    Risk in VaR = maximum likely loss

'"   Variance-covariance approach to VaR has the same theoretical
'    basis as portfolio theory. Not so the historical simulation
'    approach and the Monte Carlo simulation approach to VaR.

'"   Portfolio theory is limited to market risk, while VaR can be
'    applied to a much broader range of problems (credit, liquidity,
'    operational risks etc.)

'"   VaR can better accommodate statistical issues like non-normal returns

'"   VaR can be applied for firm-wide risk management and provides better
'    rules than portfolio theory to guide investment, hedging and portfolio
'    management decisions.
  
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
End Function


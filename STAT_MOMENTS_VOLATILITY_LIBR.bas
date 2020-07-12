Attribute VB_Name = "STAT_MOMENTS_VOLATILITY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : UPDATE_VOLATILITY_FUNC
'DESCRIPTION   : UPDATE STANDARD DEVIATION BASED ON NEW SPOT PRICE
'LIBRARY       : STATISTICS
'GROUP         : VOLATILITY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function UPDATE_VOLATILITY_FUNC(ByVal OLD_SIGMA As Double, _
ByVal OLD_SPOT As Double, _
ByVal NEW_SPOT As Double, _
Optional ByVal LAMBDA As Double = 0.99, _
Optional ByVal COUNT_BASIS As Double = 250)
On Error GoTo ERROR_LABEL
UPDATE_VOLATILITY_FUNC = Sqr(LAMBDA * OLD_SIGMA ^ 2 + (1 - LAMBDA) * (Log(NEW_SPOT / OLD_SPOT)) ^ 2 * COUNT_BASIS)
Exit Function
ERROR_LABEL:
UPDATE_VOLATILITY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : FORWARD_VOLATILITY_FUNC
'DESCRIPTION   : FORWARD FORWARD VOLATILITY
'LIBRARY       : STATISTICS
'GROUP         : VOLATILITY
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function FORWARD_VOLATILITY_FUNC(ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByVal END_SIGMA As Double, _
ByVal REF_DATE As Date, _
ByVal REF_SIGMA As Double)
On Error GoTo ERROR_LABEL
If END_DATE > REF_DATE Then
    FORWARD_VOLATILITY_FUNC = Sqr(MAXIMUM_FUNC((END_SIGMA ^ 2 * (END_DATE - START_DATE) - REF_SIGMA ^ 2 * (REF_DATE - START_DATE)) / (END_DATE - REF_DATE), 0))
Else
    FORWARD_VOLATILITY_FUNC = END_SIGMA
End If
Exit Function
ERROR_LABEL:
FORWARD_VOLATILITY_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_STDEV_FUNC
'DESCRIPTION   : Estimates standard deviation based on a sample. The standard
'deviation is a measure of how widely values are dispersed from the average
'value (the mean).
'LIBRARY       : STATISTICS
'GROUP         : VOLATILITY
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_STDEV_FUNC(ByRef DATA_RNG As Variant)

Dim j As Long
Dim i As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double
Dim TEMP4_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To 1, 1 To NCOLUMNS)


For j = 1 To NCOLUMNS
    TEMP2_VAL = 0
    TEMP3_VAL = 0
    TEMP4_VAL = 0
    For i = 1 To NROWS
        TEMP2_VAL = TEMP2_VAL + DATA_MATRIX(i, j)
    Next i
    TEMP1_VAL = TEMP2_VAL / NROWS
    For i = 1 To NROWS
        TEMP3_VAL = TEMP3_VAL + (DATA_MATRIX(i, j) - TEMP1_VAL) ^ 2
        TEMP4_VAL = TEMP4_VAL + (DATA_MATRIX(i, j) - TEMP1_VAL)
    Next i
    TEMP_MATRIX(1, j) = Sqr((TEMP3_VAL - TEMP4_VAL * TEMP4_VAL / NROWS) / (NROWS - 1)) 'Rounding error corrected var formula.
Next j

MATRIX_STDEV_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_STDEV_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_STDEVP_FUNC
'DESCRIPTION   : Calculates standard deviation based on the entire population
'given as arguments
'LIBRARY       : STATISTICS
'GROUP         : VOLATILITY
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_STDEVP_FUNC(ByRef DATA_RNG As Variant)

Dim j As Long
Dim i As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double
Dim TEMP4_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To 1, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    TEMP2_VAL = 0: TEMP3_VAL = 0 'TEMP4_VAL = 0
    For i = 1 To NROWS: TEMP2_VAL = TEMP2_VAL + DATA_MATRIX(i, j): Next i
    TEMP1_VAL = TEMP2_VAL / NROWS
    For i = 1 To NROWS
        TEMP3_VAL = TEMP3_VAL + (DATA_MATRIX(i, j) - TEMP1_VAL) ^ 2
    'TEMP4_VAL = TEMP4_VAL + (DATA_MATRIX(i, j) - TEMP1_VAL)
    Next i
    TEMP_MATRIX(1, j) = Sqr(TEMP3_VAL / NROWS)
    'TEMP_MATRIX(1, j) = Sqr(NROWS / (NROWS - 1)) * Sqr((TEMP3_VAL - TEMP4_VAL * TEMP4_VAL / NROWS) / (NROWS - 1))
    'Rounding error corrected var formula.
Next j

MATRIX_STDEVP_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_STDEVP_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_AVERAGE_VOLATILITY_FUNC
'DESCRIPTION   : Computes average volatility from a matrix of returns
'LIBRARY       : STATISTICS
'GROUP         : VOLATILITY
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_AVERAGE_VOLATILITY_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

DATA_MATRIX = MATRIX_COVARIANCE_FRAME3_FUNC(DATA_MATRIX, 0, 0)
NCOLUMNS = UBound(DATA_MATRIX, 2)
TEMP_SUM = 0
For i = 1 To NCOLUMNS
    TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, i) ^ 0.5
Next i
MATRIX_AVERAGE_VOLATILITY_FUNC = TEMP_SUM / NCOLUMNS

Exit Function
ERROR_LABEL:
MATRIX_AVERAGE_VOLATILITY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_MOVING_AVERAGE_VOLATILITY_FUNC
'DESCRIPTION   : Volatility Estimate with Moving Average
'LIBRARY       : STATISTICS
'GROUP         : VOLATILITY
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_MOVING_AVERAGE_VOLATILITY_FUNC(ByRef DATA_RNG As Variant, _
ByVal MA_FACTOR As Variant, _
ByVal LAMBDA As Double, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

'IF DATA is in Descending Order Then:
'=REVERSE_FUNC(VECTOR_MOVING_AVERAGE_VOLATILITY_FUNC(MATRIX_REVERSE_FUNC(""),MA_FACTOR,lamda,1))

Dim i As Long
Dim NROWS As Long

Dim DATA_VECTOR As Variant

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)

NROWS = UBound(DATA_VECTOR)

ReDim TEMP1_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    TEMP1_VECTOR(i, 1) = DATA_VECTOR(i, 1) ^ 2 'SQUARED RETURNS
Next i

TEMP2_VECTOR = VECTOR_MOVING_AVERAGE_FUNC(TEMP1_VECTOR, MA_FACTOR) 'MOVING AVERAGE SQUARED RETURNS

ReDim TEMP_MATRIX(0 To NROWS, 1 To 4)

TEMP_MATRIX(0, 1) = "RETURN"
TEMP_MATRIX(0, 2) = "RETURN ^ 2"
TEMP_MATRIX(0, 3) = "MA RETURN ^ 0.5"
TEMP_MATRIX(0, 4) = "SIGMA ESTIMATE" 'EWMA

TEMP_MATRIX(1, 1) = DATA_VECTOR(1, 1)
TEMP_MATRIX(1, 2) = TEMP1_VECTOR(1, 1)
TEMP_MATRIX(1, 3) = Sqr(TEMP2_VECTOR(1, 1))
TEMP_MATRIX(1, 4) = 0

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = TEMP1_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = Sqr(TEMP2_VECTOR(i, 1))
    TEMP_MATRIX(i, 4) = Sqr(LAMBDA * TEMP_MATRIX(i - 1, 4) ^ 2 + (1 - LAMBDA) * TEMP_MATRIX(i - 1, 2))
Next i

VECTOR_MOVING_AVERAGE_VOLATILITY_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
VECTOR_MOVING_AVERAGE_VOLATILITY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_VOLATILITY_FORECAST_FUNC

'DESCRIPTION   : Calculates volatilty forecasts based on the exponentially
'weighted moving average model (=RiskMetrics approach)

'LIBRARY       : STATISTICS
'GROUP         : VOLATILITY
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_VOLATILITY_FORECAST_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal LAMBDA As Double = 0.9, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim MEAN_VAL As Double

Dim DATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)

NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

If VERSION = 0 Then
    For j = 1 To NROWS
        TEMP_SUM = TEMP_SUM + DATA_VECTOR(j, 1)
        MEAN_VAL = TEMP_SUM / j
    Next j
Else
    TEMP_SUM = 0
    MEAN_VAL = 0
End If

TEMP_VECTOR(1, 1) = ((1 - LAMBDA) * DATA_VECTOR(1, 1) ^ 2) ^ 0.5

For i = 2 To NROWS
    TEMP_VECTOR(i, 1) = (LAMBDA * TEMP_VECTOR(i - 1, 1) ^ 2 + (1 - LAMBDA) * (DATA_VECTOR(i, 1) - MEAN_VAL) ^ 2) ^ 0.5
Next i

VECTOR_VOLATILITY_FORECAST_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_VOLATILITY_FORECAST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_EWMA_VOLATILITY_FUNC
'DESCRIPTION   : Calculates exponentially weighted moving average model
'LIBRARY       : STATISTICS
'GROUP         : VOLATILITY
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************
    
Function MATRIX_EWMA_VOLATILITY_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal LAMBDA As Double = 0, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then: DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
If DATA_TYPE <> 0 Then: DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    TEMP_MATRIX(1, j) = DATA_MATRIX(1, j)
    For i = 2 To NROWS
        TEMP_MATRIX(i, j) = LAMBDA * DATA_MATRIX(i, j) + (1 - LAMBDA) * _
                        TEMP_MATRIX(i - 1, j)
    Next i
Next j

MATRIX_EWMA_VOLATILITY_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_EWMA_VOLATILITY_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_HP_VOLATILITY_FILTER_FUNC

'DESCRIPTION   : Hodrick-Prescott Filter: A very useful volatility smoothing
'algorithm --> LAMBDA: 1 = minor smooting, 10 = medium smoothing,
'100 = massive smoothing

'LIBRARY       : STATISTICS
'GROUP         : VOLATILITY
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_HP_VOLATILITY_FILTER_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal LAMBDA As Double = 0, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim H1() As Double
Dim H2() As Double
Dim H3() As Double
Dim H4() As Double
Dim H5() As Double

Dim HH1() As Double
Dim HH2() As Double
Dim HH3() As Double
Dim HH4() As Double
Dim HH5() As Double

Dim HHH1() As Double
Dim HHH2() As Double
Dim HHH3() As Double

Dim TEMP1_MATRIX() As Double
Dim TEMP2_MATRIX() As Double
Dim TEMP3_MATRIX() As Double

Dim TEMP_MATRIX() As Double

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then: DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
If DATA_TYPE <> 0 Then: DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
For i = 1 To NCOLUMNS Step 1
    For j = 1 To NROWS Step 1
        TEMP_MATRIX(j, i) = DATA_MATRIX(j, i)
    Next j
Next i

If NROWS <= 3 Then
    MATRIX_HP_VOLATILITY_FILTER_FUNC = TEMP_MATRIX
Else
    'creates pentadiagonal Matrix'
    ReDim TEMP1_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    ReDim TEMP2_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    ReDim TEMP3_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    
    For j = 1 To NCOLUMNS Step 1
        TEMP1_MATRIX(1, j) = 1 + LAMBDA
        TEMP2_MATRIX(1, j) = -2 * LAMBDA
        TEMP3_MATRIX(1, j) = LAMBDA
        For i = 2 To NROWS - 1 Step 1
            TEMP1_MATRIX(i, j) = 6 * LAMBDA + 1
            TEMP2_MATRIX(i, j) = -4 * LAMBDA
            TEMP3_MATRIX(i, j) = LAMBDA
        Next i
        TEMP1_MATRIX(2, j) = 5 * LAMBDA + 1
        TEMP1_MATRIX(NROWS, j) = 1 + LAMBDA
        TEMP1_MATRIX(NROWS - 1, j) = 5 * LAMBDA + 1
        TEMP2_MATRIX(1, j) = -2 * LAMBDA
        TEMP2_MATRIX(NROWS - 1, j) = -2 * LAMBDA
        TEMP2_MATRIX(NROWS, j) = 0
        TEMP3_MATRIX(NROWS - 1, j) = 0
        TEMP3_MATRIX(NROWS, j) = 0
    Next j
    'Solving system of linear equations'
    
    ReDim H1(1 To NCOLUMNS)
    ReDim H2(1 To NCOLUMNS)
    ReDim H3(1 To NCOLUMNS)
    ReDim H4(1 To NCOLUMNS)
    ReDim H5(1 To NCOLUMNS)
    
    ReDim HH1(1 To NCOLUMNS)
    ReDim HH2(1 To NCOLUMNS)
    ReDim HH3(1 To NCOLUMNS)
    ReDim HH4(1 To NCOLUMNS)
    ReDim HH5(1 To NCOLUMNS)
    
    ReDim HHH1(1 To NCOLUMNS)
    ReDim HHH2(1 To NCOLUMNS)
    ReDim HHH3(1 To NCOLUMNS)
    
    'Forward'
    For j = 1 To NCOLUMNS Step 1
        For i = 1 To NROWS Step 1
            HHH1(j) = TEMP1_MATRIX(i, j) - H4(j) * H1(j) - HH5(j) * HH2(j)
            HHH2(j) = TEMP2_MATRIX(i, j)
            HH1(j) = H1(j)
            H1(j) = (HHH2(j) - H4(j) * H2(j)) / HHH1(j)
            TEMP2_MATRIX(i, j) = H1(j)
            HHH3(j) = TEMP3_MATRIX(i, j)
            HH2(j) = H2(j)
            H2(j) = HHH3(j) / HHH1(j)
            TEMP3_MATRIX(i, j) = H2(j)
            TEMP1_MATRIX(i, j) = (TEMP_MATRIX(i, j) - HH3(j) * _
                                HH5(j) - H3(j) * H4(j)) / HHH1(j)
            HH3(j) = H3(j)
            H3(j) = TEMP1_MATRIX(i, j)
            H4(j) = HHH2(j) - H5(j) * HH1(j)
            HH5(j) = H5(j)
            H5(j) = HHH3(j)
        Next i
    H2(j) = 0
    H1(j) = TEMP1_MATRIX(NROWS, j)
    TEMP_MATRIX(NROWS, j) = H1(j)
    'Backward'
        For i = NROWS To 1 Step -1
            TEMP_MATRIX(i, j) = TEMP1_MATRIX(i, j) - TEMP2_MATRIX(i, j) * _
                                H1(j) - TEMP3_MATRIX(i, j) * H2(j)
            H2(j) = H1(j)
            H1(j) = TEMP_MATRIX(i, j)
        Next i
    Next j
    MATRIX_HP_VOLATILITY_FILTER_FUNC = TEMP_MATRIX
End If

Exit Function
ERROR_LABEL:
MATRIX_HP_VOLATILITY_FILTER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : SHRINK_VOLATILITY_VECTOR_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : VOLATILITY
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function SHRINK_VOLATILITY_VECTOR_FUNC(ByRef VOLATILITY_RNG As Variant, _
ByVal SHRINKAGE_FACTOR As Double)

Dim i As Long
Dim NSIZE As Long

Dim MEAN_VAL As Double
Dim VOLATILITY_VECTOR As Variant

On Error GoTo ERROR_LABEL

VOLATILITY_VECTOR = VOLATILITY_RNG
If UBound(VOLATILITY_VECTOR, 1) = 1 Then: _
    VOLATILITY_VECTOR = MATRIX_TRANSPOSE_FUNC(VOLATILITY_VECTOR)

NSIZE = UBound(VOLATILITY_VECTOR, 1)
MEAN_VAL = MATRIX_MEAN_FUNC(VOLATILITY_VECTOR)(1, 1)
For i = 1 To NSIZE
    VOLATILITY_VECTOR(i, 1) = VOLATILITY_VECTOR(i, 1) + SHRINKAGE_FACTOR * _
                            (MEAN_VAL - VOLATILITY_VECTOR(i, 1))
Next i
SHRINK_VOLATILITY_VECTOR_FUNC = VOLATILITY_VECTOR

Exit Function
ERROR_LABEL:
SHRINK_VOLATILITY_VECTOR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_VOLATILITY_COVARIANCE_FUNC
'DESCRIPTION   : Compute standard deviation matrix from covariance matrix
'LIBRARY       : STATISTICS
'GROUP         : VOLATILITY
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_VOLATILITY_COVARIANCE_FUNC(ByRef COVARIANCE_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VECTOR As Variant
Dim COVARIANCE_MATRIX As Variant

On Error GoTo ERROR_LABEL

COVARIANCE_MATRIX = COVARIANCE_RNG
NROWS = UBound(COVARIANCE_MATRIX, 1)
NCOLUMNS = UBound(COVARIANCE_MATRIX, 2)

If NROWS <> NCOLUMNS Then: GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To NCOLUMNS, 1 To 1)

For i = 1 To NCOLUMNS
    If IsNumeric(COVARIANCE_MATRIX(i, i)) And Not IsEmpty(COVARIANCE_MATRIX(i, i)) Then
        TEMP_VECTOR(i, 1) = COVARIANCE_MATRIX(i, i) ^ 0.5
    Else
        TEMP_VECTOR(i, 1) = "N/A"
    End If
Next i

MATRIX_VOLATILITY_COVARIANCE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_VOLATILITY_COVARIANCE_FUNC = Err.number
End Function

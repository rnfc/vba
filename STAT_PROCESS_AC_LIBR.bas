Attribute VB_Name = "STAT_PROCESS_AC_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : AUTO_CORREL_PARTIAL_SAMPLE_COEF_FUNC
'DESCRIPTION   : Find sample partial autocorrelation coefficients
'LIBRARY       : STATISTICS
'GROUP         : AUTO CORREL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function AUTO_CORREL_PARTIAL_SAMPLE_COEF_FUNC(ByRef DATA_RNG As Variant, _
ByVal NO_LAGS As Long, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim MEAN_VAL As Double

Dim FACTOR_VAL As Double

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant
Dim TEMP3_VECTOR As Variant

Dim TEMP_MATRIX As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)

NSIZE = UBound(DATA_VECTOR, 1)
ReDim TEMP_MATRIX(1 To NO_LAGS, 3)
NROWS = NSIZE + NO_LAGS

'Put data in deviations from mean form

MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
ReDim TEMP1_VECTOR(1 To NROWS, 1)
For i = 1 To NSIZE
    TEMP1_VECTOR(i, 1) = DATA_VECTOR(i, 1) - MEAN_VAL
Next i
TEMP3_VECTOR = AUTO_CORREL_CIRCULAR_LAG_FUNC(TEMP1_VECTOR, NROWS, 0, 0)
For i = 1 To NO_LAGS
    FACTOR_VAL = MATRIX_SUM_PRODUCT_FUNC(TEMP3_VECTOR, TEMP1_VECTOR) / MATRIX_SUM_PRODUCT_FUNC(TEMP3_VECTOR, TEMP3_VECTOR)
    TEMP2_VECTOR = TEMP1_VECTOR
    For j = 1 To NROWS
        TEMP1_VECTOR(j, 1) = TEMP2_VECTOR(j, 1) + (-1 * FACTOR_VAL) * TEMP3_VECTOR(j, 1)
    Next j
    For j = 1 To NROWS
        TEMP3_VECTOR(j, 1) = TEMP3_VECTOR(j, 1) + (-1 * FACTOR_VAL) * TEMP2_VECTOR(j, 1)
    Next j
    TEMP3_VECTOR = AUTO_CORREL_CIRCULAR_LAG_FUNC(TEMP3_VECTOR, NROWS, 0, 0)
    TEMP_MATRIX(i, 1) = FACTOR_VAL
    TEMP_MATRIX(i, 2) = 2 / Sqr(NSIZE) ' 2 corresponds to ~95% confidence
    TEMP_MATRIX(i, 3) = -TEMP_MATRIX(i, 2)
Next i

AUTO_CORREL_PARTIAL_SAMPLE_COEF_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
AUTO_CORREL_PARTIAL_SAMPLE_COEF_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : AUTO_CORREL_FIRST_SAMPLE_COEF_FUNC
'DESCRIPTION   : Find sample autocorrelation coefficients
'LIBRARY       : STATISTICS
'GROUP         : AUTO CORREL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function AUTO_CORREL_FIRST_SAMPLE_COEF_FUNC(ByRef DATA_RNG As Variant, _
ByVal NO_LAGS As Long, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim FACTOR_VAL As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)

NSIZE = UBound(DATA_VECTOR, 1)
ReDim TEMP_MATRIX(1 To NO_LAGS, 3)
NROWS = NSIZE + NO_LAGS

SIGMA_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)
FACTOR_VAL = SIGMA_VAL * SIGMA_VAL

' put data in deviations from mean form
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
ReDim TEMP_VECTOR(1 To NSIZE, 1)
For i = 1 To NSIZE
    TEMP_VECTOR(i, 1) = DATA_VECTOR(i, 1) - MEAN_VAL
Next i

For i = 1 To NO_LAGS
    DATA_VECTOR = AUTO_CORREL_CIRCULAR_LAG_FUNC(DATA_VECTOR, NSIZE, 0, 0)
    TEMP_MATRIX(i, 1) = MATRIX_SUM_PRODUCT_FUNC(TEMP_VECTOR, DATA_VECTOR) / (NSIZE * FACTOR_VAL)
    TEMP_MATRIX(i, 2) = 2 / Sqr(NSIZE)
    TEMP_MATRIX(i, 3) = -TEMP_MATRIX(i, 2)
Next i

AUTO_CORREL_FIRST_SAMPLE_COEF_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
AUTO_CORREL_FIRST_SAMPLE_COEF_FUNC = Err.number
End Function
    

'************************************************************************************
'************************************************************************************
'FUNCTION      : AUTO_CORREL_CIRCULAR_LAG_FUNC
'DESCRIPTION   : Circular lag function
'LIBRARY       : STATISTICS
'GROUP         : AUTO CORREL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function AUTO_CORREL_CIRCULAR_LAG_FUNC(ByRef DATA_RNG As Variant, _
ByVal NO_LAGS As Long, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NO_LAGS, 1 To 1)
TEMP_VECTOR(1, 1) = DATA_VECTOR(NO_LAGS, 1)
For i = 2 To NO_LAGS
    TEMP_VECTOR(i, 1) = DATA_VECTOR(i - 1, 1)
Next i

AUTO_CORREL_CIRCULAR_LAG_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
AUTO_CORREL_CIRCULAR_LAG_FUNC = Err.number
End Function

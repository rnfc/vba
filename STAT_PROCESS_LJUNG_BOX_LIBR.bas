Attribute VB_Name = "STAT_PROCESS_LJUNG_BOX_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : LJUNG_BOX_STATISTICS_FUNC
'DESCRIPTION   : Ljung-Box Q-Statistics & critical value
'http://en.wikipedia.org/wiki/Ljung%E2%80%93Box_test
'LIBRARY       : STATISTICS
'GROUP         : LJUNG_BOX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function LJUNG_BOX_STATISTICS_FUNC(ByRef DATA_RNG As Variant, _
ByVal NO_LAGS As Long, _
Optional ByVal CONFIDENCE_VAL As Double = 0.05, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim TEMP_SUM As Double

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

ReDim TEMP_MATRIX(1 To NO_LAGS, 1 To 2)
For j = 1 To NO_LAGS
    TEMP_VECTOR = AUTO_CORREL_FIRST_SAMPLE_COEF_FUNC(DATA_VECTOR, j, 0, 0)
    TEMP_SUM = 0
    For i = 1 To j: TEMP_SUM = TEMP_SUM + (TEMP_VECTOR(i, 1) ^ 2) / (NROWS - i): Next i
    TEMP_MATRIX(j, 1) = NROWS * (NROWS + 2) * TEMP_SUM
    TEMP_MATRIX(j, 2) = INVERSE_CHI_SQUARED_DIST_FUNC(1 - CONFIDENCE_VAL, j, False) 'Critical Value
Next j

LJUNG_BOX_STATISTICS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
LJUNG_BOX_STATISTICS_FUNC = Err.number
End Function

Attribute VB_Name = "STAT_PROCESS_BW_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : BLUNDELL_WARD_FILTER_FUNC
'DESCRIPTION   : Regression to Determine the Coefficient a1 in the
'Blundell/Ward Filter {autocorrelation test}
'LIBRARY       : STATISTICS
'GROUP         : BLUNDELL_WARD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function BLUNDELL_WARD_FILTER_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'r*(t) = 1/(1-a1) * r(t) - a1/(1-a1) * r(t-1)
'r*(t)... the "decorrelated" return time series
'r(t)... the original return time series
'r(t-1)... the lagged (by one period9 return time series
'a1... a coefficient from the regression below...
'r(T) = a0 + a1 * r(T - 1) + e(T)
'a0... a constant
'e(t)... the usual regression error term

'This function has the advantage that the mean return remains more or
'less unchanged. So calculating risk-adjusted performance will result
'in much less biased results.

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim DATE_VECTOR As Variant
Dim DATA_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim YDATA_VECTOR As Variant
Dim XDATA_VECTOR As Variant

Dim OLS_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

If DATA_TYPE <> 0 Then
    j = 1: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
End If
NROWS = UBound(DATE_VECTOR, 1)

ReDim YDATA_VECTOR(1 To NROWS - 1, 1 To 1)
ReDim XDATA_VECTOR(1 To NROWS - 1, 1 To 1)
For i = 2 To NROWS
    YDATA_VECTOR(i - 1, 1) = DATA_VECTOR(i, 1)
    XDATA_VECTOR(i - 1, 1) = DATA_VECTOR(i - 1, 1)
Next i

OLS_MATRIX = REGRESSION_SIMPLE_COEF_FUNC(XDATA_VECTOR, YDATA_VECTOR)
If OUTPUT <> 0 Then
    BLUNDELL_WARD_FILTER_FUNC = OLS_MATRIX
    Exit Function
End If

ReDim TEMP_MATRIX(0 To NROWS, 1 To 4)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "ORIGINAL SERIES"
TEMP_MATRIX(0, 3) = "LAGGED SERIES"
TEMP_MATRIX(0, 4) = "FILTERED SERIES"

i = 1
TEMP_MATRIX(i, 1) = DATE_VECTOR(i + j, 1)
TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
TEMP_MATRIX(i, 3) = ""
TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) / (1 - OLS_MATRIX(1, 1)) - TEMP_MATRIX(i, 2) * OLS_MATRIX(1, 1) / (1 - OLS_MATRIX(1, 1))
For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATE_VECTOR(i + j, 1)
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = DATA_VECTOR(i - 1, 1)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) / (1 - OLS_MATRIX(1, 1)) - (TEMP_MATRIX(i - 1, 2) * OLS_MATRIX(1, 1)) / (1 - OLS_MATRIX(1, 1))
Next i

BLUNDELL_WARD_FILTER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
BLUNDELL_WARD_FILTER_FUNC = Err.number
End Function

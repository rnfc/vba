Attribute VB_Name = "STAT_REGRESSION_MAPE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : REGRESSION_MAPE_FUNC
'DESCRIPTION   : The following function takes a set of data on two variables,
'such as prices and corresponding demands, and then estimates the
'best-fitting linear, exponential, and power curves for these data.
'It also calculates the corresponding mean absolute percentage error (MAPE).

'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_MAPE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function REGRESSION_MAPE_FUNC(ByRef YDATA_RNG As Variant, _
ByRef XDATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim EXP_SUM As Double
Dim POWER_SUM As Double
Dim LINEAR_SUM As Double

Dim LOG_YDATA As Variant
Dim LOG_XDATA As Variant

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant

Dim MAPE_MATRIX As Variant
Dim RESULT_MATRIX As Variant

On Error GoTo ERROR_LABEL

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
NROWS = UBound(YDATA_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then
    XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)
End If
If UBound(YDATA_VECTOR, 1) <> UBound(XDATA_MATRIX, 1) Then: GoTo ERROR_LABEL

ReDim LOG_YDATA(1 To NROWS, 1 To 1)
ReDim LOG_XDATA(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    LOG_YDATA(i, 1) = Log(YDATA_VECTOR(i, 1))
    LOG_XDATA(i, 1) = Log(XDATA_MATRIX(i, 1))
Next i

ReDim RESULT_MATRIX(1 To 3, 1 To 3)
'LINEAR
RESULT_MATRIX(1, 1) = REGRESSION_SIMPLE_COEF_FUNC(XDATA_MATRIX, YDATA_VECTOR)(2, 1) 'INTERCEPT
RESULT_MATRIX(2, 1) = REGRESSION_SIMPLE_COEF_FUNC(XDATA_MATRIX, YDATA_VECTOR)(1, 1) 'SLOPE

'POWER
RESULT_MATRIX(1, 2) = Exp(REGRESSION_SIMPLE_COEF_FUNC(LOG_XDATA, LOG_YDATA)(2, 1))
RESULT_MATRIX(2, 2) = REGRESSION_SIMPLE_COEF_FUNC(LOG_XDATA, LOG_YDATA)(1, 1)

'EXPONENTIAL
RESULT_MATRIX(1, 3) = Exp(REGRESSION_SIMPLE_COEF_FUNC(XDATA_MATRIX, LOG_YDATA)(2, 1))
RESULT_MATRIX(2, 3) = REGRESSION_SIMPLE_COEF_FUNC(XDATA_MATRIX, LOG_YDATA)(1, 1)

ReDim MAPE_MATRIX(1 To NROWS, 1 To 3)

For i = 1 To NROWS
    MAPE_MATRIX(i, 1) = Abs(YDATA_VECTOR(i, 1) - (RESULT_MATRIX(1, 1) + RESULT_MATRIX(2, 1) * XDATA_MATRIX(i, 1))) / YDATA_VECTOR(i, 1) 'LINEAR
    LINEAR_SUM = LINEAR_SUM + MAPE_MATRIX(i, 1)
    MAPE_MATRIX(i, 2) = Abs(YDATA_VECTOR(i, 1) - (RESULT_MATRIX(1, 2) * XDATA_MATRIX(i, 1) ^ RESULT_MATRIX(2, 2))) / YDATA_VECTOR(i, 1) 'POWER
    POWER_SUM = POWER_SUM + MAPE_MATRIX(i, 2)
    MAPE_MATRIX(i, 3) = Abs(YDATA_VECTOR(i, 1) - (RESULT_MATRIX(1, 3) * Exp(RESULT_MATRIX(2, 3) * XDATA_MATRIX(i, 1)))) / YDATA_VECTOR(i, 1) 'EXPONENTIAL
    EXP_SUM = EXP_SUM + MAPE_MATRIX(i, 3)
Next i

RESULT_MATRIX(3, 1) = LINEAR_SUM / NROWS 'MAPE
RESULT_MATRIX(3, 2) = POWER_SUM / NROWS 'MAPE
RESULT_MATRIX(3, 3) = EXP_SUM / NROWS 'MAPE

Select Case OUTPUT
Case 0
    REGRESSION_MAPE_FUNC = RESULT_MATRIX
Case 1
    REGRESSION_MAPE_FUNC = MAPE_MATRIX
Case Else
    REGRESSION_MAPE_FUNC = Array(RESULT_MATRIX, MAPE_MATRIX)
End Select

Exit Function
ERROR_LABEL:
REGRESSION_MAPE_FUNC = Err.number
End Function


Attribute VB_Name = "STAT_REGRESSION_COEF_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : REGRESSION_SIMPLE_COEF_FUNC
'DESCRIPTION   : Least-squares Regression Function
'LIBRARY       : MATRIX
'GROUP         : LEAST_SQUARE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function REGRESSION_SIMPLE_COEF_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByRef ZDATA_RNG As Variant)
    
Dim i As Long
Dim NROWS As Long

Dim SX_VAL As Double
Dim SY_VAL As Double
Dim SXX_VAL As Double
Dim SXY_VAL As Double

Dim TEMP_FACT As Double
Dim TEMP_SUM As Double

Dim TEMP_VECTOR As Variant
Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant
Dim ZDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If
YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If UBound(YDATA_VECTOR, 1) <> UBound(XDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(YDATA_VECTOR, 1)

If IsArray(ZDATA_RNG) = True Then
    ZDATA_VECTOR = ZDATA_RNG
    If UBound(ZDATA_VECTOR, 1) = 1 Then
        ZDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(ZDATA_VECTOR)
    End If
    If UBound(ZDATA_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
    For i = 1 To NROWS: ZDATA_VECTOR(i, 1) = ZDATA_VECTOR(i, 1) * ZDATA_VECTOR(i, 1): Next i
Else
    ReDim ZDATA_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS: ZDATA_VECTOR(i, 1) = 1: Next i
End If

SX_VAL = 0: SY_VAL = 0
SXX_VAL = 0: SXY_VAL = 0
TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + ZDATA_VECTOR(i, 1)
    SX_VAL = SX_VAL + XDATA_VECTOR(i, 1) * ZDATA_VECTOR(i, 1)
    SY_VAL = SY_VAL + YDATA_VECTOR(i, 1) * ZDATA_VECTOR(i, 1)
    SXX_VAL = SXX_VAL + XDATA_VECTOR(i, 1) * XDATA_VECTOR(i, 1) * ZDATA_VECTOR(i, 1)
    SXY_VAL = SXY_VAL + XDATA_VECTOR(i, 1) * YDATA_VECTOR(i, 1) * ZDATA_VECTOR(i, 1)
Next i
TEMP_FACT = 1 / (TEMP_SUM * SXX_VAL - SX_VAL * SX_VAL)

ReDim TEMP_VECTOR(1 To 5, 1 To 1)
TEMP_VECTOR(1, 1) = (TEMP_SUM * SXY_VAL - SX_VAL * SY_VAL) * TEMP_FACT 'BETA --> Slope
TEMP_VECTOR(2, 1) = (SXX_VAL * SY_VAL - SX_VAL * SXY_VAL) * TEMP_FACT 'ALPHA --> Intercept
TEMP_VECTOR(3, 1) = Sqr(TEMP_SUM * TEMP_FACT) 'SIGMA SLOPE
TEMP_VECTOR(4, 1) = Sqr(SXX_VAL * TEMP_FACT) 'SIGMA INTERCEPT
If (NROWS - 2) <> 0 Then
    TEMP_SUM = 0
    For i = 1 To NROWS
        SY_VAL = TEMP_VECTOR(2, 1) + TEMP_VECTOR(1, 1) * XDATA_VECTOR(i, 1) 'YFIT
        TEMP_SUM = TEMP_SUM + (YDATA_VECTOR(i, 1) - SY_VAL) ^ 2
    Next i
    TEMP_VECTOR(5, 1) = (TEMP_SUM / (NROWS - 2)) ^ 0.5 'RMSE
End If
REGRESSION_SIMPLE_COEF_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
REGRESSION_SIMPLE_COEF_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : REGRESSION_MULT_COEF_FUNC

'DESCRIPTION   : Coefficients for a line by using the "least squares" method
'to calculate a straight line that best fits your data, and then returns an array
'that describes the line

'LIBRARY       : MATRIX
'GROUP         : LEAST_SQUARE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 11/29/2008
'************************************************************************************
'************************************************************************************

Function REGRESSION_MULT_COEF_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = True, _
Optional ByVal MATRIX_INVERSE_TYPE As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim X_MATRIX As Variant
Dim XT_MATRIX As Variant
Dim XTX_MATRIX As Variant
Dim XTY_MATRIX As Variant

Dim XTXI_MATRIX As Variant
Dim XTXIXT_MATRIX As Variant

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1)
NCOLUMNS = UBound(XDATA_MATRIX, 2)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If NROWS <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

'----------------------------------------------------------------------------------------
Select Case INTERCEPT_FLAG
'----------------------------------------------------------------------------------------
Case True
'----------------------------------------------------------------------------------------
    ReDim X_MATRIX(1 To NROWS, 1 To NCOLUMNS + 1)
    ReDim XT_MATRIX(1 To NCOLUMNS + 1, 1 To NROWS)
    
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS + 1
            If j = 1 Then
                X_MATRIX(i, 1) = 1
                XT_MATRIX(1, i) = 1
            Else
                X_MATRIX(i, j) = XDATA_MATRIX(i, j - 1)
                XT_MATRIX(j, i) = XDATA_MATRIX(i, j - 1)
            End If
        Next j
    Next i
        
    XTX_MATRIX = MMULT_FUNC(XT_MATRIX, X_MATRIX, 70) 'X'X
    XTXI_MATRIX = MATRIX_INVERSE_FUNC(XTX_MATRIX, MATRIX_INVERSE_TYPE) 'X'X -1
    XTXIXT_MATRIX = MMULT_FUNC(XTXI_MATRIX, XT_MATRIX, 70) 'ESTIMATES

    REGRESSION_MULT_COEF_FUNC = MMULT_FUNC(XTXIXT_MATRIX, YDATA_VECTOR, 70)
'----------------------------------------------------------------------------------------
'ENTRY IN COEFFICIENTS_VECTOR(1,1) --> Intercept = Alpha
'----------------------------------------------------------------------------------------
Case False
'----------------------------------------------------------------------------------------
    XT_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)
    XTX_MATRIX = MMULT_FUNC(XT_MATRIX, XDATA_MATRIX, 70) 'X'X
    XTXI_MATRIX = MATRIX_INVERSE_FUNC(XTX_MATRIX, MATRIX_INVERSE_TYPE) 'X'X -1
    XTY_MATRIX = MMULT_FUNC(XT_MATRIX, YDATA_VECTOR, 70)

    REGRESSION_MULT_COEF_FUNC = MMULT_FUNC(XTXI_MATRIX, XTY_MATRIX, 70)
'----------------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
REGRESSION_MULT_COEF_FUNC = Err.number
End Function

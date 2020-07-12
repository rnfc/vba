Attribute VB_Name = "STAT_REGRESSION_SVD_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SVD_REGRESSION_FUNC

'DESCRIPTION   : Returns the coefficients vector [a0, a1,..am] of linear
'regression function; f = a0 + a1*x1 + a2*x2 +a3*x3 +... am*xm
'and the standard deviation of estimate
'input Y(n) dependent variable vector
'input X(n x m) range of independent variables (n > m)
'input INTERCEPT_FLAG = True/False calculate the Y intercept
'Output vector (m+1) [a0, a1, a2,...am] uses the SVD decomposition method

'LIBRARY       : MATRIX
'GROUP         : REGRESSION_SVD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_SVD_REGRESSION_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim YR_VAL As Double
Dim SSR_VAL As Double

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant

Dim TEMP_VECTOR As Variant

Dim ESR_VECTOR As Variant
Dim ALFA_VECTOR As Variant
Dim COEF_VECTOR As Variant

Dim UDATA_MATRIX As Variant
Dim SDATA_MATRIX As Variant
Dim VDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1)
NCOLUMNS = UBound(XDATA_MATRIX, 2)

ReDim COEF_VECTOR(0 To NCOLUMNS, 1 To 1)
ReDim TEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
For j = 1 To NCOLUMNS
    TEMP_VECTOR(j, 1) = 0
    For i = 1 To NROWS
        TEMP_VECTOR(j, 1) = TEMP_VECTOR(j, 1) + XDATA_MATRIX(i, j) ^ 2
    Next i
    TEMP_VECTOR(j, 1) = Sqr(TEMP_VECTOR(j, 1))
Next j

If INTERCEPT_FLAG = True Then
    ReDim UDATA_MATRIX(1 To NROWS, 1 To NCOLUMNS + 1)
    For i = 1 To NROWS
        UDATA_MATRIX(i, 1) = 1
        For j = 1 To NCOLUMNS
            UDATA_MATRIX(i, j + 1) = XDATA_MATRIX(i, j) / TEMP_VECTOR(j, 1)
        Next j
    Next i
    NSIZE = NCOLUMNS + 1
Else
    ReDim UDATA_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            UDATA_MATRIX(i, j) = XDATA_MATRIX(i, j) / TEMP_VECTOR(j, 1)
        Next j
    Next i
    NSIZE = NCOLUMNS
End If

ALFA_VECTOR = MATRIX_SVD_DECOMPOSITION_FUNC(UDATA_MATRIX, NROWS, NSIZE, 0)
UDATA_MATRIX = ALFA_VECTOR(LBound(ALFA_VECTOR))
VDATA_MATRIX = ALFA_VECTOR(LBound(ALFA_VECTOR) + 1)
SDATA_MATRIX = ALFA_VECTOR(LBound(ALFA_VECTOR) + 2)
'-------------------------------------------------------------------------------
Erase ALFA_VECTOR
ReDim ALFA_VECTOR(1 To NSIZE, 1 To 1)
'-------------------------------------------------------------------------------
For i = 1 To NSIZE
'-------------------------------------------------------------------------------
    For k = 1 To NROWS
        ALFA_VECTOR(i, 1) = ALFA_VECTOR(i, 1) + _
            UDATA_MATRIX(k, i) * YDATA_VECTOR(k, 1)
    Next k
    ALFA_VECTOR(i, 1) = ALFA_VECTOR(i, 1) / SDATA_MATRIX(i, 1)
'-------------------------------------------------------------------------------
Next i
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
If INTERCEPT_FLAG = True Then
'-------------------------------------------------------------------------------
    For i = 1 To NSIZE
        For k = 1 To NSIZE
            COEF_VECTOR(i - 1, 1) = COEF_VECTOR(i - 1, 1) + _
                ALFA_VECTOR(k, 1) * VDATA_MATRIX(i, k)
        Next k
    Next i
'-------------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------------
    COEF_VECTOR(0, 1) = 0
    For i = 1 To NSIZE
        For k = 1 To NSIZE
            COEF_VECTOR(i, 1) = COEF_VECTOR(i, 1) + _
                ALFA_VECTOR(k, 1) * VDATA_MATRIX(i, k)
        Next k
    Next i
'-------------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------------

'rescaling coefficients
For i = 1 To NCOLUMNS
    COEF_VECTOR(i, 1) = COEF_VECTOR(i, 1) / TEMP_VECTOR(i, 1)
Next i

'-------------------------------------------------------------------------------
'regression residual standard error
For i = 1 To NROWS
'-------------------------------------------------------------------------------
    YR_VAL = COEF_VECTOR(0, 1)
    For j = 1 To NCOLUMNS
        YR_VAL = YR_VAL + COEF_VECTOR(j, 1) * XDATA_MATRIX(i, j)
    Next j
    SSR_VAL = SSR_VAL + (YR_VAL - YDATA_VECTOR(i, 1)) ^ 2
'-------------------------------------------------------------------------------
Next i
'-------------------------------------------------------------------------------

If NROWS > NSIZE Then
    SSR_VAL = Sqr(SSR_VAL / (NROWS - NSIZE))
Else
    SSR_VAL = 0
End If

ReDim ESR_VECTOR(0 To NCOLUMNS, 1 To 1)

'-------------------------------------------------------------------------------
If INTERCEPT_FLAG = True Then
'-------------------------------------------------------------------------------
    For i = 0 To NCOLUMNS
        For j = 1 To NSIZE
            ESR_VECTOR(i, 1) = ESR_VECTOR(i, 1) + (VDATA_MATRIX(i + 1, j) _
                / SDATA_MATRIX(j, 1)) ^ 2
        Next j
        ESR_VECTOR(i, 1) = Sqr(ESR_VECTOR(i, 1)) * SSR_VAL
    Next i
'-------------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------------
    ESR_VECTOR(0, 1) = 0
    For i = 1 To NCOLUMNS
        For j = 1 To NSIZE
            ESR_VECTOR(i, 1) = ESR_VECTOR(i, 1) + (VDATA_MATRIX(i, j) _
                / SDATA_MATRIX(j, 1)) ^ 2
        Next j
        ESR_VECTOR(i, 1) = Sqr(ESR_VECTOR(i, 1)) * SSR_VAL
    Next i
'-------------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------------

For i = 1 To NCOLUMNS
    ESR_VECTOR(i, 1) = ESR_VECTOR(i, 1) / TEMP_VECTOR(i, 1)
Next i

ReDim TEMP_VECTOR(1 To NCOLUMNS + 1, 1 To 2)

For i = 1 To NCOLUMNS + 1
    TEMP_VECTOR(i, 1) = COEF_VECTOR(i - 1, 1)
    TEMP_VECTOR(i, 2) = ESR_VECTOR(i - 1, 1)
Next i

'ReDim TEMP_VECTOR(1 To 2, 1 To NCOLUMNS + 1)
'For i = 1 To NCOLUMNS + 1
'    TEMP_VECTOR(1, i) = COEF_VECTOR(NCOLUMNS + 1 - i,1)
'    TEMP_VECTOR(2, i) = ESR_VECTOR(NCOLUMNS + 1 - i,1)
'Next i

MATRIX_SVD_REGRESSION_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_SVD_REGRESSION_FUNC = Err.number
End Function

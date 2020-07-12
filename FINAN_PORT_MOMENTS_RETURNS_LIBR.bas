Attribute VB_Name = "FINAN_PORT_MOMENTS_RETURNS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_WEIGHTED_RETURN1_FUNC
'DESCRIPTION   : Portfolio Returns And/Or StDev
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_RETURNS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_WEIGHTED_RETURN1_FUNC(ByRef DATA_RNG As Variant, _
ByRef WEIGHTS_RNG As Variant, _
Optional ByVal COUNT_BASIS As Double = 252, _
Optional ByVal DATA_TYPE As Integer = 1, _
Optional ByVal LOG_SCALE As Integer = 1, _
Optional ByVal OUTPUT As Variant = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim WEIGHTS_VECTOR As Variant '--> Weight Vector

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then
    DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
End If
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 1) = 1 Then
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
End If
If UBound(DATA_MATRIX, 2) <> UBound(WEIGHTS_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    TEMP1_SUM = 0
    For j = 1 To NCOLUMNS
        TEMP1_SUM = TEMP1_SUM + WEIGHTS_VECTOR(j, 1) * DATA_MATRIX(i, j)
    Next j
    TEMP_VECTOR(i, 1) = TEMP1_SUM
    TEMP2_SUM = TEMP2_SUM + TEMP_VECTOR(i, 1)
Next i

TEMP1_SUM = 0
For i = 1 To NROWS
    TEMP1_SUM = TEMP1_SUM + (TEMP_VECTOR(i, 1) - (TEMP2_SUM / NROWS)) ^ 2
Next i

Select Case OUTPUT
Case 0 ' Port Mean
    PORT_WEIGHTED_RETURN1_FUNC = TEMP2_SUM / NROWS * COUNT_BASIS
Case 1 ' Port Sigma
    PORT_WEIGHTED_RETURN1_FUNC = ((TEMP1_SUM / NROWS) * COUNT_BASIS) ^ 0.5
Case 2
    PORT_WEIGHTED_RETURN1_FUNC = TEMP_VECTOR
Case Else
    PORT_WEIGHTED_RETURN1_FUNC = Array(TEMP2_SUM / NROWS * COUNT_BASIS, ((TEMP1_SUM / NROWS) * COUNT_BASIS) ^ 0.5, TEMP_VECTOR)
End Select

Exit Function
ERROR_LABEL:
PORT_WEIGHTED_RETURN1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_WEIGHTED_RETURN2_FUNC
'DESCRIPTION   : Computes return for portfolio weight from returns e
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_RETURNS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_WEIGHTED_RETURN2_FUNC(ByRef RETURNS_RNG As Variant, _
ByRef WEIGHTS_RNG As Variant)

Dim i As Long
Dim NSIZE As Long

Dim TEMP_SUM As Double
Dim WEIGHTS_VECTOR As Variant
Dim RETURNS_VECTOR As Variant

On Error GoTo ERROR_LABEL

RETURNS_VECTOR = RETURNS_RNG
If UBound(RETURNS_VECTOR, 1) = 1 Then
    RETURNS_VECTOR = MATRIX_TRANSPOSE_FUNC(RETURNS_VECTOR)
End If

WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 1) = 1 Then
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
End If
NSIZE = UBound(WEIGHTS_VECTOR, 1)
If UBound(WEIGHTS_VECTOR, 1) <> UBound(RETURNS_VECTOR, 1) Then: GoTo ERROR_LABEL
TEMP_SUM = 0
For i = 1 To NSIZE: TEMP_SUM = TEMP_SUM + WEIGHTS_VECTOR(i, 1) * RETURNS_VECTOR(i, 1): Next i
PORT_WEIGHTED_RETURN2_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
PORT_WEIGHTED_RETURN2_FUNC = Err.number
End Function

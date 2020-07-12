Attribute VB_Name = "MATRIX_DOMINANCE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_DEBUG_PRINT_FUNC
'DESCRIPTION   : Compute the dominance factor for a square matrix
'LIBRARY       : MATRIX
'GROUP         : DOMINANCE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_DOMINANCE_FACTOR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MAX As Double
Dim TEMP_MIN As Double
Dim TEMP_SUM As Double
Dim TEMP_MEAN As Double

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

TEMP_MEAN = 0
TEMP_MAX = 0
TEMP_MIN = 10 ^ 300

For i = 1 To NROWS
    TEMP_SUM = 0
    For j = 1 To NROWS
        TEMP_SUM = TEMP_SUM + Abs(DATA_MATRIX(i, j))
    Next j
    TEMP_VECTOR(i, 1) = Abs(DATA_MATRIX(i, i)) / TEMP_SUM
    TEMP_MEAN = TEMP_MEAN + TEMP_VECTOR(i, 1)
    If TEMP_VECTOR(i, 1) < TEMP_MIN Then TEMP_MIN = TEMP_VECTOR(i, 1)
    If TEMP_VECTOR(i, 1) > TEMP_MAX Then TEMP_MAX = TEMP_VECTOR(i, 1)
Next i

Select Case OUTPUT
Case 0
    MATRIX_DOMINANCE_FACTOR_FUNC = TEMP_MEAN / NROWS
Case Else
    MATRIX_DOMINANCE_FACTOR_FUNC = TEMP_VECTOR
End Select

Exit Function
ERROR_LABEL:
MATRIX_DOMINANCE_FACTOR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_DIAGONAL_DOMINANCE_FUNC
'DESCRIPTION   : This algorithm addensate the biggest values around the first diagonal
'in order to make diagonal dominant the system matrix A*x = b using only rows exchange

'LIBRARY       : MATRIX
'GROUP         : DOMINANCE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_DIAGONAL_DOMINANCE_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim TEMP_VALUE As Variant
Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant

Dim TEMP_MATRIX As Variant

Dim SWAP_FLAG As Boolean

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_RNG
DATA_VECTOR = VECTOR_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

SWAP_FLAG = False

NROWS = UBound(DATA_MATRIX, 1)
k = 0
ReDim TEMP_MATRIX(1 To NROWS ^ 2, 1 To 3)
For i = 1 To NROWS
    For j = 1 To NROWS
        k = k + 1
        TEMP_MATRIX(k, 1) = i
        TEMP_MATRIX(k, 2) = j
        TEMP_MATRIX(k, 3) = Abs(DATA_MATRIX(i, j))
    Next j
Next i
NSIZE = UBound(TEMP_MATRIX)
TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 3, 0)

l = 0
Do
    SWAP_FLAG = False
    k = 0
    Do
        k = k + 1
        i = TEMP_MATRIX(k, 1)
        j = TEMP_MATRIX(k, 2)
        TEMP_VALUE = Abs(DATA_MATRIX(i, j) * _
                DATA_MATRIX(j, i)) - Abs(DATA_MATRIX(i, i) * DATA_MATRIX(j, j))
        If TEMP_VALUE <= 0 Then
            TEMP_VALUE = Abs(DATA_MATRIX(i, j)) + _
                    Abs(DATA_MATRIX(j, i)) - Abs(DATA_MATRIX(i, i)) - _
                            Abs(DATA_MATRIX(j, j))
        End If
        If TEMP_VALUE > 0 Then 'swap the rows
            DATA_MATRIX = MATRIX_SWAP_ROW_FUNC(DATA_MATRIX, j, i)
            DATA_VECTOR = MATRIX_SWAP_ROW_FUNC(DATA_VECTOR, j, i)
            TEMP_MATRIX = MATRIX_SPARSE_SWAP_ROW_FUNC(TEMP_MATRIX, j, i, 0)
            l = l + 1
            SWAP_FLAG = True
        End If
    Loop Until k = NSIZE
Loop Until SWAP_FLAG = False Or l > NROWS

Select Case OUTPUT
    Case 0
        MATRIX_DIAGONAL_DOMINANCE_FUNC = DATA_MATRIX
    Case Else
        MATRIX_DIAGONAL_DOMINANCE_FUNC = DATA_VECTOR
End Select

Exit Function
ERROR_LABEL:
MATRIX_DIAGONAL_DOMINANCE_FUNC = Err.number
End Function

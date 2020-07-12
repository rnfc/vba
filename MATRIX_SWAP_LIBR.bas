Attribute VB_Name = "MATRIX_SWAP_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SWAP_ROW_FUNC
'DESCRIPTION   : Swaps rows k and i
'LIBRARY       : MATRIX
'GROUP         : SWAP
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SWAP_ROW_FUNC(ByRef DATA_RNG As Variant, _
ByVal k As Long, _
ByVal i As Long)

Dim j As Long
Dim TEMP_VALUE As Variant

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL
    
DATA_MATRIX = DATA_RNG

If IS_2D_ARRAY_FUNC(DATA_MATRIX) Then
    For j = LBound(DATA_MATRIX, 2) To UBound(DATA_MATRIX, 2)
        TEMP_VALUE = DATA_MATRIX(i, j)
        DATA_MATRIX(i, j) = DATA_MATRIX(k, j)
        DATA_MATRIX(k, j) = TEMP_VALUE
    Next j
ElseIf IS_1D_ARRAY_FUNC(DATA_MATRIX) Then
    For j = LBound(DATA_MATRIX, 2) To UBound(DATA_MATRIX, 2)
        TEMP_VALUE = DATA_MATRIX(i)
        DATA_MATRIX(i) = DATA_MATRIX(k)
        DATA_MATRIX(k) = TEMP_VALUE
    Next j
Else
    GoTo ERROR_LABEL
End If

MATRIX_SWAP_ROW_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SWAP_ROW_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SWAP_COLUMN_FUNC
'DESCRIPTION   : Swaps columns k and i
'LIBRARY       : MATRIX
'GROUP         : SWAP
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SWAP_COLUMN_FUNC(ByRef DATA_RNG As Variant, _
ByVal k As Long, _
ByVal j As Long)

Dim i As Long
Dim TEMP_VALUE As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL
    
DATA_MATRIX = DATA_RNG
For i = LBound(DATA_MATRIX, 1) To UBound(DATA_MATRIX, 1)
    TEMP_VALUE = DATA_MATRIX(i, j)
    DATA_MATRIX(i, j) = DATA_MATRIX(i, k)
    DATA_MATRIX(i, k) = TEMP_VALUE
Next i

MATRIX_SWAP_COLUMN_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SWAP_COLUMN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_SWAP_COLUMN_FUNC
'DESCRIPTION   : Swaps columns k and i (for complex matrix)
'LIBRARY       : MATRIX
'GROUP         : SWAP
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_SWAP_COLUMN_FUNC(ByRef DATA_RNG As Variant, _
ByVal k As Long, _
ByVal j As Long)
    
Dim i As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VALUE As Variant
Dim DATA_MATRIX As Variant
    
On Error GoTo ERROR_LABEL
    
DATA_MATRIX = DATA_RNG
    
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
For i = 1 To NROWS
    TEMP_VALUE = DATA_MATRIX(i, j)
    DATA_MATRIX(i, j) = DATA_MATRIX(i, k)
    DATA_MATRIX(i, k) = TEMP_VALUE
Next i
    '
If NCOLUMNS = 2 * NROWS Then
'complex matrix (NROWS x 2*NROWS)
    For i = 1 To NROWS
        TEMP_VALUE = DATA_MATRIX(i, j + NROWS)
        DATA_MATRIX(i, j + NROWS) = DATA_MATRIX(i, k + NROWS)
        DATA_MATRIX(i, k + NROWS) = TEMP_VALUE
    Next i
End If
    
COMPLEX_MATRIX_SWAP_COLUMN_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_SWAP_COLUMN_FUNC = Err.number
End Function

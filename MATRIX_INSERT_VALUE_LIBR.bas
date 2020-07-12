Attribute VB_Name = "MATRIX_INSERT_VALUE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ARRAY_INSERT_VALUE_FUNC
'DESCRIPTION   : Inserts a value in a sorted array in the sorted position
'LIBRARY       : MATRIX
'GROUP         : INSERT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function ARRAY_INSERT_VALUE_FUNC(ByRef DATA_MATRIX As Variant, _
ByVal REF_VALUE As Variant)
  
Dim i As Long
Dim j As Long

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim INSERT_FLAG As Boolean

Dim TEMP_VECTOR As Variant
'Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

'DATA_MATRIX = DATA_RNG

If IS_1D_ARRAY_FUNC(DATA_MATRIX) = True Then
    SROW = LBound(DATA_MATRIX, 1)
    NROWS = UBound(DATA_MATRIX, 1)
    ReDim TEMP_VECTOR(SROW To NROWS + 1)
    INSERT_FLAG = False
    j = SROW
    For i = SROW To NROWS
        If DATA_MATRIX(i) > REF_VALUE And INSERT_FLAG = False Then
            TEMP_VECTOR(j) = REF_VALUE
            j = j + 1
            INSERT_FLAG = True
        End If
        TEMP_VECTOR(j) = DATA_MATRIX(i)
        j = j + 1
    Next i
    If INSERT_FLAG = False Then: TEMP_VECTOR(j) = REF_VALUE
    ARRAY_INSERT_VALUE_FUNC = TEMP_VECTOR
ElseIf IS_2D_ARRAY_FUNC(DATA_MATRIX) = True Then
    SROW = LBound(DATA_MATRIX, 1)
    NROWS = UBound(DATA_MATRIX, 1)
    SCOLUMN = LBound(DATA_MATRIX, 2)
    NCOLUMNS = UBound(DATA_MATRIX, 2)
    ReDim TEMP_VECTOR(SROW To NROWS + 1, SCOLUMN To NCOLUMNS)
    INSERT_FLAG = False
    j = SROW
    For i = SROW To NROWS
        If DATA_MATRIX(i, 1) > REF_VALUE And INSERT_FLAG = False Then
            TEMP_VECTOR(j, 1) = REF_VALUE
            j = j + 1
            INSERT_FLAG = True
        End If
        TEMP_VECTOR(j, 1) = DATA_MATRIX(i, 1)
        j = j + 1
    Next i
    If INSERT_FLAG = False Then: TEMP_VECTOR(j, 1) = REF_VALUE
    ARRAY_INSERT_VALUE_FUNC = TEMP_VECTOR
Else
  GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
ARRAY_INSERT_VALUE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_INSERT_VALUE_FUNC
'DESCRIPTION   : Insert element in a matrix
'LIBRARY       : MATRIX
'GROUP         : INSERT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************
           
Function MATRIX_INSERT_VALUE_FUNC(ByRef DATA_MATRIX As Variant, _
ByRef DATA_VECTOR As Variant, _
Optional ByRef AROW As Long = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

On Error GoTo ERROR_LABEL

If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If (AROW) = 0 Then 'searh for the last entrie
    For i = 1 To NROWS
        If DATA_MATRIX(i, 1) = "" Then Exit For
    Next i
    AROW = i + 1
End If

If AROW > NROWS Then
    DATA_MATRIX = MATRIX_RESIZE_FUNC(DATA_MATRIX, 2 * NROWS, NCOLUMNS)
End If

For j = 1 To NCOLUMNS
    DATA_MATRIX(AROW, j) = DATA_VECTOR(j, 1)
Next j

MATRIX_INSERT_VALUE_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_INSERT_VALUE_FUNC = Err.number
End Function

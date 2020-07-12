Attribute VB_Name = "MATRIX_REPLACE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_REPLACE_DATA_FUNC
'DESCRIPTION   : Returns a matrix in which a specified substring has been replaced
'with another substring a specified number of times.

'LIBRARY       : MATRIX
'GROUP         : REPLACE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_REPLACE_DATA_FUNC(ByRef DATA_RNG As Variant, _
ByRef FROM_RNG As Variant, _
ByRef TO_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim FROM_VECTOR As Variant
Dim TO_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

FROM_VECTOR = FROM_RNG
If UBound(FROM_VECTOR, 1) = 1 Then
    FROM_VECTOR = MATRIX_TRANSPOSE_FUNC(FROM_VECTOR)
End If

TO_VECTOR = TO_RNG
If UBound(TO_VECTOR, 1) = 1 Then
    TO_VECTOR = MATRIX_TRANSPOSE_FUNC(TO_VECTOR)
End If

If UBound(FROM_VECTOR, 1) <> UBound(TO_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

For j = 1 To NCOLUMNS
    For i = 1 To NROWS
         For k = LBound(FROM_VECTOR, 1) To UBound(FROM_VECTOR, 1)
             DATA_MATRIX(i, j) = Replace(DATA_MATRIX(i, j), _
                                 FROM_VECTOR(k, 1), TO_VECTOR(k, 1), 1, -1, 0)
         Next k
    Next i
Next j

MATRIX_REPLACE_DATA_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_REPLACE_DATA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_REPLACE_VALUES_FUNC
'DESCRIPTION   : Returns a matrix in which a specified substring has been replaced
'with another substring.

'LIBRARY       : MATRIX
'GROUP         : REPLACE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_REPLACE_VALUES_FUNC(ByRef DATA_RNG As Variant, _
ByVal FROM_VALUE As Variant, _
ByVal TO_VALUE As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
k = 1
    For i = 1 To NROWS
        If DATA_MATRIX(i, j) = FROM_VALUE Then
           TEMP_MATRIX(k, j) = TO_VALUE
        Else
           TEMP_MATRIX(k, j) = DATA_MATRIX(i, j)
        End If
         k = k + 1
    Next i
Next j

MATRIX_REPLACE_VALUES_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_REPLACE_VALUES_FUNC = Err.number
End Function



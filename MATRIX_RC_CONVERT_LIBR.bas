Attribute VB_Name = "MATRIX_RC_CONVERT_LIBR"

'// PERFECT

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_VECTOR_CONVERT_FUNC
'DESCRIPTION   : This transforms the column of data to a two dimensional table.
'The block size of the data in ColumnData is specified by NSIZE.
'LIBRARY       : MATRIX
'GROUP         : RC_CONVERT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_VECTOR_CONVERT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NSIZE As Long = 5)
    
Dim i As Long            ' Row index
Dim j As Long            ' Column index
Dim k As Long

Dim NROWS As Long

Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

NROWS = UBound(DATA_VECTOR, 1)
If NROWS Mod NSIZE = 0 Then
    NROWS = CLng(NROWS / NSIZE)
Else
    NROWS = 1 + CLng(NROWS / NSIZE)
End If

ReDim TEMP_MATRIX(1 To NROWS, 1 To NSIZE)

k = 0
For i = 1 To NROWS
    For j = 1 To NSIZE
        k = k + 1
        If k > UBound(DATA_VECTOR, 1) Then: Exit For
        TEMP_MATRIX(i, j) = DATA_VECTOR(k, 1)
    Next j
Next i

MATRIX_VECTOR_CONVERT_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
MATRIX_VECTOR_CONVERT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ARRAY_CONVERT_FUNC
'DESCRIPTION   : Convert Matrix to Array
'LIBRARY       : MATRIX
'GROUP         : RC_CONVERT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ARRAY_CONVERT_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_ARR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)
SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

NSIZE = (NROWS - SROW + 1) * (NCOLUMNS - SCOLUMN + 1)

ReDim TEMP_ARR(SROW To NSIZE)

k = SROW
For j = SCOLUMN To NCOLUMNS
    For i = SROW To NROWS
        TEMP_ARR(k) = DATA_MATRIX(i, j)
        k = k + 1
    Next i
Next j

MATRIX_ARRAY_CONVERT_FUNC = TEMP_ARR

Exit Function
ERROR_LABEL:
MATRIX_ARRAY_CONVERT_FUNC = Err.number
End Function

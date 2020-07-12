Attribute VB_Name = "MATRIX_RC_GET_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GET_ROW_FUNC
'DESCRIPTION   : Extract row from array
'LIBRARY       : MATRIX
'GROUP         : RC_GET
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

'// PERFECT

Function MATRIX_GET_ROW_FUNC(ByRef DATA_RNG As Variant, _
ByVal AROW As Long, _
ByVal BASE_ROW As Long)

Dim i As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(BASE_ROW To BASE_ROW, SCOLUMN To NCOLUMNS)

For i = SCOLUMN To NCOLUMNS
     TEMP_MATRIX(BASE_ROW, i) = DATA_MATRIX(AROW, i)
Next i

MATRIX_GET_ROW_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_GET_ROW_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GET_COLUMN_FUNC
'DESCRIPTION   : Extract column from array
'LIBRARY       : MATRIX
'GROUP         : RC_GET
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

'// PERFECT

Function MATRIX_GET_COLUMN_FUNC(ByRef DATA_RNG As Variant, _
ByVal ACOLUMN As Long, _
ByVal BASE_COLUMN As Long)

Dim i As Long

Dim SROW As Long
Dim NROWS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(SROW To NROWS, BASE_COLUMN To BASE_COLUMN)
For i = SROW To NROWS
     TEMP_MATRIX(i, BASE_COLUMN) = DATA_MATRIX(i, ACOLUMN)
Next i

MATRIX_GET_COLUMN_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_GET_COLUMN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ARRAY_GET_VECTOR_FUNC
'DESCRIPTION   : Extract sub-matrix from array
'LIBRARY       : MATRIX
'GROUP         : RC_GET
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

'// PERFECT

Function ARRAY_GET_VECTOR_FUNC(ByRef DATA_RNG As Variant, _
ByVal SCOLUMN As Long, _
ByVal NCOLUMNS As Long, _
ByVal SROW As Long, _
ByVal NROWS As Long)

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

'----------------------------------------------------------------------
If IS_2D_ARRAY_FUNC(DATA_MATRIX) = True Then
'----------------------------------------------------------------------

    ReDim TEMP_MATRIX(1 To (NROWS - SROW + 1), 1 To (NCOLUMNS - SCOLUMN + 1))
    jj = 1
    For j = SCOLUMN To NCOLUMNS
        ii = 1
        For i = SROW To NROWS
             TEMP_MATRIX(ii, jj) = DATA_MATRIX(i, j)
             ii = ii + 1
        Next i
        jj = jj + 1
    Next j
'----------------------------------------------------------------------
ElseIf IS_1D_ARRAY_FUNC(DATA_MATRIX) = True Then
'----------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To (NROWS - SROW + 1), 1 To 1)
    j = 1
    For i = SROW To NROWS
        TEMP_MATRIX(j, 1) = DATA_MATRIX(i)
        j = j + 1
    Next i
'----------------------------------------------------------------------
Else
'----------------------------------------------------------------------
    GoTo ERROR_LABEL
'----------------------------------------------------------------------
End If
'----------------------------------------------------------------------

ARRAY_GET_VECTOR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ARRAY_GET_VECTOR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GET_DIAGONAL_FUNC
'DESCRIPTION   : Extract diagonal from matrix
'LIBRARY       : MATRIX
'GROUP         : RC_GET
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_GET_DIAGONAL_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NSIZE As Long = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If ((NSIZE > 0 And NSIZE > NCOLUMNS) Or _
    (NSIZE < 0 And NSIZE < -NROWS)) Then
    GoTo ERROR_LABEL ' --> requested diagonal out of range
End If

ReDim TEMP_MATRIX(SROW To NROWS, SCOLUMN To NCOLUMNS)

l = MINIMUM_FUNC(NCOLUMNS, NROWS + NSIZE)
For j = SCOLUMN To l
    k = MAXIMUM_FUNC(SROW, j - NSIZE)
    For i = k To NROWS
        TEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
    Next i
Next j
  
MATRIX_GET_DIAGONAL_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_GET_DIAGONAL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GET_SUB_MATRIX_FUNC
'DESCRIPTION   : Return the sub matrix of pivot ij
'LIBRARY       : MATRIX
'GROUP         : RC_GET
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_GET_SUB_MATRIX_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal SROW As Long = -1, _
Optional ByVal NROWS As Long = -1, _
Optional ByVal SCOLUMN As Long = -1, _
Optional ByVal NCOLUMNS As Long = -1)
  
Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL
  
DATA_MATRIX = DATA_RNG
If SROW = -1 Then: SROW = LBound(DATA_MATRIX, 1)
If NROWS = -1 Then: NROWS = UBound(DATA_MATRIX, 1)

If SCOLUMN = -1 Then: SCOLUMN = LBound(DATA_MATRIX, 2)
If NCOLUMNS = -1 Then: NCOLUMNS = UBound(DATA_MATRIX, 2)
  
ReDim TEMP_MATRIX(1 To NROWS - SROW + 1, 1 To NCOLUMNS - SCOLUMN + 1)
ii = 1: jj = 1
For j = SCOLUMN To NCOLUMNS
    ii = 1
    For i = SROW To NROWS
      TEMP_MATRIX(ii, jj) = DATA_MATRIX(i, j)
      ii = ii + 1
    Next i
    jj = jj + 1
Next j

MATRIX_GET_SUB_MATRIX_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_GET_SUB_MATRIX_FUNC = Err.number
End Function

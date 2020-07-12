Attribute VB_Name = "MATRIX_RC_REMOVE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_REMOVE_ROWS_FUNC
'DESCRIPTION   : Delete row(s) from array
'LIBRARY       : MATRIX
'GROUP         : RC_REMOVE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

'// PERFECT

Function MATRIX_REMOVE_ROWS_FUNC(ByRef DATA_RNG As Variant, _
ByVal START_ROW As Long, _
Optional ByVal NO_ROWS As Long = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
  
If NO_ROWS = 0 Then
    MATRIX_REMOVE_ROWS_FUNC = DATA_MATRIX
    Exit Function
End If

SROW = LBound(DATA_MATRIX, 1)
SCOLUMN = LBound(DATA_MATRIX, 2)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If (NO_ROWS < 0) Then: NO_ROWS = 1
If NO_ROWS >= (NROWS - SROW) + 1 Then: GoTo ERROR_LABEL
  
ReDim TEMP_MATRIX(SROW To (NROWS - NO_ROWS), SCOLUMN To NCOLUMNS)
    
For j = SCOLUMN To NCOLUMNS
    k = SROW
    For i = SROW To NROWS
        If i = START_ROW Then: i = i + NO_ROWS
        If (i > NROWS) Or (k > UBound(TEMP_MATRIX, 1)) Then: GoTo 1983
        TEMP_MATRIX(k, j) = DATA_MATRIX(i, j)
        k = k + 1
    Next i
1983:
Next j
  
MATRIX_REMOVE_ROWS_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_REMOVE_ROWS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_REMOVE_COLUMNS_FUNC
'DESCRIPTION   : Delete column(s) from array
'LIBRARY       : MATRIX
'GROUP         : RC_REMOVE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

'// PERFECT

Function MATRIX_REMOVE_COLUMNS_FUNC(ByRef DATA_RNG As Variant, _
ByVal START_COLUMN As Long, _
Optional ByVal NO_COLUMNS As Long = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
  
If NO_COLUMNS = 0 Then
    MATRIX_REMOVE_COLUMNS_FUNC = DATA_MATRIX
    Exit Function
End If

SROW = LBound(DATA_MATRIX, 1)
SCOLUMN = LBound(DATA_MATRIX, 2)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If (NO_COLUMNS < 0) Then: NO_COLUMNS = 1
If NO_COLUMNS >= (NCOLUMNS - SCOLUMN) + 1 Then: GoTo ERROR_LABEL
        
ReDim TEMP_MATRIX(SROW To NROWS, SCOLUMN To (NCOLUMNS - NO_COLUMNS))
    
For j = SROW To NROWS
    k = SCOLUMN
    For i = SCOLUMN To NCOLUMNS
        If i = START_COLUMN Then: i = i + NO_COLUMNS
        If (i > NCOLUMNS) Or (k > UBound(TEMP_MATRIX, 2)) Then: GoTo 1983
        TEMP_MATRIX(j, k) = DATA_MATRIX(j, i)
        k = k + 1
    Next i
1983:
Next j

MATRIX_REMOVE_COLUMNS_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_REMOVE_COLUMNS_FUNC = Err.number
End Function

Attribute VB_Name = "MATRIX_COUNT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_COUNT_NUMERICS_FUNC
'DESCRIPTION   : COUNT NUMERIC ENTRIES
'LIBRARY       : MATRIX
'GROUP         : COUNT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_COUNT_NUMERICS_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG

If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If ARRAY_DIMENSION_FUNC(DATA_VECTOR) < 2 Then: _
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)
j = 0
For i = 1 To NROWS
    If IsNumeric(DATA_VECTOR(i, 1)) And Not IsEmpty(DATA_VECTOR(i, 1)) Then
        j = j + 1
    End If
Next i
VECTOR_COUNT_NUMERICS_FUNC = j

Exit Function
ERROR_LABEL:
VECTOR_COUNT_NUMERICS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_COUNT_UNIQUES_FUNC
'DESCRIPTION   : COUNT UNIQUE ENTRIES WITHIN A VECTOR
'LIBRARY       : MATRIX
'GROUP         : COUNT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_COUNT_UNIQUES_FUNC(ByRef DATA_RNG As Variant)
'0 = ""
    Dim i As Long
    Dim j As Long
    Dim NROWS As Long
    
    Dim TEMP_VALUE As Variant
    Dim DATA_VECTOR As Variant
    
    On Error GoTo ERROR_LABEL
    
    DATA_VECTOR = DATA_RNG
    
    NROWS = UBound(DATA_VECTOR, 1)
    
    DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
    
    j = 0
    TEMP_VALUE = DATA_VECTOR(1, 1)

    For i = 1 To NROWS
         If DATA_VECTOR(i, 1) > TEMP_VALUE Then
            j = j + 1
            TEMP_VALUE = DATA_VECTOR(i, 1)
         End If
    Next i

    VECTOR_COUNT_UNIQUES_FUNC = j + 1
    
Exit Function
ERROR_LABEL:
VECTOR_COUNT_UNIQUES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_COUNT_BLANKS_FUNC
'DESCRIPTION   : COUNT BLANK ENTRIES WITHIN ARRAY
'LIBRARY       : STATISTICS
'GROUP         : COUNT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_COUNT_BLANKS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal SROW As Long = 1, _
Optional ByVal SCOLUMN As Long = 1)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If SROW = 0 Then SROW = LBound(DATA_MATRIX, 1)
If SCOLUMN = 0 Then SCOLUMN = LBound(DATA_MATRIX, 2)

j = 0
For i = SROW To NROWS
   If IsEmpty(DATA_MATRIX(i, SCOLUMN)) Or _
   (DATA_MATRIX(i, SCOLUMN) = "") Or _
   (DATA_MATRIX(i, SCOLUMN) = " ") Then
        j = j + 1
   End If
Next i

MATRIX_COUNT_BLANKS_FUNC = j

Exit Function
ERROR_LABEL:
MATRIX_COUNT_BLANKS_FUNC = Err.number
End Function


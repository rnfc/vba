Attribute VB_Name = "MATRIX_RC_TRANSPOSE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRANSPOSE_FUNC
'DESCRIPTION   : TRANSPOSE AN ARRAY (FROM N X M TO M X N)
'LIBRARY       : MATRIX
'GROUP         : RC_TRANSPOSE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRANSPOSE_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

If IS_2D_ARRAY_FUNC(DATA_MATRIX) = True Then
    SROW = LBound(DATA_MATRIX, 1)
    SCOLUMN = LBound(DATA_MATRIX, 2)

    NROWS = UBound(DATA_MATRIX, 1)
    NCOLUMNS = UBound(DATA_MATRIX, 2)

    ReDim TEMP_MATRIX(SCOLUMN To NCOLUMNS, SROW To NROWS)
    For j = SCOLUMN To NCOLUMNS
        For i = SROW To NROWS
            TEMP_MATRIX(j, i) = DATA_MATRIX(i, j)
        Next i
    Next j
Else
    If IS_1D_ARRAY_FUNC(DATA_MATRIX) Then

        SROW = LBound(DATA_MATRIX)
        NROWS = UBound(DATA_MATRIX)
        
        ReDim TEMP_MATRIX(SROW To NROWS, SROW To SROW)
        For i = SROW To NROWS
            TEMP_MATRIX(i, SROW) = DATA_MATRIX(i)
        Next i
    Else
        GoTo ERROR_LABEL
    End If
End If

MATRIX_TRANSPOSE_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_TRANSPOSE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_REVERSE_FUNC
'DESCRIPTION   : REVERSE THE ENTRIES IN AN ARRAY
'LIBRARY       : MATRIX
'GROUP         : RC_TRANSPOSE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_REVERSE_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

'-------------------------------------------------------------------------
If IS_2D_ARRAY_FUNC(DATA_MATRIX) = True Then
'-------------------------------------------------------------------------
    NCOLUMNS = UBound(DATA_MATRIX, 2)
    NROWS = UBound(DATA_MATRIX, 1)

    SCOLUMN = LBound(DATA_MATRIX, 2)
    SROW = LBound(DATA_MATRIX, 1)

    ReDim TEMP_MATRIX(SROW To NROWS, SCOLUMN To NCOLUMNS)

    For j = SCOLUMN To NCOLUMNS
        For i = SROW To NROWS
            TEMP_MATRIX(NROWS + SROW - i, j) = DATA_MATRIX(i, j)
        Next i
    Next j
'-------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------
    If IS_1D_ARRAY_FUNC(DATA_MATRIX) Then
    
        SROW = LBound(DATA_MATRIX, 1)
        NROWS = UBound(DATA_MATRIX, 1)
        
        ReDim TEMP_MATRIX(SROW To NROWS, SROW To SROW)
        
        For j = SCOLUMN To NCOLUMNS
            For i = SROW To NROWS
                TEMP_MATRIX(NROWS + SROW - i, SROW) = DATA_MATRIX(i)
            Next i
        Next j
    Else
        GoTo ERROR_LABEL
    End If
'-------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------

MATRIX_REVERSE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_REVERSE_FUNC = Err.number
End Function

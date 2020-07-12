Attribute VB_Name = "MATRIX_RC_BASE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CHANGE_BASE_ZERO_FUNC
'DESCRIPTION   : Change the base of the array to zero (for rows & columns)
'LIBRARY       : MATRIX
'GROUP         : RC_BASE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CHANGE_BASE_ZERO_FUNC(ByRef DATA_RNG As Variant)

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
NROWS = UBound(DATA_MATRIX, 1) - LBound(DATA_MATRIX, 1) + 1
NCOLUMNS = UBound(DATA_MATRIX, 2) - LBound(DATA_MATRIX, 2) + 1

ReDim TEMP_MATRIX(0 To NROWS - 1, 0 To NCOLUMNS - 1)

SROW = 0
For i = LBound(DATA_MATRIX, 1) To UBound(DATA_MATRIX, 1)
    SCOLUMN = 0
    For j = LBound(DATA_MATRIX, 2) To UBound(DATA_MATRIX, 2)
        TEMP_MATRIX(SROW, SCOLUMN) = DATA_MATRIX(i, j)
        SCOLUMN = SCOLUMN + 1
    Next j
    SROW = SROW + 1
Next i
MATRIX_CHANGE_BASE_ZERO_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CHANGE_BASE_ZERO_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CHANGE_BASE_ZERO_FUNC
'DESCRIPTION   : Change the base of the array to one (for rows & columns)
'LIBRARY       : MATRIX
'GROUP         : RC_BASE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_CHANGE_BASE_ONE_FUNC(ByRef DATA_RNG As Variant)

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
NROWS = UBound(DATA_MATRIX, 1) - LBound(DATA_MATRIX, 1) + 1
NCOLUMNS = UBound(DATA_MATRIX, 2) - LBound(DATA_MATRIX, 2) + 1

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

SROW = 1
For i = LBound(DATA_MATRIX, 1) To UBound(DATA_MATRIX, 1)
    SCOLUMN = 1
    For j = LBound(DATA_MATRIX, 2) To UBound(DATA_MATRIX, 2)
        TEMP_MATRIX(SROW, SCOLUMN) = DATA_MATRIX(i, j)
        SCOLUMN = SCOLUMN + 1
    Next j
    SROW = SROW + 1
Next i

MATRIX_CHANGE_BASE_ONE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CHANGE_BASE_ONE_FUNC = Err.number
End Function

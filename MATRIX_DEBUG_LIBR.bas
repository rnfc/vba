Attribute VB_Name = "MATRIX_DEBUG_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_DEBUG_PRINT_FUNC
'DESCRIPTION   : Print the entries of a vector in the Immediate Windows
'LIBRARY       : MATRIX
'GROUP         : DEBUG
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_DEBUG_PRINT_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim SROW As Long
Dim NROWS As Long

Dim DATA_ARR As Variant

On Error GoTo ERROR_LABEL
    
VECTOR_DEBUG_PRINT_FUNC = False
    
DATA_ARR = DATA_RNG
SROW = LBound(DATA_ARR)
NROWS = UBound(DATA_ARR)

For i = SROW To NROWS
    Debug.Print DATA_ARR(i)
Next i

VECTOR_DEBUG_PRINT_FUNC = True

Exit Function
ERROR_LABEL:
VECTOR_DEBUG_PRINT_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_DEBUG_PRINT_FUNC
'DESCRIPTION   : Print the entries of a matrix in the Immediate Windows
'LIBRARY       : MATRIX
'GROUP         : DEBUG
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_DEBUG_PRINT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal TAB_FACTOR As Double = 24)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

MATRIX_DEBUG_PRINT_FUNC = False
DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

For i = SROW To NROWS
    For j = SCOLUMN To NCOLUMNS
        Debug.Print DATA_MATRIX(i, j); Tab(TAB_FACTOR * j);
    Next j
Next i

MATRIX_DEBUG_PRINT_FUNC = True

Exit Function
ERROR_LABEL:
MATRIX_DEBUG_PRINT_FUNC = False
End Function

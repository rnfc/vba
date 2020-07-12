Attribute VB_Name = "MATRIX_ARITHM_POWER_ROOT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_POWER_FUNC
'DESCRIPTION   : Returns the result of a matrix raised to a power
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_POWER_ROOT Power
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_POWER_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal POWER_VAL As Double = 2, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

If POWER_VAL = 1 Then
    MATRIX_ELEMENTS_POWER_FUNC = DATA_MATRIX
    Exit Function
End If
    
'---------------------------------------------------------------------------
Select Case VERSION
'---------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------
    NROWS = UBound(DATA_MATRIX, 1)
    NCOLUMNS = UBound(DATA_MATRIX, 2)
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            TEMP_MATRIX(i, j) = DATA_MATRIX(i, j) ^ POWER_VAL
        Next j
    Next i
    MATRIX_ELEMENTS_POWER_FUNC = TEMP_MATRIX
'---------------------------------------------------------------------------
Case 1
'---------------------------------------------------------------------------
    MATRIX_ELEMENTS_POWER_FUNC = MMULT_FUNC(MATRIX_ELEMENTS_POWER_FUNC(DATA_MATRIX, POWER_VAL - 1, 1), DATA_MATRIX)
'---------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------
    TEMP_MATRIX = DATA_MATRIX
    For i = 1 To POWER_VAL - 1
        TEMP_MATRIX = MMULT_FUNC(DATA_MATRIX, TEMP_MATRIX)
    Next i
    MATRIX_ELEMENTS_POWER_FUNC = TEMP_MATRIX
'---------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_POWER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_ELEMENTS_SQUARE_ROOT_FUNC
'DESCRIPTION   : Returns the square root of each entry in a vector
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_POWER_ROOT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_ELEMENTS_SQUARE_ROOT_FUNC(ByRef DATA_RNG As Variant)
    
Dim i As Long
Dim NROWS As Long
Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: _
DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    TEMP_VECTOR(i, 1) = Sqr(DATA_VECTOR(i, 1))
Next i
VECTOR_ELEMENTS_SQUARE_ROOT_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_ELEMENTS_SQUARE_ROOT_FUNC = Err.number
End Function

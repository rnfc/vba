Attribute VB_Name = "MATRIX_ARITHM_SIGN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_SIGN_VALUES_FUNC
'DESCRIPTION   : Returns the Sign of each entry in a vector
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_SIGN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Function VECTOR_SIGN_VALUES_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)
    
Dim i As Long
    
Dim SROW As Long
Dim NROWS As Long
    
Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant
    
On Error GoTo ERROR_LABEL
    
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: _
DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
    
SROW = LBound(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)
    
'------------------------------------------------------------------
Select Case VERSION
'------------------------------------------------------------------
Case 0
'------------------------------------------------------------------
    If IS_1D_ARRAY_FUNC(DATA_VECTOR) Then
        ReDim TEMP_VECTOR(SROW To NROWS)
        For i = SROW To NROWS
            TEMP_VECTOR(i) = Sgn(DATA_VECTOR(i))
        Next i
    ElseIf IS_2D_ARRAY_FUNC(DATA_VECTOR) Then
        ReDim TEMP_VECTOR(SROW To NROWS)
        For i = SROW To NROWS
            TEMP_VECTOR(i) = Sgn(DATA_VECTOR(i, 1))
        Next i
    Else
        GoTo ERROR_LABEL
    End If
'------------------------------------------------------------------
Case Else
'------------------------------------------------------------------
    If IS_1D_ARRAY_FUNC(DATA_VECTOR) Then
        ReDim TEMP_VECTOR(SROW To NROWS, 1 To 1)
        For i = SROW To NROWS
            TEMP_VECTOR(i, 1) = Sgn(DATA_VECTOR(i))
        Next i
    ElseIf IS_2D_ARRAY_FUNC(DATA_VECTOR) Then
        ReDim TEMP_VECTOR(SROW To NROWS, 1 To 1)
        For i = SROW To NROWS
            TEMP_VECTOR(i, 1) = Sgn(DATA_VECTOR(i, 1))
        Next i
    Else
        GoTo ERROR_LABEL
    End If
'------------------------------------------------------------------
'------------------------------------------------------------------
End Select

VECTOR_SIGN_VALUES_FUNC = TEMP_VECTOR
Exit Function
ERROR_LABEL:
VECTOR_SIGN_VALUES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CHANGE_SIGN_COLUMN_FUNC
'DESCRIPTION   : Change sign in a column
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_SIGN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_CHANGE_SIGN_COLUMN_FUNC(ByRef DATA_RNG As Variant, _
ByVal j As Long)

Dim i As Long
Dim SROW As Long
Dim NROWS As Long
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

For i = SROW To NROWS
    DATA_MATRIX(i, j) = -DATA_MATRIX(i, j)
Next i

MATRIX_CHANGE_SIGN_COLUMN_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CHANGE_SIGN_COLUMN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CHANGE_SIGN_ROW_FUNC
'DESCRIPTION   : Change sign in a row
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_SIGN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CHANGE_SIGN_ROW_FUNC(ByRef DATA_RNG As Variant, _
ByVal i As Long)

Dim j As Long
Dim SCOLUMN As Long
Dim NCOLUMNS As Long
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

For j = SCOLUMN To NCOLUMNS
    DATA_MATRIX(i, j) = -DATA_MATRIX(i, j)
Next j

MATRIX_CHANGE_SIGN_ROW_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CHANGE_SIGN_ROW_FUNC = Err.number
End Function

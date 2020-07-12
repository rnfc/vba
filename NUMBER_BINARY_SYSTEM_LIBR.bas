Attribute VB_Name = "NUMBER_BINARY_SYSTEM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_SYSTEM_SOLVER_FUNC
'DESCRIPTION   : Binary Systems support function
'LIBRARY       : NUMBER_BINARY
'GROUP         : SYSTEM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function BINARY_SYSTEM_SOLVER_FUNC(ByRef BIN_MAT_RNG As Variant, _
ByRef BIN_VEC_RNG As Variant, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal SWITCH_FLAG As Boolean = False)

Dim i As Long
Dim j As Long

Dim NCOLUMNS As Long

Dim TEMP_VALUE As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

Dim BINARY_MATRIX As Variant
Dim BINARY_VECTOR As Variant

On Error GoTo ERROR_LABEL

BINARY_MATRIX = BIN_MAT_RNG

BINARY_VECTOR = BIN_VEC_RNG
If UBound(BINARY_VECTOR, 2) > UBound(BINARY_VECTOR, 1) Then
    BINARY_VECTOR = MATRIX_TRANSPOSE_FUNC(BINARY_VECTOR)
    'transform any vector in a vertical vector
End If

NCOLUMNS = UBound(BINARY_MATRIX, 2)  'number of variables
NROWS = UBound(BINARY_MATRIX, 1)   'number of equations

If NROWS <> UBound(BINARY_VECTOR, 1) Then: GoTo ERROR_LABEL
'matrix and vector must be the same rows

If UBound(BINARY_VECTOR, 2) <> 1 Then: GoTo ERROR_LABEL
'constant term must be a vector

DATA_MATRIX = CALL_BINARY_SYSTEM_SOLVER_FUNC(BINARY_MATRIX, BINARY_VECTOR)

NROWS = UBound(DATA_MATRIX, 1)


If NROWS = 0 Then: GoTo ERROR_LABEL ' no solutions found"

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        If VERSION = 0 Then
            TEMP_VALUE = DATA_MATRIX(i, j) 'least significant digit to left
        Else
            TEMP_VALUE = DATA_MATRIX(i, NCOLUMNS - j + 1)
            'least significant digit to right
        End If
        If SWITCH_FLAG = True Then 'change 1-0 with true/false
            If TEMP_VALUE = 1 Then
                TEMP_VALUE = True
            Else
                TEMP_VALUE = False
            End If
        End If
        TEMP_MATRIX(i, j) = TEMP_VALUE
    Next i
Next j

BINARY_SYSTEM_SOLVER_FUNC = TEMP_MATRIX
     
Exit Function
ERROR_LABEL:
BINARY_SYSTEM_SOLVER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_BINARY_SYSTEM_SOLVER_FUNC
'DESCRIPTION   : Solve binary system of n equation and m variables with
'brute force attack (adapted for m < 12)
'LIBRARY       : NUMBER_BINARY
'GROUP         : SYSTEM
'ID            : 002
'LAST UPDATE   : 12/08/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

' BINARY_MATRIX: binary matrix system (n x m)
' BINARY_VECTOR: constant binary terms vector  (n x 1)

Private Function CALL_BINARY_SYSTEM_SOLVER_FUNC(ByRef BIN_MAT_RNG As Variant, _
ByRef BIN_VEC_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim BINARY_MATRIX As Variant
Dim BINARY_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

BINARY_MATRIX = BIN_MAT_RNG
BINARY_VECTOR = BIN_VEC_RNG
If UBound(BINARY_VECTOR, 1) = 1 Then: _
    BINARY_VECTOR = MATRIX_TRANSPOSE_FUNC(BINARY_VECTOR)

NROWS = UBound(BINARY_MATRIX, 1)     'number of equations
NCOLUMNS = UBound(BINARY_MATRIX, 2)    'number of variables

ATEMP_VECTOR = BINARY_GENERATOR_FUNC(NCOLUMNS, 0)
BTEMP_VECTOR = BINARY_MATRIX_MULT_FUNC(BINARY_MATRIX, ATEMP_VECTOR)

ReDim DATA_MATRIX(1 To UBound(ATEMP_VECTOR, 1), 1 To NCOLUMNS)
'search for solutions

l = 0
For i = 1 To UBound(BTEMP_VECTOR, 1) 'check BTEMP_VECTOR = BINARY_VECTOR
    For j = 1 To NROWS
        If BTEMP_VECTOR(i, j) <> BINARY_VECTOR(j, 1) Then Exit For
    Next j
    If j > NROWS Then 'one solution find
        l = l + 1
        For k = 1 To NCOLUMNS
            DATA_MATRIX(l, k) = ATEMP_VECTOR(i, k)
        Next k
    End If
Next i
NSIZE = l
If NSIZE > 0 Then
    'return solution find
    ReDim TEMP_MATRIX(1 To NSIZE, 1 To NCOLUMNS)
    For i = 1 To NSIZE
        For j = 1 To NCOLUMNS
            TEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
        Next j
    Next i
Else
    ReDim TEMP_MATRIX(0, 0) 'no solution find
End If

CALL_BINARY_SYSTEM_SOLVER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_BINARY_SYSTEM_SOLVER_FUNC = Err.number
End Function

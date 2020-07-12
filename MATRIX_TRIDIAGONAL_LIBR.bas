Attribute VB_Name = "MATRIX_TRIDIAGONAL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIDIAGONAL_VALIDATE_FUNC
'DESCRIPTION   : Check if a matrix is tridiagonal
'LIBRARY       : MATRIX
'GROUP         : TRIDIAGONAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRIDIAGONAL_VALIDATE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal epsilon As Double = 5 * 10 ^ -16)

'1 = tridiagonal
'2 = tridiagonal uniform
'0 = else

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim UNIT_FLAG As Boolean
Dim TRIDIAG_FLAG As Boolean

Dim DATA_MATRIX As Variant
    
On Error GoTo ERROR_LABEL
    
DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
     
If NROWS <> NCOLUMNS Then GoTo 1984
For i = 1 To NROWS
    For j = i + 2 To NROWS
        If Abs(DATA_MATRIX(i, j)) > epsilon Then GoTo 1983
        If Abs(DATA_MATRIX(j, i)) > epsilon Then GoTo 1983
    Next j
Next i
TRIDIAG_FLAG = True
     
1983: If TRIDIAG_FLAG Then
        For i = 2 To NROWS - 1
            If Abs(DATA_MATRIX(i, i + 1) - _
                    DATA_MATRIX(i - 1, i)) > epsilon Then GoTo 1984
            If Abs(DATA_MATRIX(i + 1, i) - _
                    DATA_MATRIX(i, i - 1)) > epsilon Then GoTo 1984
        Next i
    
        For i = 2 To NROWS
            If Abs(DATA_MATRIX(i, i) - _
                    DATA_MATRIX(i - 1, i - 1)) > epsilon Then GoTo 1984
        Next i
        UNIT_FLAG = True
     End If
    
1984: If TRIDIAG_FLAG Then
        If UNIT_FLAG Then k = 2 Else k = 1
     Else
        k = 0
     End If
     
     MATRIX_TRIDIAGONAL_VALIDATE_FUNC = k

Exit Function
ERROR_LABEL:
MATRIX_TRIDIAGONAL_VALIDATE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIDIAGONAL_CONVERT_FUNC
'DESCRIPTION   : Convert into a tridiagonal matrix
'LIBRARY       : MATRIX
'GROUP         : TRIDIAGONAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_TRIDIAGONAL_CONVERT_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)

If NCOLUMNS = 3 Then
    For i = 1 To NROWS
        If i > 1 Then TEMP_MATRIX(i, i - 1) = DATA_MATRIX(i, 1)
         TEMP_MATRIX(i, i) = DATA_MATRIX(i, 2)
        If i < NROWS Then TEMP_MATRIX(i, i + 1) = DATA_MATRIX(i, 3)
    Next i
    MATRIX_TRIDIAGONAL_CONVERT_FUNC = TEMP_MATRIX
ElseIf NROWS = NCOLUMNS Then 'load from a square matrix
    MATRIX_TRIDIAGONAL_CONVERT_FUNC = DATA_MATRIX
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
MATRIX_TRIDIAGONAL_CONVERT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIDIAGONAL_LOAD_FUNC
'DESCRIPTION   : Load from square tridiagonal matrix
'LIBRARY       : MATRIX
'GROUP         : TRIDIAGONAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRIDIAGONAL_LOAD_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 3)

If NROWS = NCOLUMNS Then
    For i = 1 To NROWS
        If i > 1 Then TEMP_MATRIX(i, 1) = DATA_MATRIX(i, i - 1)
        TEMP_MATRIX(i, 2) = DATA_MATRIX(i, i)
        If i < NROWS Then TEMP_MATRIX(i, 3) = DATA_MATRIX(i, i + 1)
    Next i
ElseIf NCOLUMNS = 3 Then 'load from 3 vectors
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
        TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
        TEMP_MATRIX(i, 3) = DATA_MATRIX(i, 3)
    Next i
Else
    GoTo ERROR_LABEL
End If

MATRIX_TRIDIAGONAL_LOAD_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_TRIDIAGONAL_LOAD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIDIAGONAL_SYMMETRIZE_FUNC
'DESCRIPTION   : Symmetrize tridiagonal matrix
'LIBRARY       : MATRIX
'GROUP         : TRIDIAGONAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRIDIAGONAL_SYMMETRIZE_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim SCALE_VAL As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)

For i = 2 To NROWS
    If DATA_MATRIX(i - 1, 3) = 0 Then
        DATA_MATRIX(i, 1) = 0
    ElseIf DATA_MATRIX(i, 1) = 0 Then
        DATA_MATRIX(i - 1, 3) = 0
    Else
        SCALE_VAL = Sqr(DATA_MATRIX(i, 1) / DATA_MATRIX(i - 1, 3))
        DATA_MATRIX(i, 1) = DATA_MATRIX(i, 1) / SCALE_VAL
        DATA_MATRIX(i - 1, 3) = DATA_MATRIX(i - 1, 3) * SCALE_VAL
    End If
Next

MATRIX_TRIDIAGONAL_SYMMETRIZE_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_TRIDIAGONAL_SYMMETRIZE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIDIAGONAL_MULT_FUNC
'DESCRIPTION   : Multiply tridiagonal matrix
'LIBRARY       : MATRIX
'GROUP         : TRIDIAGONAL
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRIDIAGONAL_MULT_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

'DATA_MATRIX(NROWS x 3) for DATA_MATRIX matrix B(NROWS x NCOLUMNS) ==> C=A*B

'This function performs the tridiagonal multiplication of two matrices
'MATRIX_RNG: a three diagonals matrix
'VECTOR_RNG: can be a vector (n x 1) or even a rectangular matrix (n x m)
'The result is a vector (n x1) or a matrix (n x m)
'This function accepts both tridiagonal square (n x n) matrices and
'(n x 3) rectangular matrices.

DATA_MATRIX = MATRIX_RNG
DATA_VECTOR = VECTOR_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_VECTOR, 2)

If UBound(DATA_MATRIX, 2) <> 3 Or _
UBound(DATA_MATRIX, 1) <> UBound(DATA_VECTOR, 1) _
Then GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    TEMP_MATRIX(1, j) = DATA_MATRIX(1, 2) * DATA_VECTOR(1, j) + _
        DATA_MATRIX(1, 3) * DATA_VECTOR(2, j)
    For i = 2 To NROWS - 1
       TEMP_MATRIX(i, j) = DATA_MATRIX(i, 1) * DATA_VECTOR(i - 1, j) + _
        DATA_MATRIX(i, 2) * DATA_VECTOR(i, j) + _
        DATA_MATRIX(i, 3) * DATA_VECTOR(i + 1, j)
    Next i
    TEMP_MATRIX(NROWS, j) = DATA_MATRIX(NROWS, 1) * _
        DATA_VECTOR(NROWS - 1, j) + _
        DATA_MATRIX(NROWS, 2) * DATA_VECTOR(NROWS, j)
Next j

MATRIX_TRIDIAGONAL_MULT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_TRIDIAGONAL_MULT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIDIAGONAL_GJ_LINEAR_SYSTEM_FUNC
'DESCRIPTION   : Routine for solving tridiagonal linear system
'LIBRARY       : MATRIX
'GROUP         : TRIDIAGONAL
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRIDIAGONAL_GJ_LINEAR_SYSTEM_FUNC(ByRef MATRIX_RNG As Variant, _
Optional ByRef VECTOR_RNG As Variant, _
Optional ByVal epsilon As Double = 10 ^ -13, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Double
Dim DETERM_VAL As Double
Dim VERSION As Integer

Dim DETERM_FLAG As Boolean
Dim DATA_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

'Remember that for a linear system:
' A X = B
'A = (n x 3) matrix contains the subdiagonal lower,
'    diagonal, subdiagonal upper
'B = (n x m) A the matrix of constant terms.At the end the
'    matrix b contains the solution X

DETERM_FLAG = True
DATA_MATRIX = MATRIX_RNG

If IsArray(VECTOR_RNG) = True Then
    DATA_VECTOR = VECTOR_RNG
    If UBound(DATA_VECTOR, 1) = 1 Then
        DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
    End If
Else
    ReDim DATA_VECTOR(1 To UBound(DATA_MATRIX, 1), 1 To 1)
End If

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_VECTOR, 2)

DATA_MATRIX(1, 1) = 0
DETERM_VAL = 1

For i = 1 To NROWS - 1
    VERSION = 0
    For j = 1 To 3
        If Abs(DATA_MATRIX(i, j)) > epsilon Then VERSION = 1: Exit For
    Next
    If VERSION = 0 Then
        DETERM_VAL = 0: GoTo 1983  'singular matrix
    End If
    
    DATA_MATRIX(i, 1) = DATA_MATRIX(i, 2)
    DATA_MATRIX(i, 2) = DATA_MATRIX(i, 3)
    DATA_MATRIX(i, 3) = 0
    
    If Abs(DATA_MATRIX(i + 1, 1)) > epsilon Then
        If Abs(DATA_MATRIX(i, 1)) < Abs(DATA_MATRIX(i + 1, 1)) Then
            If DETERM_FLAG Then DETERM_VAL = -DETERM_VAL
            DATA_MATRIX = MATRIX_SWAP_ROW_FUNC(DATA_MATRIX, i + 1, i)
            DATA_VECTOR = MATRIX_SWAP_ROW_FUNC(DATA_VECTOR, i + 1, i)
        End If
        
        If DETERM_FLAG Then DETERM_VAL = DETERM_VAL * DATA_MATRIX(i, 1)
        TEMP_VAL = -DATA_MATRIX(i + 1, 1) / DATA_MATRIX(i, 1)
        DATA_MATRIX = MATRIX_LINEAR_ROWS_COMBINATION_FUNC(DATA_MATRIX, i + 1, i, TEMP_VAL)
        DATA_VECTOR = MATRIX_LINEAR_ROWS_COMBINATION_FUNC(DATA_VECTOR, i + 1, i, TEMP_VAL)
    End If
    
    DATA_MATRIX(i + 1, 1) = 0
Next i

'determinant computation

DATA_MATRIX(NROWS, 1) = DATA_MATRIX(NROWS, 2)
DATA_MATRIX(NROWS, 2) = DATA_MATRIX(NROWS, 3)
DATA_MATRIX(NROWS, 3) = 0

If Abs(DATA_MATRIX(NROWS, 1)) <= epsilon Then
    DATA_MATRIX(NROWS, 1) = 0
    DETERM_VAL = 0 '"singular"
End If

If DETERM_FLAG Then
    DETERM_VAL = DETERM_VAL * DATA_MATRIX(NROWS, 1)
    If Abs(DETERM_VAL) <= epsilon Then
        DETERM_VAL = 0
        GoTo 1983 'singular matrix
    End If
End If

For i = 1 To NROWS '1984 last row
    DATA_MATRIX(i, 2) = DATA_MATRIX(i, 2) / DATA_MATRIX(i, 1)
    DATA_MATRIX(i, 3) = DATA_MATRIX(i, 3) / DATA_MATRIX(i, 1)
    
    For j = 1 To NCOLUMNS
        DATA_VECTOR(i, j) = DATA_VECTOR(i, j) / DATA_MATRIX(i, 1)
    Next j
Next i

For i = NROWS - 1 To 1 Step -1 'backsubstitution
    For j = 1 To NCOLUMNS
        DATA_VECTOR(i, j) = DATA_VECTOR(i, j) - _
            DATA_MATRIX(i, 2) * DATA_VECTOR(i + 1, j)
        If i < NROWS - 1 Then DATA_VECTOR(i, j) = _
            DATA_VECTOR(i, j) - DATA_MATRIX(i, 3) * DATA_VECTOR(i + 2, j)
    Next j
Next i

1983:

Select Case OUTPUT
    Case 0
        MATRIX_TRIDIAGONAL_GJ_LINEAR_SYSTEM_FUNC = DATA_VECTOR
    Case 1
        MATRIX_TRIDIAGONAL_GJ_LINEAR_SYSTEM_FUNC = DETERM_VAL
    Case 2
        MATRIX_TRIDIAGONAL_GJ_LINEAR_SYSTEM_FUNC = DATA_MATRIX
    Case Else
        MATRIX_TRIDIAGONAL_GJ_LINEAR_SYSTEM_FUNC = Array(DATA_VECTOR, DETERM_VAL, DATA_MATRIX)
End Select

Exit Function
ERROR_LABEL:
MATRIX_TRIDIAGONAL_GJ_LINEAR_SYSTEM_FUNC = Err.number
End Function

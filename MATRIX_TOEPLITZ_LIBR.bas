Attribute VB_Name = "MATRIX_TOEPLITZ_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TOEPLITZ_GENERATE_FUNC

'DESCRIPTION   : Generate a Toeplitz matrix

'Toeplitz matrices, also called band-diagonal or diagonal-toeplitz matrices,
'have all theirs elements constant along each sub-diagonal

'Example of Toeplitz matrices are:

'Vector has always odd elements M = 2*NROWS-1, where NROWS is
'the dimension of the Toeplitz matrix

' 5   9   0   6   8
' 2   5   9   0   6
'-3   2   5   9   0
'-1  -3   2   5   9
'-2  -1  -3   2   5

'LIBRARY       : MATRIX
'GROUP         : TOEPLITZ
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TOEPLITZ_GENERATE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 1)

Dim i As Long
Dim j As Long

Dim NSIZE As Long

Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
    
NSIZE = (UBound(DATA_VECTOR, 1) - LBound(DATA_VECTOR, 1) + 2)
If NSIZE Mod 2 <> 0 Then: GoTo ERROR_LABEL
NSIZE = NSIZE / 2

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

'----------------------------------------------------------------------------------
Select Case VERSION
'----------------------------------------------------------------------------------
Case 0
'----------------------------------------------------------------------------------
    TEMP_MATRIX(1, 1) = DATA_VECTOR(1, 1)
    For i = 2 To NSIZE
        TEMP_MATRIX(i, 1) = DATA_VECTOR(i, 1)
        TEMP_MATRIX(1, i) = DATA_VECTOR(NSIZE + i - 1, 1)
    Next i
    For i = 2 To NSIZE
        For j = 2 To NSIZE
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i - 1, j - 1)
        Next j
    Next i
'----------------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------------
    TEMP_MATRIX(1, 1) = DATA_VECTOR(NSIZE, 1)
    For i = 2 To NSIZE
        TEMP_MATRIX(i, 1) = DATA_VECTOR(NSIZE - i + 1, 1)
        TEMP_MATRIX(1, i) = DATA_VECTOR(NSIZE + i - 1, 1)
    Next i
    For i = 2 To NSIZE
        For j = 2 To NSIZE
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i - 1, j - 1)
        Next j
    Next i
'----------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------
    
MATRIX_TOEPLITZ_GENERATE_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
MATRIX_TOEPLITZ_GENERATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TOEPLITZ_VALIDATE_FUNC

'DESCRIPTION   : Check Toeplitz matrix; return -1 if is not square
'otherwise return the error distance form the Toeplitz form

'LIBRARY       : MATRIX
'GROUP         : TOEPLITZ
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TOEPLITZ_VALIDATE_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim NSIZE  As Long

Dim MAX_VAL As Double
Dim MIN_VAL As Double
Dim TEMP_VAL As Double

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NSIZE = UBound(DATA_MATRIX, 1)
If NSIZE <> UBound(DATA_MATRIX, 2) Then
    TEMP_VAL = -1
Else
    For i = 1 To NSIZE
        MAX_VAL = Abs(DATA_MATRIX(i, 1))
        MIN_VAL = MAX_VAL
        For j = 2 To NSIZE - i + 1
            If Abs(DATA_MATRIX(i + j - 1, j)) > MAX_VAL Then MAX_VAL = Abs(DATA_MATRIX(i + j - 1, j))
            If Abs(DATA_MATRIX(i + j - 1, j)) < MIN_VAL Then MIN_VAL = Abs(DATA_MATRIX(i + j - 1, j))
        Next j
        TEMP_VAL = TEMP_VAL + MAX_VAL - MIN_VAL
    Next i
    
    For j = 2 To NSIZE
        MAX_VAL = Abs(DATA_MATRIX(1, j))
        MIN_VAL = MAX_VAL
        For i = 2 To NSIZE - j + 1
            If Abs(DATA_MATRIX(i, i + j - 1)) > MAX_VAL Then MAX_VAL = Abs(DATA_MATRIX(i, i + j - 1))
            If Abs(DATA_MATRIX(i, i + j - 1)) < MIN_VAL Then MIN_VAL = Abs(DATA_MATRIX(i, i + j - 1))
        Next i
        TEMP_VAL = TEMP_VAL + MAX_VAL - MIN_VAL
    Next j
End If
MATRIX_TOEPLITZ_VALIDATE_FUNC = TEMP_VAL / (2 * NSIZE) 'average absolute error
  
Exit Function
ERROR_LABEL:
MATRIX_TOEPLITZ_VALIDATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TOEPLITZ_MULT_VECTOR_ELEMENTS_FUNC
'DESCRIPTION   : Perform the multiplication between a Toeplitz matrix
'and a vector
'LIBRARY       : MATRIX
'GROUP         : TOEPLITZ
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TOEPLITZ_MULT_VECTOR_ELEMENTS_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim TEMP_VECTOR As Variant

Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

'Parameter DATA_MATRIX = Toeplitz matrix written in the compact  (2n-1)
'vector form. Parametr VECTOR_RNG is the vector of constant terms

'DATA_MATRIX: vector (2*NROWS-1)
'VECTOR_RNG: vector (NROWS)

DATA_MATRIX = MATRIX_RNG
If UBound(DATA_MATRIX, 1) = 1 Then
    DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
End If

DATA_VECTOR = VECTOR_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

NROWS = UBound(DATA_VECTOR, 1)
NSIZE = UBound(DATA_MATRIX, 1)

If 2 * NROWS - 1 <> NSIZE Or (NSIZE + 1) Mod 2 <> 0 Then GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    For j = 1 To NROWS
        NSIZE = NROWS - i + j
        TEMP_VECTOR(i, 1) = TEMP_VECTOR(i, 1) + DATA_MATRIX(NSIZE, 1) * DATA_VECTOR(j, 1)
    Next j
Next i

MATRIX_TOEPLITZ_MULT_VECTOR_ELEMENTS_FUNC = TEMP_VECTOR
  
Exit Function
ERROR_LABEL:
MATRIX_TOEPLITZ_MULT_VECTOR_ELEMENTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TOEPLITZ_LINEAR_SYSTEM_FUNC

'DESCRIPTION   : Solves Toeplitz linear system by the Levinson’s method
'SysLinTpz saves more than 700% of the elaboration time. So it is
'adapt for large matrices. But of course there is also a drawback:
'not all Toeplitz linear system can be computed. Sometime the algorithm
'fails even if the Toeplitz matrix is not singular. When
'this happens we have to come back to the SysLin function
'It has been demonstrated that if the matrix is diagonal dominant the
'Levinson’s method has successful.

'LIBRARY       : MATRIX
'GROUP         : TOEPLITZ
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_TOEPLITZ_LINEAR_SYSTEM_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant)

Dim h As Long
Dim i As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double

Dim DATA_MATRIX As Variant
Dim DATA_VECTOR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant
Dim RTEMP_VECTOR As Variant
Dim XTEMP_VECTOR As Variant
Dim STEMP_VECTOR As Variant

Dim ERROR_VAL As Double
Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 5 * 10 ^ -14
DATA_MATRIX = MATRIX_RNG

DATA_VECTOR = VECTOR_RNG 'Const Range
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

'-----------------------------------------------------------------------------
If UBound(DATA_MATRIX, 2) = 1 Then
'-----------------------------------------------------------------------------
    k = UBound(DATA_MATRIX, 1)
    If (k + 1) Mod 2 <> 0 Then GoTo ERROR_LABEL
    
    NROWS = (k + 1) / 2
    NSIZE = UBound(DATA_VECTOR, 1)
    
    If NROWS <> NSIZE Then GoTo ERROR_LABEL
    
    ReDim RTEMP_VECTOR(1 To k, 1 To 1)
    ReDim XTEMP_VECTOR(1 To NROWS, 1 To 1)
    ReDim STEMP_VECTOR(1 To NROWS, 1 To 1)
    
    For i = 1 To NROWS
        STEMP_VECTOR(i, 1) = DATA_VECTOR(i, 1)
    Next i
    For i = 1 To k
        RTEMP_VECTOR(i, 1) = DATA_MATRIX(i, 1)
    Next i
'-----------------------------------------------------------------------------
Else
'-----------------------------------------------------------------------------
    NROWS = UBound(DATA_MATRIX, 1)
    NCOLUMNS = UBound(DATA_MATRIX, 2)
    NSIZE = UBound(DATA_VECTOR, 1)
    
    If NROWS <> NSIZE Or NROWS <> NCOLUMNS Then GoTo ERROR_LABEL
    ERROR_VAL = MATRIX_TOEPLITZ_VALIDATE_FUNC(DATA_MATRIX)
    If ERROR_VAL < 0 Or ERROR_VAL > epsilon Then GoTo ERROR_LABEL
    k = 2 * NROWS - 1
    
    ReDim RTEMP_VECTOR(1 To k, 1 To 1)
    ReDim STEMP_VECTOR(1 To NROWS, 1 To 1)
    ReDim XTEMP_VECTOR(1 To NROWS, 1 To 1)
    
    For i = 1 To NROWS
        STEMP_VECTOR(i, 1) = DATA_VECTOR(i, 1)
    Next i
    For i = 1 To NROWS
        RTEMP_VECTOR(i, 1) = DATA_MATRIX(NROWS - i + 1, 1)
    Next i
    For i = 1 To NROWS - 1
        RTEMP_VECTOR(NROWS + i, 1) = DATA_MATRIX(1, i + 1)
    Next i
'-----------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------

ReDim BTEMP_VECTOR(NROWS - 1, 1 To 1)
ReDim CTEMP_VECTOR(NROWS - 1, 1 To 1)
ReDim ATEMP_VECTOR(2 * NROWS - 1, 1 To 1)

'rearrange input vector
For i = 1 To NROWS
  ATEMP_VECTOR(i, 1) = RTEMP_VECTOR(NROWS + i - 1, 1)
Next i

For i = 1 To NROWS - 1
  ATEMP_VECTOR(NROWS + i, 1) = RTEMP_VECTOR(NROWS - i, 1)
Next i

h = 1

If (NROWS < 1) Then: GoTo 1983
ATEMP_VAL = ATEMP_VECTOR(1, 1)

If ATEMP_VAL = 0 Then: GoTo 1983
XTEMP_VECTOR(1, 1) = STEMP_VECTOR(1, 1) / ATEMP_VAL

If (NROWS = 1) Then: GoTo 1983

'  Recurrent process for solving the system with the Toeplitz matrix.

For k = 2 To NROWS ' Compute multiples of the first and last columns
'of the inverse of the principal minor of order k.
    DTEMP_VAL = ATEMP_VECTOR(NROWS + k - 1, 1)
    ETEMP_VAL = ATEMP_VECTOR(k, 1)
  If (k > 2) Then
    BTEMP_VECTOR(k - 1, 1) = BTEMP_VAL
    For i = 1 To k - 2
        DTEMP_VAL = DTEMP_VAL + ATEMP_VECTOR(NROWS + i, 1) * BTEMP_VECTOR(k - i, 1)
        ETEMP_VAL = ETEMP_VAL + ATEMP_VECTOR(i + 1, 1) * CTEMP_VECTOR(i, 1)
    Next i
  End If
  If ATEMP_VAL = 0 Then: GoTo 1983
  BTEMP_VAL = -DTEMP_VAL / ATEMP_VAL
  CTEMP_VAL = -ETEMP_VAL / ATEMP_VAL
  ATEMP_VAL = ATEMP_VAL + DTEMP_VAL * CTEMP_VAL

  If (k > 2) Then
    ETEMP_VAL = CTEMP_VECTOR(1, 1)
    CTEMP_VECTOR(k - 1, 1) = 0
    For i = 2 To k - 1
      DTEMP_VAL = CTEMP_VECTOR(i, 1)
      CTEMP_VECTOR(i, 1) = BTEMP_VECTOR(i, 1) * CTEMP_VAL + ETEMP_VAL
      BTEMP_VECTOR(i, 1) = BTEMP_VECTOR(i, 1) + ETEMP_VAL * BTEMP_VAL
      ETEMP_VAL = DTEMP_VAL
    Next i
  End If
  CTEMP_VECTOR(1, 1) = CTEMP_VAL
  
'  Compute the solution of the system with the principal minor of order k.
  DTEMP_VAL = 0
  For i = 1 To k - 1
      DTEMP_VAL = DTEMP_VAL + ATEMP_VECTOR(NROWS + i, 1) * XTEMP_VECTOR(k - i, 1)
  Next i
  If ATEMP_VAL = 0 Then: GoTo 1983
  ETEMP_VAL = (STEMP_VECTOR(k, 1) - DTEMP_VAL) / ATEMP_VAL
  For i = 1 To k - 1
      XTEMP_VECTOR(i, 1) = XTEMP_VECTOR(i, 1) + CTEMP_VECTOR(i, 1) * ETEMP_VAL
  Next i
  XTEMP_VECTOR(k, 1) = ETEMP_VAL
Next k
h = 0
  
1983:

If h <> 0 Then GoTo ERROR_LABEL
MATRIX_TOEPLITZ_LINEAR_SYSTEM_FUNC = XTEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_TOEPLITZ_LINEAR_SYSTEM_FUNC = Err.number
End Function

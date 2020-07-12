Attribute VB_Name = "MATRIX_CHOLESKY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CHOLESKY_FUNC

'DESCRIPTION   : This function returns the Cholesky decomposition of a
'symmetric matrix: A = L*L^T; where A is a symmetric matrix,
'L is a lower triangular matrix.

'This decomposition works only if A is positive definite. That is:
' v * A * v > 0 Inv(A)*v or, in other words, when the eigenvalues
'of A are all-positive.

'This function always returns  a matrix. Inspecting the diagonal
'elements of the returned matrix we can discover if the matrix is
'positive definite: if the diagonal elements are all positive then
'the matrix A is also positive definite.

'If we see the decomposition of matrix A; the triangular matrix L has
'all diagonal elements positive; then the matrix A is positive definite
'and the eigenvalues are all positive.

'On the contrary, if the decomposition of the matrix B shows a negative
'number; then we can say that B is not positive definite and at least
'one of its eigenvalues is negative.

'This decomposition is useful also for solving the so called generalized
'eigenproblem.

'LIBRARY       : MATRIX
'GROUP         : CHOLESKY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CHOLESKY_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 0.0000000000001

DATA_MATRIX = DATA_RNG
NCOLUMNS = UBound(DATA_MATRIX, 1)
NSIZE = UBound(DATA_MATRIX, 2)

If NCOLUMNS <> NSIZE Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    TEMP_SUM = 0
    For k = 1 To j - 1
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(j, k) ^ 2
    Next k
    TEMP_MATRIX(j, j) = DATA_MATRIX(j, j) - TEMP_SUM
    If TEMP_MATRIX(j, j) <= tolerance Then Exit For
    'the matrix can not be decomp; matrix not positive definite
    TEMP_MATRIX(j, j) = Sqr(TEMP_MATRIX(j, j))
    For i = j + 1 To NCOLUMNS
        TEMP_SUM = 0
        For k = 1 To j - 1
            TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, k) * TEMP_MATRIX(j, k)
        Next k
        TEMP_MATRIX(i, j) = (DATA_MATRIX(i, j) - TEMP_SUM) / TEMP_MATRIX(j, j)
    Next i
Next j

MATRIX_CHOLESKY_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_CHOLESKY_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CHOLESKY_INVERSE_FUNC

'DESCRIPTION   : Given an upper or lower triangle of TEMP_MATRIX symmetric positive
'definite matrix, the algorithm generates matrix TEMP_MATRIX^-1 and
'saves the upper or lower triangle depending on the input.
'
'Input parameters:
'    TEMP_MATRIX -   matrix to be inverted (upper or lower triangle).
'                    Array with elements [1..NSIZE].
'    NSIZE       -   size of matrix TEMP_MATRIX.
'    UPPER_FLAG  -   storage format.
'                If UPPER_FLAG = True, then the upper triangle of matrix
'                   TEMP_MATRIX is given, otherwise the lower triangle is given.
'Output parameters:
'    TEMP_MATRIX       -   inverse of matrix TEMP_MATRIX.
'                Array with elements [1..NSIZE, 1..NSIZE].
'                If UPPER_FLAG = True, then the upper triangle of
'                matrix TEMP_MATRIX^-1 is used, and the elements below
'                the main diagonal are not used nor changed. The
'                same applies if UPPER_FLAG = False.
'RESULT_FLAG:
'    True, if the matrix is positive definite.
'    False, if the matrix is not positive definite (and it could not be
'    inverted by this algorithm).

'LIBRARY       : MATRIX
'GROUP         : CHOLESKY
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009

'************************************************************************************
'************************************************************************************


Function MATRIX_CHOLESKY_INVERSE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NSIZE As Long = 0, _
Optional ByVal UPPER_FLAG As Boolean = True)
    
Dim ii As Long
Dim jj As Long

Dim RESULT_FLAG As Boolean
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL
If NSIZE = 0 Then: NSIZE = UBound(DATA_MATRIX, 1)

DATA_MATRIX = _
   MATRIX_CHOLESKY_DECOMPOSITION_FUNC(DATA_MATRIX, NSIZE, UPPER_FLAG, RESULT_FLAG)
    If RESULT_FLAG = False Then: GoTo ERROR_LABEL
DATA_MATRIX = _
    MATRIX_PD_SYMMETRIC_INVERSE_FUNC(DATA_MATRIX, NSIZE, UPPER_FLAG, RESULT_FLAG)
    If RESULT_FLAG = False Then: GoTo ERROR_LABEL

If UPPER_FLAG = True Then 'routine has inverted values in lower half of diagonal
'must adjust upper half
    For ii = 1 To NSIZE
        For jj = 1 To NSIZE
            If ii > jj Then: DATA_MATRIX(ii, jj) = DATA_MATRIX(jj, ii)
        Next jj
    Next ii
End If

MATRIX_CHOLESKY_INVERSE_FUNC = DATA_MATRIX
    
Exit Function
ERROR_LABEL:
MATRIX_CHOLESKY_INVERSE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CHOLESKY_DECOMPOSITION_FUNC

'DESCRIPTION   : This function returns the Cholesky decomposition of a
'symmetric matrix. This decomposition is useful also for solving the
'so called generalized eigenproblem.

'    NSIZE        -   size of matrix TEMP_MATRIX.
'    UPPER_FLAG   –   storage format.
'                 If UPPER_FLAG = True, then matrix TEMP_MATRIX is
'                 given as TEMP_MATRIX = U'*U (matrix contains upper triangle).
'                 Similarly, if UPPER_FLAG = False, then TEMP_MATRIX = L*L'.
'RESULT_FLAG:
'    True, if the inversion succeeded.
'    False, if matrix TEMP_MATRIX contains zero elements on its main diagonal.
'    Matrix TEMP_MATRIX could not be inverted.

'LIBRARY       : MATRIX
'GROUP         : CHOLESKY
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Private Function MATRIX_CHOLESKY_DECOMPOSITION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NSIZE As Long = 0, _
Optional ByVal UPPER_FLAG As Boolean = True, _
Optional ByRef RESULT_FLAG As Boolean)

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim DATA_MATRIX As Variant
    
On Error GoTo ERROR_LABEL
    
DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL
If NSIZE = 0 Then: NSIZE = UBound(DATA_MATRIX, 1)

RESULT_FLAG = True
If NSIZE = 0 Then
   MATRIX_CHOLESKY_DECOMPOSITION_FUNC = DATA_MATRIX
    Exit Function
End If
'----------------------------------------------------------------------------------
If UPPER_FLAG Then
'----------------------------------------------------------------------------------
    ' Compute the Cholesky factorization DATA_MATRIX = U'*U.
    For jj = 1 To NSIZE Step 1
        
        '
        ' Compute U(jj,jj) and test for non-positive-definiteness.
        '
        i = jj - 1
        BTEMP_VAL = 0
        For kk = 1 To i Step 1
            BTEMP_VAL = BTEMP_VAL + _
                DATA_MATRIX(kk, jj) * DATA_MATRIX(kk, jj)
        Next kk
        ATEMP_VAL = DATA_MATRIX(jj, jj) - BTEMP_VAL
        If ATEMP_VAL <= 0 Then
            RESULT_FLAG = False
           MATRIX_CHOLESKY_DECOMPOSITION_FUNC = DATA_MATRIX
            Exit Function
        End If
        ATEMP_VAL = Sqr(ATEMP_VAL)
        DATA_MATRIX(jj, jj) = ATEMP_VAL
        
        '
        ' Compute elements J+1:NSIZE of row jj.
        '
        If jj < NSIZE Then
            For ii = jj + 1 To NSIZE Step 1
                i = jj - 1
                BTEMP_VAL = 0
                For kk = 1 To i Step 1
                    BTEMP_VAL = BTEMP_VAL + _
                        DATA_MATRIX(kk, ii) * DATA_MATRIX(kk, jj)
                Next kk
                DATA_MATRIX(jj, ii) = DATA_MATRIX(jj, ii) - BTEMP_VAL
            Next ii
            BTEMP_VAL = 1 / ATEMP_VAL
            j = jj + 1
            For kk = j To NSIZE Step 1
                DATA_MATRIX(jj, kk) = BTEMP_VAL * DATA_MATRIX(jj, kk)
            Next kk
        End If
    Next jj

'----------------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------------
    ' Compute the Cholesky factorization DATA_MATRIX = L*L'.
    For jj = 1 To NSIZE Step 1
        
        '
        ' Compute L(jj,jj) and test for non-positive-definiteness.
        '
        i = jj - 1
        BTEMP_VAL = 0
        For kk = 1 To i Step 1
            BTEMP_VAL = BTEMP_VAL + _
                DATA_MATRIX(jj, kk) * DATA_MATRIX(jj, kk)
        Next kk
        ATEMP_VAL = DATA_MATRIX(jj, jj) - BTEMP_VAL
        If ATEMP_VAL <= 0 Then
            RESULT_FLAG = False
           MATRIX_CHOLESKY_DECOMPOSITION_FUNC = DATA_MATRIX
            Exit Function
        End If
        ATEMP_VAL = Sqr(ATEMP_VAL)
        DATA_MATRIX(jj, jj) = ATEMP_VAL
        
        '
        ' Compute elements J+1:NSIZE of column jj.
        '
        If jj < NSIZE Then
            For ii = jj + 1 To NSIZE Step 1
                i = jj - 1
                BTEMP_VAL = 0
                For kk = 1 To i Step 1
                    BTEMP_VAL = BTEMP_VAL + _
                        DATA_MATRIX(ii, kk) * DATA_MATRIX(jj, kk)
                Next kk
                DATA_MATRIX(ii, jj) = DATA_MATRIX(ii, jj) - BTEMP_VAL
            Next ii
            BTEMP_VAL = 1 / ATEMP_VAL
            j = jj + 1
            For kk = j To NSIZE Step 1
                DATA_MATRIX(kk, jj) = BTEMP_VAL * DATA_MATRIX(kk, jj)
            Next kk
        End If
    Next jj
'----------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------

MATRIX_CHOLESKY_DECOMPOSITION_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CHOLESKY_DECOMPOSITION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_PD_SYMMETRIC_INVERSE_FUNC

'DESCRIPTION   : Inversion of TEMP_MATRIX symmetric positive definite
'matrix which is given by Cholesky decomposition.
'
'Input parameters:
'    TEMP_MATRIX  -   Cholesky decomposition of the matrix to be inverted:
'                 A=U’*U or TEMP_MATRIX = L*L'.
'                 Output of MATRIX_CHOLESKY_DECOMPOSITION_FUNC subroutine.
'                 Array with elements [1..NSIZE, 1..NSIZE].
'    NSIZE        -   size of matrix TEMP_MATRIX.
'    UPPER_FLAG   –   storage format.
'                 If UPPER_FLAG = True, then matrix TEMP_MATRIX is
'                 given as TEMP_MATRIX = U'*U (matrix contains upper triangle).
'                 Similarly, if UPPER_FLAG = False, then TEMP_MATRIX = L*L'.
'
'Output parameters:
'    TEMP_MATRIX  -   upper or lower triangle of symmetric matrix
'                 TEMP_MATRIX^-1, depending on the value of UPPER_FLAG.
'
'RESULT_FLAG:
'    True, if the inversion succeeded.
'    False, if matrix TEMP_MATRIX contains zero elements on its main diagonal.
'    Matrix TEMP_MATRIX could not be inverted.

'LIBRARY       : MATRIX
'GROUP         : CHOLESKY
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Private Function MATRIX_PD_SYMMETRIC_INVERSE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NSIZE As Long = 0, _
Optional ByVal UPPER_FLAG As Boolean = True, _
Optional ByRef RESULT_FLAG As Boolean)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim ATEMP_ARR() As Double
Dim BTEMP_ARR() As Double

Dim DATA_MATRIX As Variant

RESULT_FLAG = True
DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL
If NSIZE = 0 Then: NSIZE = UBound(DATA_MATRIX, 1)

ReDim ATEMP_ARR(1 To NSIZE)
ReDim BTEMP_ARR(1 To NSIZE)
'--------------------------------------------------------------------------------
If UPPER_FLAG Then
'--------------------------------------------------------------------------------
    ' Compute inverse of upper triangular matrix.
    For jj = 1 To NSIZE Step 1
        If DATA_MATRIX(jj, jj) = 0 Then
            RESULT_FLAG = False
            MATRIX_PD_SYMMETRIC_INVERSE_FUNC = DATA_MATRIX
            Exit Function
        End If
        i = jj - 1
        DATA_MATRIX(jj, jj) = 1 / DATA_MATRIX(jj, jj)
        BTEMP_VAL = -DATA_MATRIX(jj, jj)
        
        '
        ' Compute elements 1:jj-1 of jj-th column.
        '
        For kk = 1 To i Step 1
            ATEMP_ARR(kk) = DATA_MATRIX(kk, jj)
        Next kk
        For ii = 1 To jj - 1 Step 1
            ATEMP_VAL = 0
            For kk = ii To i Step 1
                ATEMP_VAL = ATEMP_VAL + _
                    DATA_MATRIX(ii, kk) * DATA_MATRIX(kk, jj)
            Next kk
            DATA_MATRIX(ii, jj) = ATEMP_VAL
        Next ii
        For kk = 1 To i Step 1
            DATA_MATRIX(kk, jj) = BTEMP_VAL * DATA_MATRIX(kk, jj)
        Next kk
    Next jj
    
    '
    ' InvA = InvU * InvU'
    '
    For ii = 1 To NSIZE Step 1
        CTEMP_VAL = DATA_MATRIX(ii, ii)
        If ii < NSIZE Then
            ATEMP_VAL = 0
            For kk = ii To NSIZE Step 1
                ATEMP_VAL = ATEMP_VAL + _
                    DATA_MATRIX(ii, kk) * DATA_MATRIX(ii, kk)
            Next kk
            DATA_MATRIX(ii, ii) = ATEMP_VAL
            k = ii + 1
            For h = 1 To ii - 1 Step 1
                ATEMP_VAL = 0
                For kk = k To NSIZE Step 1
                    ATEMP_VAL = ATEMP_VAL + _
                        DATA_MATRIX(h, kk) * DATA_MATRIX(ii, kk)
                Next kk
                DATA_MATRIX(h, ii) = _
                    DATA_MATRIX(h, ii) * CTEMP_VAL + ATEMP_VAL
            Next h
        Else
            For kk = 1 To ii Step 1
                DATA_MATRIX(kk, ii) = CTEMP_VAL * DATA_MATRIX(kk, ii)
            Next kk
        End If
    Next ii
'--------------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------------
    ' Compute inverse of lower triangular matrix.
    For jj = NSIZE To 1 Step -1
        If DATA_MATRIX(jj, jj) = 0 Then
            RESULT_FLAG = False
            MATRIX_PD_SYMMETRIC_INVERSE_FUNC = DATA_MATRIX
            Exit Function
        End If
        DATA_MATRIX(jj, jj) = 1 / DATA_MATRIX(jj, jj)
        BTEMP_VAL = -DATA_MATRIX(jj, jj)
        If jj < NSIZE Then
            
            '
            ' Compute elements j+1:NSIZE of jj-th column.
            '
            l = NSIZE - jj
            j = jj + 1
            For kk = j To NSIZE Step 1
                ATEMP_ARR(kk) = DATA_MATRIX(kk, jj)
            Next kk
            For ii = jj + 1 To NSIZE Step 1
                ATEMP_VAL = 0
                For kk = j To ii Step 1
                    ATEMP_VAL = ATEMP_VAL + _
                        DATA_MATRIX(ii, kk) * ATEMP_ARR(kk)
                Next kk
                DATA_MATRIX(ii, jj) = ATEMP_VAL
            Next ii
            For kk = j To NSIZE Step 1
                DATA_MATRIX(kk, jj) = BTEMP_VAL * DATA_MATRIX(kk, jj)
            Next kk
        End If
    Next jj

    ' InvA = InvL' * InvL

    For ii = 1 To NSIZE Step 1
        CTEMP_VAL = DATA_MATRIX(ii, ii)
        If ii < NSIZE Then
            ATEMP_VAL = 0
            For kk = ii To NSIZE Step 1
                ATEMP_VAL = ATEMP_VAL + _
                    DATA_MATRIX(kk, ii) * DATA_MATRIX(kk, ii)
            Next kk
            DATA_MATRIX(ii, ii) = ATEMP_VAL
            k = ii + 1
            For h = 1 To ii - 1 Step 1
                ATEMP_VAL = 0
                For kk = k To NSIZE Step 1
                    ATEMP_VAL = ATEMP_VAL + _
                        DATA_MATRIX(kk, h) * DATA_MATRIX(kk, ii)
                Next kk
                DATA_MATRIX(ii, h) = _
                    CTEMP_VAL * DATA_MATRIX(ii, h) + ATEMP_VAL
            Next h
        Else
            For kk = 1 To ii Step 1
                DATA_MATRIX(ii, kk) = CTEMP_VAL * DATA_MATRIX(ii, kk)
            Next kk
        End If
    Next ii
'--------------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------------

MATRIX_PD_SYMMETRIC_INVERSE_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_PD_SYMMETRIC_INVERSE_FUNC = Err.number
End Function

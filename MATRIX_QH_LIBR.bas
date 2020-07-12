Attribute VB_Name = "MATRIX_QH_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_QH_DECOMPOSITION_FUNC
'DESCRIPTION   : This function performs the QH decomposition of the square matrix

'A with the vector b
'A = Q*H*Q^t

'where:
'A is a square (n x n) matrix
'b is a (n x 1) vector
'Q is an orthogonal matrix
'H is an Hessenberg matrix.
'If A is symmetric, then H is a tridiagonal matrix

'The functions returns a matrix (n x 2n), where the first (n x n)
'block contains Q and the second (n x n) block  contains H.

'returns the Q H factorization matrices where
'Q = orthonormal and H = hessemberg being [Q^t]*[A]*[Q]=[H]
'it uses the Arnoldi's algorithm.

'For symmetric matrix uses the faster Lanczos' algorithm

'LIBRARY       : MATRIX
'GROUP         : QH
'ID            : 001



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function MATRIX_QH_DECOMPOSITION_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant, _
Optional ByRef nLOOPS As Long = 0, _
Optional ByVal epsilon As Double = 5 * 10 ^ -16)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

Dim ALPHA_VAL As Double
Dim BETA_VAL As Double
Dim NORM_VAL As Double

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant
Dim DTEMP_MATRIX As Variant

Dim SYMM_FLAG As Boolean

On Error GoTo ERROR_LABEL

ATEMP_MATRIX = MATRIX_RNG
ATEMP_VECTOR = VECTOR_RNG

If UBound(ATEMP_MATRIX, 1) <> UBound(ATEMP_MATRIX, 2) Then GoTo ERROR_LABEL
If UBound(ATEMP_VECTOR, 1) <> UBound(ATEMP_MATRIX, 1) Then GoTo ERROR_LABEL

NROWS = UBound(ATEMP_MATRIX, 1)
If nLOOPS = 0 Then nLOOPS = NROWS

ReDim BTEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    BTEMP_VECTOR(i, 1) = ATEMP_VECTOR(i, 1)
Next i

SYMM_FLAG = False
    For i = 1 To UBound(ATEMP_MATRIX, 1)
        For j = 1 To UBound(ATEMP_MATRIX, 2)
            If ATEMP_MATRIX(i, j) <> ATEMP_MATRIX(j, i) Then: GoTo 1983
        Next j
    Next i
SYMM_FLAG = True
1983:

'------------------------------------------------------------------------------
If SYMM_FLAG Then
'------------------------------------------------------------------------------
'Lanczos factorization of the linear system Ax = b
'returns the Lanczos's factorization matrices Q,H
'where [Q^t]*[A]*[Q]=[H]


    ReDim BTEMP_MATRIX(1 To NROWS, 1 To NROWS)
    ReDim CTEMP_MATRIX(1 To NROWS, 1 To NROWS)
    ReDim CTEMP_VECTOR(1 To NROWS, 1 To 1)

    'compute the norm of the vector
    NORM_VAL = VECTOR_EUCLIDEAN_NORM_FUNC(BTEMP_VECTOR)
    For i = 1 To NROWS
        BTEMP_MATRIX(i, 1) = BTEMP_VECTOR(i, 1) / NORM_VAL
    Next i
    h = 1 'nLOOPS COUNTER
    Do
        For i = 1 To NROWS
            CTEMP_VECTOR(i, 1) = 0
            For j = 1 To NROWS
                CTEMP_VECTOR(i, 1) = CTEMP_VECTOR(i, 1) + _
                    ATEMP_MATRIX(i, j) * BTEMP_MATRIX(j, h)
            Next j
        Next i
    
        ALPHA_VAL = 0
        For i = 1 To NROWS
            ALPHA_VAL = ALPHA_VAL + CTEMP_VECTOR(i, 1) * BTEMP_MATRIX(i, h)
        Next i
    
        If h = 1 Then
            For i = 1 To NROWS
                CTEMP_VECTOR(i, 1) = CTEMP_VECTOR(i, 1) - ALPHA_VAL * _
                        BTEMP_MATRIX(i, h)
            Next i
        Else
            For i = 1 To NROWS
                CTEMP_VECTOR(i, 1) = CTEMP_VECTOR(i, 1) - ALPHA_VAL * _
                    BTEMP_MATRIX(i, h) - BETA_VAL * BTEMP_MATRIX(i, h - 1)
            Next i
        End If
        
        BETA_VAL = VECTOR_EUCLIDEAN_NORM_FUNC(CTEMP_VECTOR)
        CTEMP_MATRIX(h, h) = ALPHA_VAL
        
        If h > 1 Then CTEMP_MATRIX(h - 1, h) = CTEMP_MATRIX(h, h - 1)
        If h = NROWS Or h = nLOOPS Then Exit Do
        
        CTEMP_MATRIX(h + 1, h) = BETA_VAL
        For i = 1 To NROWS
            BTEMP_MATRIX(i, h + 1) = CTEMP_VECTOR(i, 1) / BETA_VAL
        Next i
        h = h + 1
    Loop Until h = NROWS + 1 Or Abs(CTEMP_MATRIX(h, h - 1)) < epsilon
'------------------------------------------------------------------------------
Else
'------------------------------------------------------------------------------

'Arnoldi factorization of the linear system Ax = b
'returns the Arnoldi's factorization matrices Q,H where [Q^t]*[A]*[Q]=[H]

    ReDim BTEMP_MATRIX(1 To NROWS, 1 To NROWS)
    ReDim CTEMP_MATRIX(1 To NROWS, 1 To NROWS)
    ReDim CTEMP_VECTOR(1 To NROWS, 1 To 1)
    'compute the norm of the vector
    NORM_VAL = VECTOR_EUCLIDEAN_NORM_FUNC(BTEMP_VECTOR)
    For i = 1 To NROWS
        BTEMP_MATRIX(i, 1) = BTEMP_VECTOR(i, 1) / NORM_VAL
    Next i
    
    h = 1 'nLOOPS COUNTER
    Do
        For i = 1 To NROWS
            CTEMP_VECTOR(i, 1) = 0
            For j = 1 To NROWS
                CTEMP_VECTOR(i, 1) = CTEMP_VECTOR(i, 1) + _
                    ATEMP_MATRIX(i, j) * BTEMP_MATRIX(j, h)
            Next j
        Next i
        For i = 1 To h
            CTEMP_MATRIX(i, h) = 0
            For k = 1 To NROWS
                CTEMP_MATRIX(i, h) = CTEMP_MATRIX(i, h) + _
                    BTEMP_MATRIX(k, i) * CTEMP_VECTOR(k, 1)
            Next k
            For k = 1 To NROWS
                CTEMP_VECTOR(k, 1) = CTEMP_VECTOR(k, 1) - CTEMP_MATRIX(i, h) * _
                    BTEMP_MATRIX(k, i)
            Next k
        Next i
        If h = NROWS Or h = nLOOPS Then Exit Do
        CTEMP_MATRIX(h + 1, h) = VECTOR_EUCLIDEAN_NORM_FUNC(CTEMP_VECTOR)
        For k = 1 To NROWS
            BTEMP_MATRIX(k, h + 1) = CTEMP_VECTOR(k, 1) / CTEMP_MATRIX(h + 1, h)
        Next k
        h = h + 1
    Loop Until h = NROWS + 1 Or Abs(CTEMP_MATRIX(h, h - 1)) < epsilon

'------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------

'arrange the output matrix
ReDim DTEMP_MATRIX(1 To NROWS, 1 To 2 * NROWS)
For i = 1 To NROWS
    For j = 1 To NROWS
        DTEMP_MATRIX(i, j) = BTEMP_MATRIX(i, j)
        DTEMP_MATRIX(i, j + NROWS) = CTEMP_MATRIX(i, j)
    Next j
Next i

MATRIX_QH_DECOMPOSITION_FUNC = DTEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_QH_DECOMPOSITION_FUNC = Err.number
End Function

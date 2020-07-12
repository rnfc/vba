Attribute VB_Name = "MATRIX_QR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.

                            
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_QR_DECOMPOSITION_FUNC
'DESCRIPTION   : 'This function(*) performs the QR decomposition

'A = Q * R
'where:
'A is a rectangular (m x n) matrix, with m >= n.
'Q is an orthogonal (m x n) matrix
'R is an upper triangular (n x n) matrix.

'This function returns a matrix (m x (n + n)), where
'the first (m x n) block is Q and the first n rows of
'the second (m x n) block is R. The last m - n rows of
'the second block are all zero.

'The QR decomposition forms the basis of an efficient method for
'calculating the eigenvalues.

'LIBRARY       : MATRIX
'GROUP         : QR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************


Function MATRIX_QR_DECOMPOSITION_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim TEMP_VALUE As Double

Dim TEMP_VECTOR As Variant
Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
If NCOLUMNS = NROWS Then NSIZE = NCOLUMNS - 1 Else NSIZE = NCOLUMNS

For k = 1 To NSIZE 'compute the modulus of vector k
    ATEMP_SUM = 0
    For i = 1 To NROWS
        ATEMP_SUM = ATEMP_SUM + DATA_MATRIX(i, k) ^ 2
    Next i
    ATEMP_SUM = Sqr(ATEMP_SUM)
    'normalize vector and load TEMP_VECTOR
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = DATA_MATRIX(i, k) / ATEMP_SUM
    Next i
    'computes D= (TEMP_VECTOR(k,1)^2+v(k+1)^2+...TEMP_VECTOR(NCOLUMNS,1)^2)^(1/2)
    BTEMP_SUM = 0
    For i = k To NROWS
        BTEMP_SUM = BTEMP_SUM + TEMP_VECTOR(i, 1) ^ 2
    Next i
    BTEMP_SUM = Sqr(BTEMP_SUM)
    If TEMP_VECTOR(k, 1) > 0 Then BTEMP_SUM = -BTEMP_SUM
    'compute TEMP_VECTOR
    For i = 1 To NROWS
        If i < k Then
            TEMP_VECTOR(i, 1) = 0
        ElseIf i = k Then
            TEMP_VECTOR(k, 1) = Sqr((1 - TEMP_VECTOR(k, 1) / BTEMP_SUM) / 2)
            TEMP_VALUE = -BTEMP_SUM * TEMP_VECTOR(k, 1)
        Else
            TEMP_VECTOR(i, 1) = TEMP_VECTOR(i, 1) / TEMP_VALUE / 2
        End If
    Next i
    ATEMP_MATRIX = MATRIX_HOUSEHOLDER_FUNC(TEMP_VECTOR)
    DATA_MATRIX = MMULT_FUNC(ATEMP_MATRIX, DATA_MATRIX, 70)
    If k = 1 Then
        BTEMP_MATRIX = ATEMP_MATRIX
    Else
        BTEMP_MATRIX = MMULT_FUNC(ATEMP_MATRIX, BTEMP_MATRIX, 70)
    End If
Next k
BTEMP_MATRIX = MATRIX_TRANSPOSE_FUNC(BTEMP_MATRIX) 'make positive the diagonal elements of

For i = 1 To NCOLUMNS
    If DATA_MATRIX(i, i) < 0 Then
        DATA_MATRIX = MATRIX_CHANGE_SIGN_ROW_FUNC(DATA_MATRIX, i)
        BTEMP_MATRIX = MATRIX_CHANGE_SIGN_COLUMN_FUNC(BTEMP_MATRIX, i)
    End If
Next i

ReDim ATEMP_MATRIX(1 To NROWS, 1 To 2 * NCOLUMNS) 'load matrix out QR
For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        ATEMP_MATRIX(i, j) = BTEMP_MATRIX(i, j) '-BTEMP_MATRIX(i, j)
        If i <= NCOLUMNS Then
            ATEMP_MATRIX(i, j + NCOLUMNS) = DATA_MATRIX(i, j) '-DATA_MATRIX(i, j)
        End If
    Next j
Next i

MATRIX_QR_DECOMPOSITION_FUNC = ATEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_QR_DECOMPOSITION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_QR_DECOMPOSITION_FUNC
'DESCRIPTION   : This function print the QR decomposition by segments (M=Q or M=R)
'LIBRARY       : MATRIX
'GROUP         : QR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function PRINT_MATRIX_QR_DECOMPOSITION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
SCOLUMN = LBound(DATA_MATRIX, 2)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

DATA_MATRIX = MATRIX_QR_DECOMPOSITION_FUNC(DATA_MATRIX)

'-----------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------
Case 0 'matrix Q
'-----------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            TEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
        Next j
    Next i
'-----------------------------------------------------------------------
Case Else ''matrix R
'-----------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
    For i = 1 To NCOLUMNS
        For j = 1 To NCOLUMNS
            TEMP_MATRIX(i, j) = DATA_MATRIX(i, NCOLUMNS + j)
        Next j
    Next i
'-----------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------

PRINT_MATRIX_QR_DECOMPOSITION_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
PRINT_MATRIX_QR_DECOMPOSITION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ITERATE_MATRIX_QR_DECOMPOSITION_FUNC

'DESCRIPTION   : This function performs the diagonalization of a symmetric
'matrix by the QR iterative process. The heart of this method is the QR iterated
'decomposition. If the matrix is not symmetric the process gives a
'triangular matrix where the diagonal elements are still the eigenvalues.
'performs the iterative QR method for triangolarization/diagonalization

'LIBRARY       : MATRIX
'GROUP         : QR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ITERATE_MATRIX_QR_DECOMPOSITION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal nLOOPS As Long = 100)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
If NROWS <> NCOLUMNS Then: GoTo ERROR_LABEL

ReDim BTEMP_MATRIX(1 To NROWS, 1 To NROWS)
ReDim CTEMP_MATRIX(1 To NROWS, 1 To NROWS)

k = 1
Do Until k > nLOOPS
    ATEMP_MATRIX = MATRIX_QR_DECOMPOSITION_FUNC(DATA_MATRIX)
    GoSub 1983
    DATA_MATRIX = MMULT_FUNC(CTEMP_MATRIX, BTEMP_MATRIX, 70)
    k = k + 1
Loop

ITERATE_MATRIX_QR_DECOMPOSITION_FUNC = DATA_MATRIX

Exit Function
'--------------------------------------------------------------------------
1983:
'--------------------------------------------------------------------------
    For i = 1 To NROWS
        For j = 1 To NROWS
            BTEMP_MATRIX(i, j) = ATEMP_MATRIX(i, j)
            CTEMP_MATRIX(i, j) = ATEMP_MATRIX(i, j + NROWS)
        Next j
    Next i
'--------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------
Exit Function
ERROR_LABEL:
ITERATE_MATRIX_QR_DECOMPOSITION_FUNC = Err.number
End Function

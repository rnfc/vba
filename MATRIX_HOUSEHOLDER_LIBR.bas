Attribute VB_Name = "MATRIX_HOUSEHOLDER_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_HOUSEHOLDER_FUNC
'DESCRIPTION   : Build Houseolder matrix H
'LIBRARY       : MATRIX
'GROUP         : HOUSEHOLDER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function MATRIX_HOUSEHOLDER_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim TEMP_MOD As Double 'Modulus

Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim epsilon As Double

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

epsilon = 2 * 10 ^ -15

NROWS = UBound(DATA_VECTOR, 1)
TEMP_MOD = VECTOR_EUCLIDEAN_NORM_FUNC(DATA_VECTOR) 'normalize DATA_VECTOR

For i = 1 To NROWS
    DATA_VECTOR(i, 1) = DATA_VECTOR(i, 1) / TEMP_MOD
Next i

ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)

For i = 1 To NROWS
    For j = 1 To NROWS
        TEMP_MATRIX(i, j) = -2 * DATA_VECTOR(i, 1) * DATA_VECTOR(j, 1)
        If Abs(TEMP_MATRIX(i, j)) < epsilon Then _
            TEMP_MATRIX(i, j) = 0
        If i = j Then TEMP_MATRIX(i, j) = 1 + TEMP_MATRIX(i, j)
    Next j
Next i

MATRIX_HOUSEHOLDER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_HOUSEHOLDER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_HOUSEHOLDER2_FUNC

'DESCRIPTION   : 'Returns the Householder matrix I-2*W*Wt/|W|^2
'Note that, if an element vi of the vector is zero, then the
'Houseolder matrix has always zero the i-column and the i-row,
'except the diagonal element aii = 1. Note also that the determinant is always -1
'These matrices are used in several important algorithms as, for example,
'the QR decomposition. One, unusual, application of the Householder matrix
'is at the generation of a random symmetric matrix with given eigenvalues

'LIBRARY       : MATRIX
'GROUP         : HOUSEHOLDER
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function MATRIX_HOUSEHOLDER2_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_SCALAR As Double

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim DATA_MATRIX As Variant

Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 2 * 10 ^ -15
DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)

NCOLUMNS = UBound(DATA_MATRIX, 2)
TEMP_SUM = 0
For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j) ^ 2
    Next j
Next i

TEMP_SCALAR = TEMP_SUM / 2
ATEMP_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
BTEMP_VECTOR = MMULT_FUNC(DATA_MATRIX, ATEMP_VECTOR, 70)

For i = 1 To NROWS
    For j = 1 To NROWS
        BTEMP_VECTOR(i, j) = -BTEMP_VECTOR(i, j) / TEMP_SCALAR
        If Abs(BTEMP_VECTOR(i, j)) < epsilon Then _
            BTEMP_VECTOR(i, j) = 0
        If i = j Then BTEMP_VECTOR(i, j) = 1 + BTEMP_VECTOR(i, j)
    Next j
Next i

MATRIX_HOUSEHOLDER2_FUNC = BTEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_HOUSEHOLDER2_FUNC = Err.number
End Function

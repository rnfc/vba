Attribute VB_Name = "NUMBER_COMPLEX_POLYNOMIAL_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_COMPANION_POLYNOMIAL_FUNC
'DESCRIPTION   : Returns the complex companion matrix of a polynomial
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : POLYNOMIAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_COMPANION_POLYNOMIAL_FUNC(ByRef COEF_RNG As Variant)

Dim i As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim ACOMPLEX_OBJ As Cplx
Dim BCOMPLEX_OBJ As Cplx
Dim CCOMPLEX_OBJ As Cplx

On Error GoTo ERROR_LABEL

DATA_VECTOR = COEF_RNG

NROWS = UBound(DATA_VECTOR, 1)
NSIZE = NROWS - 1

ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2 * NSIZE)

If DATA_VECTOR(NROWS, 1) <> 1 And DATA_VECTOR(NROWS, 2) = 0 Then
    'real normalization
    For i = 1 To NROWS
       DATA_VECTOR(i, 1) = DATA_VECTOR(i, 1) / DATA_VECTOR(NROWS, 1)
       DATA_VECTOR(i, 2) = DATA_VECTOR(i, 2) / DATA_VECTOR(NROWS, 1)
    Next i
ElseIf DATA_VECTOR(NROWS, 2) <> 0 Then 'complex normalization
    ACOMPLEX_OBJ.reel = DATA_VECTOR(NROWS, 1)
    ACOMPLEX_OBJ.imag = DATA_VECTOR(NROWS, 2)
    
    For i = 1 To NROWS
       BCOMPLEX_OBJ.reel = DATA_VECTOR(i, 1)
       BCOMPLEX_OBJ.imag = DATA_VECTOR(i, 2)
       
       CCOMPLEX_OBJ = COMPLEX_QUOTIENT_OBJ_FUNC(BCOMPLEX_OBJ, ACOMPLEX_OBJ)
       DATA_VECTOR(i, 1) = CCOMPLEX_OBJ.reel
       DATA_VECTOR(i, 2) = CCOMPLEX_OBJ.imag
    Next i
End If
'build subdiagonal (lower)
For i = 1 To NSIZE - 1
    TEMP_MATRIX(i + 1, i) = 1
Next i
'insert coefficients into last column
For i = 1 To NSIZE
    TEMP_MATRIX(i, NSIZE) = -DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 2 * NSIZE) = -DATA_VECTOR(i, 2)
Next i

COMPLEX_MATRIX_COMPANION_POLYNOMIAL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_COMPANION_POLYNOMIAL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_CHARACTERISTIC_POLYNOMIAL_FUNC

'DESCRIPTION   : This function returns the complex coefficients of the
'characteristic polynomial of a complex matrix. If the matrix has dimension (n x n),
'then the polynomial has n degree and the coefficients are n+1
'As know, the roots of the characteristic polynomial are the eigenvalues
'of the matrix and vice versa. This function uses the Newton-Girard
'formulas to find all the coefficients. This function supports 3 different
'matrix formats: 1 = split, 2 = interlaced, 3 = string
'Optional parameter CPLX_FORMAT sets the complex input/output
'format (default = 1)

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : POLYNOMIAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_CHARACTERISTIC_POLYNOMIAL_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim COEF_VECTOR As Variant

Dim DATA_MATRIX As Variant
Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If CPLX_FORMAT = 2 Then DATA_MATRIX = _
            COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then DATA_MATRIX = _
            COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

If CPLX_FORMAT = 1 And (UBound(DATA_MATRIX, 1) = _
    UBound(DATA_MATRIX, 2)) Then
    ReDim Preserve DATA_MATRIX(1 To UBound(DATA_MATRIX, 1), _
    1 To 2 * UBound(DATA_MATRIX, 1))
End If
If 2 * UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then GoTo ERROR_LABEL

NROWS = UBound(DATA_MATRIX)

ReDim ATEMP_VECTOR(1 To NROWS, 1 To 2)
ReDim COEF_VECTOR(0 To NROWS, 1 To 2)

ReDim BTEMP_VECTOR(1 To 2, 1 To 1)

COEF_VECTOR(NROWS, 1) = (-1) ^ NROWS
COEF_VECTOR(NROWS, 2) = 0
ATEMP_MATRIX = DATA_MATRIX

For j = 1 To NROWS
    For i = 1 To NROWS
        ATEMP_VECTOR(j, 1) = ATEMP_VECTOR(j, 1) + ATEMP_MATRIX(i, i)
        ATEMP_VECTOR(j, 2) = ATEMP_VECTOR(j, 2) + ATEMP_MATRIX(i, NROWS + i)
    Next i
    
    BTEMP_VECTOR(1, 1) = 0
    BTEMP_VECTOR(2, 1) = 0
    
    For i = 0 To j - 1
        BTEMP_VECTOR(1, 1) = BTEMP_VECTOR(1, 1) + COEF_VECTOR(NROWS - i, 1) * _
                    ATEMP_VECTOR(j - i, 1) - COEF_VECTOR(NROWS - i, 2) * _
                            ATEMP_VECTOR(j - i, 2)
        BTEMP_VECTOR(2, 1) = BTEMP_VECTOR(2, 1) + COEF_VECTOR(NROWS - i, 1) * _
                    ATEMP_VECTOR(j - i, 2) + COEF_VECTOR(NROWS - i, 2) * _
                        ATEMP_VECTOR(j - i, 1)
    Next i
    
    COEF_VECTOR(NROWS - j, 1) = -BTEMP_VECTOR(1, 1) / j
    COEF_VECTOR(NROWS - j, 2) = -BTEMP_VECTOR(2, 1) / j
    
    If j < NROWS Then
        BTEMP_MATRIX = COMPLEX_MATRIX_MMULT_FUNC(ATEMP_MATRIX, DATA_MATRIX, _
                        CPLX_FORMAT, CPLX_CHR_STR, epsilon)
        ATEMP_MATRIX = BTEMP_MATRIX
    End If
Next j

'convert matrix for output
If CPLX_FORMAT = 2 Then COEF_VECTOR = COMPLEX_MATRIX_FORMAT_FUNC(COEF_VECTOR, 12, _
                                    CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then COEF_VECTOR = COMPLEX_MATRIX_FORMAT_FUNC(COEF_VECTOR, 13, _
                                    CPLX_CHR_STR, epsilon)
                                    
COMPLEX_MATRIX_CHARACTERISTIC_POLYNOMIAL_FUNC = COEF_VECTOR

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_CHARACTERISTIC_POLYNOMIAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_POLYNOMIAL_QRC_ROOTS_FUNC

'DESCRIPTION   : This function returns all the roots of a given complex polynomial
'Coefficients is an (n+1 x 2) array of complex coefficients
'This function use the QR algorithm for finding the eigenvalues of
'the companion matrix of the given polynomial. This process is very
'fast, robust and stable but may not be converging under certain
'conditions. If the function cannot find a root it returns “#VALUE”.
'Usually it is suitable for solving polynomials up to 10th  degree

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : POLYNOMIAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_POLYNOMIAL_QRC_ROOTS_FUNC(ByRef COEF_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 10 ^ -13)

Dim i As Long
Dim NROWS As Long

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant

Dim COEF_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

COEF_MATRIX = COEF_RNG
'build the companion matrix
DATA_MATRIX = COMPLEX_MATRIX_COMPANION_POLYNOMIAL_FUNC(COEF_MATRIX)
NROWS = UBound(DATA_MATRIX)
'find eigenvalues of companion matrix
ATEMP_VECTOR = COMPLEX_COMPANION_MATRIX_QR_EIGENVALUES_FUNC(DATA_MATRIX, _
               CPLX_FORMAT, CPLX_CHR_STR, epsilon)

ReDim BTEMP_VECTOR(1 To NROWS, 1 To 2)
For i = 1 To NROWS
    BTEMP_VECTOR(i, 1) = ATEMP_VECTOR(i, 1)
    BTEMP_VECTOR(i, 2) = ATEMP_VECTOR(i, 2)
Next i

NROWS = UBound(BTEMP_VECTOR, 1)

ReDim CTEMP_VECTOR(1 To NROWS, 1 To 3)
For i = 1 To NROWS
    CTEMP_VECTOR(i, 2) = BTEMP_VECTOR(i, 1)
    CTEMP_VECTOR(i, 3) = BTEMP_VECTOR(i, 2)
    CTEMP_VECTOR(i, 1) = Sqr(CTEMP_VECTOR(i, 2) ^ 2 + CTEMP_VECTOR(i, 3) ^ 2)
Next i

CTEMP_VECTOR = MATRIX_QUICK_SORT_FUNC(CTEMP_VECTOR, 1, 1)

ReDim BTEMP_VECTOR(1 To NROWS, 1 To 2)
For i = 1 To NROWS
    BTEMP_VECTOR(i, 1) = CTEMP_VECTOR(i, 2)
    BTEMP_VECTOR(i, 2) = CTEMP_VECTOR(i, 3)
Next i

BTEMP_VECTOR = MATRIX_TRIM_SMALL_VALUES_FUNC(BTEMP_VECTOR, epsilon)
COMPLEX_POLYNOMIAL_QRC_ROOTS_FUNC = BTEMP_VECTOR
Exit Function
ERROR_LABEL:
COMPLEX_POLYNOMIAL_QRC_ROOTS_FUNC = Err.number
End Function

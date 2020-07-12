Attribute VB_Name = "NUMBER_COMPLEX_MATRIX_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_FUNC
'DESCRIPTION   : Converts two real matrices into a complex matrix: Ar is the
'(n x m) real part and Ai is the (n x m) imaginary part. The
'real or imaginary part can be omitted. The function assumes
'the zero-matrix for the missing part.

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : MATRIX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_FUNC(Optional ByRef REEL_DATA_RNG As Variant, _
Optional ByRef IMAG_DATA_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

'Example

'A pure real matrix can be written as

'COMPLEX_MATRIX_FUNC(Ar) = [Ar] + j[0]
'A purely imaginary matrix can be written as

'COMPLEX_MATRIX_FUNC(,Ai)= [0] + j[Ai]
'  (remember the comma before the 2nd argument)

'This function supports 3 different formats: 1 = split, 2 = interlaced, 3 = string
'The optional parameter CPLX_FORMAT sets the complex input/output format (default = 1)

'This function is useful to pass a real matrix to a matrix complex function.
'For example we have to multiply a real matrix by a complex vector.
'We can use the complex mult function. But, because this function accepts complex
'matrices, we have to convert the matrix A into a complex one (with a null
'imaginary part) by this function and then pass the result to complex mult
'function.

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim REEL_MATRIX As Variant
Dim IMAG_MATRIX As Variant

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

REEL_MATRIX = REEL_DATA_RNG
IMAG_MATRIX = IMAG_DATA_RNG

If IsArray(REEL_MATRIX) = False And _
IsArray(IMAG_MATRIX) = False Then GoTo ERROR_LABEL

If IsArray(IMAG_MATRIX) = False Then
    ATEMP_MATRIX = REEL_MATRIX
    NROWS = UBound(ATEMP_MATRIX, 1)
    NCOLUMNS = UBound(ATEMP_MATRIX, 2)
    ReDim BTEMP_MATRIX(1 To NROWS, 1 To 2 * NCOLUMNS)
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            BTEMP_MATRIX(i, j) = ATEMP_MATRIX(i, j)
        Next j
    Next i
ElseIf IsArray(REEL_MATRIX) = False Then
    ATEMP_MATRIX = IMAG_MATRIX
    NROWS = UBound(ATEMP_MATRIX, 1)
    NCOLUMNS = UBound(ATEMP_MATRIX, 2)
    ReDim BTEMP_MATRIX(1 To NROWS, 1 To 2 * NCOLUMNS)
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            BTEMP_MATRIX(i, j + NCOLUMNS) = ATEMP_MATRIX(i, j)
        Next j
    Next i
Else
    ATEMP_MATRIX = REEL_MATRIX
    NROWS = UBound(ATEMP_MATRIX, 1)
    NCOLUMNS = UBound(ATEMP_MATRIX, 2)
    ReDim BTEMP_MATRIX(1 To NROWS, 1 To 2 * NCOLUMNS)
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            BTEMP_MATRIX(i, j) = ATEMP_MATRIX(i, j)
        Next j
    Next i
    
    ATEMP_MATRIX = IMAG_MATRIX
    If UBound(ATEMP_MATRIX, 1) = NROWS And UBound(ATEMP_MATRIX, 2) = NCOLUMNS Then
        For i = 1 To NROWS
            For j = 1 To NCOLUMNS
                BTEMP_MATRIX(i, j + NCOLUMNS) = ATEMP_MATRIX(i, j)
            Next j
        Next i
    Else
        GoTo ERROR_LABEL 'matrices with different size
    End If
End If

'convert matrix for output
If CPLX_FORMAT = 2 Then BTEMP_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(BTEMP_MATRIX, 12, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then BTEMP_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(BTEMP_MATRIX, 13, CPLX_CHR_STR, epsilon)

COMPLEX_MATRIX_FUNC = BTEMP_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_ABSOLUTE_VALUE_FUNC

'DESCRIPTION   : Returns the absolute value ||v|| (Euclidean Norm) of
'a complex vector v. Parameter v may be also a complex
'matrix; in this case the function returns the Frobenius
'norm of the matrix. This function supports 3 different
'formats: 1 = split, 2 = interlaced, 3 = string
'Optional parameter CPLX_FORMAT sets the complex format of
'input/output (default = 1)

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : MATRIX
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_ABSOLUTE_VALUE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TEMP_MATRIX = DATA_RNG

If CPLX_FORMAT = 2 Then TEMP_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(TEMP_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then TEMP_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(TEMP_MATRIX, 31, CPLX_CHR_STR, epsilon)

If UBound(TEMP_MATRIX, 2) Mod 2 <> 0 Then GoTo ERROR_LABEL

NROWS = UBound(TEMP_MATRIX, 1)
NCOLUMNS = UBound(TEMP_MATRIX, 2) / 2

For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        TEMP_SUM = TEMP_SUM + _
            TEMP_MATRIX(i, j) ^ 2 + TEMP_MATRIX(i, j + NCOLUMNS) ^ 2
    Next j
Next i

COMPLEX_MATRIX_ABSOLUTE_VALUE_FUNC = (TEMP_SUM) ^ 0.5

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_ABSOLUTE_VALUE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_NORMALIZED_VECTORS_FUNC
'DESCRIPTION   : Returns the normalized vectors of the matrix.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : MATRIX
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_NORMALIZED_VECTORS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NORM_OPT As Integer = 2, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 2 * 10 ^ -13) '5 * 10 ^ -15)

'The optional parameter NORM_OPT indicates what normalization is performed
'Normtype = 1.  All vector’s components are scaled to the min of the absolute
'values. Normtype = 2   (default). All non-zero vectors are length = 1
'Normtype = 3. All vector’s components are scaled to the max of the absolute
'values. This function supports 3 different complex formats: 1 = split,
'2 = interlaced, 3 = string. Optional parameter CPLX_FORMAT sets the complex
'input/output format (default = 1). The optional parameter epsilon  sets the
'minimum error level (default 2E-14). Values under this level will be set to zero.

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_VAL As Double

Dim TEMP_OBJ As Cplx

Dim XCOMPLEX_OBJ As Cplx
Dim YCOMPLEX_OBJ As Cplx

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

If CPLX_FORMAT = 2 Then DATA_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then DATA_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2) / 2

For j = 1 To 2 * NCOLUMNS
    For i = 1 To NROWS
        If Abs(DATA_MATRIX(i, j)) < (2 * 10 ^ -13) Then DATA_MATRIX(i, j) = 0
    Next i
Next j

For j = 1 To NCOLUMNS
    Select Case NORM_OPT
        Case 2  'module =1
            TEMP_SUM = 0 '
            For i = 1 To NROWS
                 TEMP_VAL = DATA_MATRIX(i, j) ^ 2 + _
                    DATA_MATRIX(i, j + NCOLUMNS) ^ 2
                 TEMP_SUM = TEMP_SUM + TEMP_VAL
            Next i
            TEMP_SUM = Sqr(TEMP_SUM)
        Case 3  'max element =1
            TEMP_SUM = 0 '
            For i = 1 To NROWS
                TEMP_VAL = DATA_MATRIX(i, j) ^ 2 + _
                    DATA_MATRIX(i, j + NCOLUMNS) ^ 2
                If TEMP_VAL > TEMP_SUM Then
                    XCOMPLEX_OBJ.reel = DATA_MATRIX(i, j)
                    XCOMPLEX_OBJ.imag = DATA_MATRIX(i, j + NCOLUMNS)
                    TEMP_SUM = TEMP_VAL
                End If
            Next i
        Case 1  'min element =1
            TEMP_SUM = 10 ^ 300 '
            For i = 1 To NROWS
                TEMP_VAL = DATA_MATRIX(i, j) ^ 2 + _
                    DATA_MATRIX(i, j + NCOLUMNS) ^ 2
                If TEMP_VAL > 1000 * (2 * 10 ^ -13) _
                        And TEMP_VAL < TEMP_SUM Then
                    XCOMPLEX_OBJ.reel = DATA_MATRIX(i, j)
                    XCOMPLEX_OBJ.imag = DATA_MATRIX(i, j + NCOLUMNS)
                    TEMP_SUM = TEMP_VAL
                End If
            Next i
    End Select
    
    If Abs(TEMP_SUM) > (2 * 10 ^ -13) Then
        If NORM_OPT = 2 Then
            For i = 1 To NROWS
                DATA_MATRIX(i, j) = DATA_MATRIX(i, j) / TEMP_SUM
                DATA_MATRIX(i, j + NCOLUMNS) = DATA_MATRIX(i, j + _
                    NCOLUMNS) / TEMP_SUM
            Next i
        Else
            For i = 1 To NROWS
                YCOMPLEX_OBJ.reel = DATA_MATRIX(i, j)
                YCOMPLEX_OBJ.imag = DATA_MATRIX(i, j + NCOLUMNS)
                TEMP_OBJ = COMPLEX_QUOTIENT_OBJ_FUNC(YCOMPLEX_OBJ, XCOMPLEX_OBJ)
                DATA_MATRIX(i, j) = TEMP_OBJ.reel
                DATA_MATRIX(i, j + NCOLUMNS) = TEMP_OBJ.imag
                If Abs(DATA_MATRIX(i, j)) < (2 * 10 ^ -13) Then _
                        DATA_MATRIX(i, j) = 0
                If Abs(DATA_MATRIX(i, j + NCOLUMNS)) < (2 * 10 ^ -13) Then _
                    DATA_MATRIX(i, j + NCOLUMNS) = 0
            Next i
        End If
    End If
Next j

If CPLX_FORMAT = 2 Then DATA_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then DATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

COMPLEX_MATRIX_NORMALIZED_VECTORS_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_NORMALIZED_VECTORS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_TRANSPOSE_FUNC

'DESCRIPTION   : Returns the transpose of a complex matrix, that is the matrix
'with rows and columns exchanged. If CONJUGATE_FLAG = True the it returns
'the transpose-conjugate of a complex matrix. This function supports
'3 different formats: 1 = split, 2 = interlaced, 3 = string.

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : MATRIX
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_TRANSPOSE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONJUGATE_FLAG As Boolean = False, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ADATA_MATRIX As Variant
Dim BDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

ADATA_MATRIX = DATA_RNG
If CPLX_FORMAT = 2 Then ADATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(ADATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then ADATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(ADATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

If UBound(ADATA_MATRIX, 2) Mod 2 <> 0 Then GoTo ERROR_LABEL
NROWS = UBound(ADATA_MATRIX, 1)
NCOLUMNS = UBound(ADATA_MATRIX, 2) / 2

ReDim BDATA_MATRIX(1 To NCOLUMNS, 1 To 2 * NROWS)
For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        BDATA_MATRIX(j, i) = ADATA_MATRIX(i, j)
        BDATA_MATRIX(j, i + NROWS) = ADATA_MATRIX(i, j + NCOLUMNS)
        If BDATA_MATRIX(j, i + NROWS) <> 0 And CONJUGATE_FLAG Then _
            BDATA_MATRIX(j, i + NROWS) = -BDATA_MATRIX(j, i + NROWS)
    Next j
Next i

'----------------------Convert matrix for output-----------------------------
If CPLX_FORMAT = 2 Then BDATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(BDATA_MATRIX, 12, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then BDATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(BDATA_MATRIX, 13, CPLX_CHR_STR, epsilon)
'----------------------------------------------------------------------------

COMPLEX_MATRIX_TRANSPOSE_FUNC = BDATA_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_TRANSPOSE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_CHARACTERISTIC_FUNC

'DESCRIPTION   : 'This function returns the complex characteristic matrix at the
'complex value z. C = A - zI
'The determinant of C is the characteristic polynomial of the matrix A
'A can be a real or complex square matrix
'z can be a real or complex number.
'This function supports 3 different formats: 1 = split, 2 = interlaced, 3 = string
'Optional parameter CPLX_FORMAT sets the complex input/output format (default = 1)

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : MATRIX
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_CHARACTERISTIC_FUNC(ByRef DATA_RNG As Variant, _
ByVal SCALAR_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

Dim NSIZE As Long

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX)
TEMP_MATRIX = MATRIX_IDENTITY_FUNC(NSIZE)  'identity matrix

ReDim Preserve TEMP_MATRIX(1 To NSIZE, 1 To 2 * NSIZE)

If CPLX_FORMAT = 2 Then TEMP_MATRIX = _
            COMPLEX_MATRIX_FORMAT_FUNC(TEMP_MATRIX, 12, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then TEMP_MATRIX = _
            COMPLEX_MATRIX_FORMAT_FUNC(TEMP_MATRIX, 13, CPLX_CHR_STR, epsilon)


If CPLX_FORMAT = 1 And (UBound(DATA_MATRIX, 1) = UBound(DATA_MATRIX, 2)) Then
    ReDim Preserve DATA_MATRIX(1 To UBound(DATA_MATRIX, 1), _
    1 To 2 * UBound(DATA_MATRIX, 1))
End If

TEMP_MATRIX = COMPLEX_MATRIX_SCALAR_MULT_FUNC(TEMP_MATRIX, SCALAR_RNG, _
                CPLX_FORMAT, CPLX_CHR_STR, epsilon)

COMPLEX_MATRIX_CHARACTERISTIC_FUNC = COMPLEX_MATRIX_SUBTRACTION_FUNC(DATA_MATRIX, TEMP_MATRIX, _
                            CPLX_FORMAT, CPLX_CHR_STR, epsilon)
    
Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_CHARACTERISTIC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_SUM_FUNC

'DESCRIPTION   : Returns the sum of two complex matrices. This function supports 3
'different formats: 1 = split, 2 = interlaced, 3 = string. The optional
'parameter Cformat sets the complex input/output format (default = 1)

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : MATRIX
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_SUM_FUNC(ByRef ADATA_RNG As Variant, _
ByRef BDATA_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15, _
Optional ByVal FACTOR_VAL As Double = 1)

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim ANROWS As Long
Dim ANCOLUMNS As Long

Dim BNROWS As Long
Dim BNCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim ADATA_MATRIX As Variant
Dim BDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

ADATA_MATRIX = ADATA_RNG
If CPLX_FORMAT = 2 Then ADATA_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(ADATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then ADATA_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(ADATA_MATRIX, 31, CPLX_CHR_STR, epsilon)
BDATA_MATRIX = BDATA_RNG

If CPLX_FORMAT = 2 Then BDATA_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(BDATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then BDATA_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(BDATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

ANROWS = UBound(ADATA_MATRIX, 1)
ANCOLUMNS = UBound(ADATA_MATRIX, 2)

BNROWS = UBound(BDATA_MATRIX, 1)
BNCOLUMNS = UBound(BDATA_MATRIX, 2)

If ANCOLUMNS Mod 2 <> 0 Or BNCOLUMNS Mod 2 <> 0 Then GoTo ERROR_LABEL

ANCOLUMNS = ANCOLUMNS / 2
BNCOLUMNS = BNCOLUMNS / 2

If ANCOLUMNS <> BNCOLUMNS Or ANROWS <> BNROWS Then GoTo ERROR_LABEL

ii = ANROWS
jj = ANCOLUMNS

ReDim TEMP_MATRIX(1 To ii, 1 To 2 * jj)

For i = 1 To ii
    For j = 1 To 2 * jj
        TEMP_MATRIX(i, j) = ADATA_MATRIX(i, j) + FACTOR_VAL * BDATA_MATRIX(i, j)
        If Abs(TEMP_MATRIX(i, j)) < (10 ^ -15) Then TEMP_MATRIX(i, j) = 0
    Next j
Next i

If CPLX_FORMAT = 2 Then TEMP_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(TEMP_MATRIX, 12, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then TEMP_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(TEMP_MATRIX, 13, CPLX_CHR_STR, epsilon)

COMPLEX_MATRIX_SUM_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_SUM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_SUBTRACTION_FUNC
'DESCRIPTION   : Performs complex matrix subtraction
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : MATRIX
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_SUBTRACTION_FUNC(ByRef ADATA_RNG As Variant, _
ByRef BDATA_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

COMPLEX_MATRIX_SUBTRACTION_FUNC = COMPLEX_MATRIX_SUM_FUNC(ADATA_RNG, BDATA_RNG, _
CPLX_FORMAT, CPLX_CHR_STR, epsilon, -1)

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_SUBTRACTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_SCALAR_PRODUCT_FUNC

'DESCRIPTION   : Returns the scalar product of two complex vectors
'This function now supports 3 different formats:
'1 = split, 2 = interlaced, 3 = string
'Complex split or interlaced matrix must have always an
'even number of columns

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : MATRIX
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_SCALAR_PRODUCT_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 3, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

Dim i As Long
Dim ANSIZE As Long
Dim BNSIZE As Long

Dim REEL_VALUE As Double
Dim IMAG_VALUE As Double

Dim DATA_MATRIX As Variant
Dim DATA_VECTOR As Variant

Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_RNG
DATA_VECTOR = VECTOR_RNG

ANSIZE = UBound(DATA_MATRIX)
BNSIZE = UBound(DATA_VECTOR)

If ANSIZE <> BNSIZE Then: GoTo ERROR_LABEL

If CPLX_FORMAT = 3 Then
    DATA_MATRIX = COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)
    DATA_VECTOR = COMPLEX_MATRIX_FORMAT_FUNC(DATA_VECTOR, 31, CPLX_CHR_STR, epsilon)
End If

REEL_VALUE = 0
IMAG_VALUE = 0

For i = 1 To ANSIZE
    REEL_VALUE = REEL_VALUE + DATA_MATRIX(i, 1) * _
        DATA_VECTOR(i, 1) + DATA_MATRIX(i, 2) * DATA_VECTOR(i, 2)
    IMAG_VALUE = IMAG_VALUE - DATA_MATRIX(i, 1) * _
        DATA_VECTOR(i, 2) + DATA_MATRIX(i, 2) * DATA_VECTOR(i, 1)
Next i

If CPLX_FORMAT = 3 Then
    COMPLEX_MATRIX_SCALAR_PRODUCT_FUNC = COMPLEX_CONVERT_COEFFICIENTS_FUNC(REEL_VALUE, IMAG_VALUE, _
        CPLX_CHR_STR, epsilon)
Else
    ReDim TEMP_VECTOR(1 To 1, 1 To 2)
    TEMP_VECTOR(1, 1) = REEL_VALUE
    TEMP_VECTOR(1, 2) = IMAG_VALUE
    
    COMPLEX_MATRIX_SCALAR_PRODUCT_FUNC = TEMP_VECTOR
End If

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_SCALAR_PRODUCT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_SCALAR_MULT_FUNC

'DESCRIPTION   : Multiplies a complex matrix by a complex scalar
'The parameter DATA_RNG is a (n x m) complex matrix or vector
'The parameter scalar can be a complex or real number, in split or string format.
'This function supports 3 different formats: 1 = split, 2 = interlaced, 3 = string
'The optional parameter CPLX_FORMAT sets the complex input/output format (default = 1)

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : MATRIX
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_SCALAR_MULT_FUNC(ByRef DATA_RNG As Variant, _
ByVal SCALAR_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim REEL_VAL As Double
Dim IMAG_VAL As Double

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

If CPLX_FORMAT = 2 Then DATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then DATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

'take the complex scalar
TEMP_VECTOR = COMPLEX_EXTRACT_NUMBER_FUNC(SCALAR_RNG, CPLX_CHR_STR)
REEL_VAL = TEMP_VECTOR(1, 1)
IMAG_VAL = TEMP_VECTOR(2, 1)

'------------------
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2) / 2

For i = 1 To NROWS 'perform the complex scalar multiplication
    For j = 1 To NCOLUMNS
        ATEMP_VAL = DATA_MATRIX(i, j)
        BTEMP_VAL = DATA_MATRIX(i, j + NCOLUMNS)
        DATA_MATRIX(i, j) = REEL_VAL * ATEMP_VAL - _
            IMAG_VAL * BTEMP_VAL     'product real
        DATA_MATRIX(i, j + NCOLUMNS) = REEL_VAL * _
            BTEMP_VAL + IMAG_VAL * ATEMP_VAL  'product imaginary
    Next j
Next i

'convert matrix for output
If CPLX_FORMAT = 2 Then DATA_MATRIX = COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 12, CPLX_CHR_STR)
If CPLX_FORMAT = 3 Then DATA_MATRIX = COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 13, CPLX_CHR_STR)

COMPLEX_MATRIX_SCALAR_MULT_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_SCALAR_MULT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_MMULT_FUNC

'DESCRIPTION   : Performs a complex matrix multiplication.
'If the dimension of the matrix M1 is (n x m)
'and M2 is (m x p) , then the product is a matrix (n x p)
'This function now supports 3 different formats: 1 = split,
'2 = interlaced, 3 = string.

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : MATRIX
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_MMULT_FUNC(ByRef ADATA_RNG As Variant, _
ByRef BDATA_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

'performs complex matrix multiplication
'TEMP_ARR = ADATA_MATRIX x BDATA_MATRIX

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim ANROWS As Long
Dim ANCOLUMNS As Long

Dim BNROWS As Long
Dim BNCOLUMNS As Long

Dim FIRST_REEL As Double
Dim FIRST_IMAG As Double

Dim SECOND_REEL As Double
Dim SECOND_IMAG As Double

Dim TEMP_MATRIX As Variant

Dim ADATA_MATRIX As Variant
Dim BDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

'------------------------------------------------------------------------
ADATA_MATRIX = ADATA_RNG
If CPLX_FORMAT = 2 Then _
    ADATA_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(ADATA_MATRIX, 21, CPLX_CHR_STR, epsilon)

If CPLX_FORMAT = 3 Then ADATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(ADATA_MATRIX, 31, CPLX_CHR_STR, epsilon)
    
'------------------------------------------------------------------------
BDATA_MATRIX = BDATA_RNG

If CPLX_FORMAT = 2 Then BDATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(BDATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then BDATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(BDATA_MATRIX, 31, CPLX_CHR_STR, epsilon)
'------------------------------------------------------------------------

ANROWS = UBound(ADATA_MATRIX, 1)
ANCOLUMNS = UBound(ADATA_MATRIX, 2)

BNROWS = UBound(BDATA_MATRIX, 1)
BNCOLUMNS = UBound(BDATA_MATRIX, 2)

If ANCOLUMNS Mod 2 <> 0 Or BNCOLUMNS Mod 2 <> 0 Then GoTo ERROR_LABEL

ANCOLUMNS = ANCOLUMNS / 2
BNCOLUMNS = BNCOLUMNS / 2

If ANCOLUMNS <> BNROWS Then GoTo ERROR_LABEL

ii = ANROWS
jj = BNCOLUMNS

ReDim TEMP_MATRIX(1 To ii, 1 To 2 * jj)

For i = 1 To ii
    For j = 1 To jj
        For k = 1 To ANCOLUMNS
            FIRST_REEL = ADATA_MATRIX(i, k)
            FIRST_IMAG = ADATA_MATRIX(i, k + ANCOLUMNS)
            
            SECOND_REEL = BDATA_MATRIX(k, j)
            SECOND_IMAG = BDATA_MATRIX(k, j + BNCOLUMNS)
            
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + FIRST_REEL * _
                SECOND_REEL - FIRST_IMAG * SECOND_IMAG
            If Abs(TEMP_MATRIX(i, j)) < epsilon Then TEMP_MATRIX(i, j) = 0
            
            TEMP_MATRIX(i, j + jj) = _
                TEMP_MATRIX(i, j + jj) + FIRST_REEL * _
                    SECOND_IMAG + FIRST_IMAG * SECOND_REEL
            If Abs(TEMP_MATRIX(i, j + jj)) < epsilon Then _
                TEMP_MATRIX(i, j + jj) = 0
        Next k
    Next j
Next i

If CPLX_FORMAT = 2 Then TEMP_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(TEMP_MATRIX, 12, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then TEMP_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(TEMP_MATRIX, 13, CPLX_CHR_STR, epsilon)

COMPLEX_MATRIX_MMULT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_MMULT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_POWER_FUNC
'DESCRIPTION   : Returns the integer power of a complex square matrix
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : MATRIX
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_POWER_FUNC(ByRef DATA_RNG As Variant, _
ByVal NSIZE As Long, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

Dim i As Long
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
TEMP_MATRIX = DATA_MATRIX

For i = 1 To NSIZE - 1
    TEMP_MATRIX = COMPLEX_MATRIX_MMULT_FUNC(DATA_MATRIX, TEMP_MATRIX, _
                  CPLX_FORMAT, CPLX_CHR_STR, epsilon)
Next i

COMPLEX_MATRIX_POWER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_POWER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_DETERMINANT_FUNC

'DESCRIPTION   : This function computes the determinant of a complex matrix.

'The argument DATA_RNG is an array (n x n ) or (n x 2n) ,
'depending of the format parameter. This function supports 3
'different formats: 1 = split, 2 = interlaced, 3 = string the

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : INVERSE & DETERMINANT
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_DETERMINANT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

Dim ERROR_STR As String

On Error GoTo ERROR_LABEL

ERROR_STR = ""
DATA_MATRIX = DATA_RNG

If CPLX_FORMAT = 2 Then _
    DATA_MATRIX = COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 21, CPLX_CHR_STR, epsilon)

If CPLX_FORMAT = 3 Then _
    DATA_MATRIX = COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

'optional complex part for format=1 and square matrices
If CPLX_FORMAT = 1 And (UBound(DATA_MATRIX, 1) = _
    UBound(DATA_MATRIX, 2)) Then
    ReDim Preserve DATA_MATRIX(1 To UBound(DATA_MATRIX, 1), _
    1 To 2 * UBound(DATA_MATRIX, 1))
End If

If 2 * UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then GoTo ERROR_LABEL

Call COMPLEX_MATRIX_GJ_SOLVER_FUNC(DATA_MATRIX, , TEMP_VECTOR, _
                                          ERROR_STR, epsilon)

If ERROR_STR <> "" Then
    ERROR_STR = "SINGULAR"
    GoTo ERROR_LABEL
End If

'------------------------convert matrix for output----------------------------
If CPLX_FORMAT = 3 Then
    COMPLEX_MATRIX_DETERMINANT_FUNC = COMPLEX_CONVERT_COEFFICIENTS_FUNC(TEMP_VECTOR(1, 1), _
    TEMP_VECTOR(2, 1), CPLX_CHR_STR, epsilon)
Else
    COMPLEX_MATRIX_DETERMINANT_FUNC = TEMP_VECTOR
End If
'------------------------------------------------------------------------------


Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_DETERMINANT_FUNC = ERROR_STR
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_INVERSE_FUNC
'DESCRIPTION   : 'Returns the inverse of a complex matrix. The complex matrix
'A must be square. This function supports 3 different formats: 1 = split,
'2 = interlaced, 3 = string; Complex split or interlaced matrix must have
'always an even number of columns.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : INVERSE & DETERMINANT
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_INVERSE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

Dim DATA_MATRIX As Variant
Dim ERROR_STR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

If CPLX_FORMAT = 2 Then _
    DATA_MATRIX = COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then _
    DATA_MATRIX = COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

'optional complex part for format=1 and square matrices
If CPLX_FORMAT = 1 And (UBound(DATA_MATRIX, 1) = _
    UBound(DATA_MATRIX, 2)) Then
    ReDim Preserve DATA_MATRIX(1 To UBound(DATA_MATRIX, 1), _
    1 To 2 * UBound(DATA_MATRIX, 1))
End If

If 2 * UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then GoTo ERROR_LABEL

Call COMPLEX_MATRIX_GJ_SOLVER_FUNC(DATA_MATRIX, , , ERROR_STR, epsilon)

If ERROR_STR <> "" Then
    COMPLEX_MATRIX_INVERSE_FUNC = "SINGULAR"
    Exit Function
End If

'------------------------convert matrix for output----------------------------
If CPLX_FORMAT = 2 Then _
    DATA_MATRIX = COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 12, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then _
    DATA_MATRIX = COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 13, CPLX_CHR_STR, epsilon)
'------------------------------------------------------------------------------

COMPLEX_MATRIX_INVERSE_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_INVERSE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_LINEAR_SYSTEM_FUNC
'DESCRIPTION   : Gauss-Jordan algorithm for complex matrix reduction with full
'Pivot method
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : LINEAR SYSTEM
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_LINEAR_SYSTEM_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

'Solve a complex linear system A*X=B

'This function solves a complex linear system by the Gauss-Jordan algorithm.
'MATRIX_RNG; is the system complex square matrix (n x 2*n)
'VECTOR_RNG; is the known complex vector (n x 2)
'As known, the above linear equation has only one solution if,
'and only if, det(A) <> 0. This function now supports 3 different
'formats: 1 = split, 2 = interlaced, 3 = string

Dim ANROWS As Long
Dim ANCOLUMNS As Long

Dim BNROWS As Long
Dim BNCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim DATA_VECTOR As Variant

Dim ERROR_STR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_RNG
If CPLX_FORMAT = 2 Then DATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then _
    DATA_MATRIX = COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

DATA_VECTOR = VECTOR_RNG
If CPLX_FORMAT = 2 Then _
    DATA_VECTOR = COMPLEX_MATRIX_FORMAT_FUNC(DATA_VECTOR, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then _
    DATA_VECTOR = COMPLEX_MATRIX_FORMAT_FUNC(DATA_VECTOR, 31, CPLX_CHR_STR, epsilon)

ANROWS = UBound(DATA_MATRIX, 1)
ANCOLUMNS = UBound(DATA_MATRIX, 2)
BNROWS = UBound(DATA_VECTOR, 1)
BNCOLUMNS = UBound(DATA_VECTOR, 2)

If ANCOLUMNS Mod 2 <> 0 Or BNCOLUMNS Mod 2 <> 0 Then GoTo ERROR_LABEL
ANCOLUMNS = ANCOLUMNS / 2
BNCOLUMNS = BNCOLUMNS / 2
If ANROWS <> ANCOLUMNS Or ANROWS <> BNROWS Then GoTo ERROR_LABEL

Call COMPLEX_MATRIX_GJ_SOLVER_FUNC(DATA_MATRIX, DATA_VECTOR, , _
                                          ERROR_STR, epsilon)

If ERROR_STR <> "" Then
    COMPLEX_MATRIX_LINEAR_SYSTEM_FUNC = "SINGULAR"
    Exit Function
End If


DATA_VECTOR = MATRIX_TRIM_SMALL_VALUES_FUNC(DATA_VECTOR, epsilon)
If CPLX_FORMAT = 2 Then DATA_VECTOR = _
    COMPLEX_MATRIX_FORMAT_FUNC(DATA_VECTOR, 12, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then DATA_VECTOR = _
    COMPLEX_MATRIX_FORMAT_FUNC(DATA_VECTOR, 13, CPLX_CHR_STR, epsilon)

COMPLEX_MATRIX_LINEAR_SYSTEM_FUNC = DATA_VECTOR

Exit Function
ERROR_LABEL:
    COMPLEX_MATRIX_LINEAR_SYSTEM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_GJ_SOLVER_FUNC
'DESCRIPTION   : Complex Gauss Jordan Routine
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : LINEAR SYSTEM
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_GJ_SOLVER_FUNC(ByRef DATA_MATRIX As Variant, _
Optional ByRef DATA_VECTOR As Variant, _
Optional ByRef DETERM_VECTOR As Variant, _
Optional ByRef ERROR_STR As Variant, _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)


'Gauss-Jordan algorithm for matrix reduction with full pivot method
'DATA_MATRIX is a matrix (NROWS x NROWS); at the end
'contains the inverse of AB is a matrix
'(NROWS x NCOLUMNS); at the end contains the solution of AX=B
'this version apply the check for too small elements: |aij|<Tiny

'DATA_MATRIX is a matrix (NROWS x 2*NROWS);

'A=[Are, Aim] at the end contains the inverse of DATA_MATRIX
'DATA_VECTOR is a matrix (NROWS x 2*NCOLUMNS);
'B=[Bre, Bim] at the end contains the solution of AX=B

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MAX_VAL As Double

Dim ACOMPLEX_OBJ As Cplx
Dim BCOMPLEX_OBJ As Cplx

Dim ACOMPLEX_MATRIX() As Cplx
Dim BCOMPLEX_MATRIX() As Cplx

Dim TEMP_MATRIX As Variant
Dim DETERM_FLAG As Boolean

On Error GoTo ERROR_LABEL

COMPLEX_MATRIX_GJ_SOLVER_FUNC = False

ERROR_STR = ""
If Not IsMissing(DETERM_VECTOR) = True Then
    DETERM_FLAG = True
    ReDim DETERM_VECTOR(1 To 2, 1 To 1)
End If

If IsArray(DATA_VECTOR) = True Then
    NCOLUMNS = UBound(DATA_VECTOR, 2) / 2
Else
    NCOLUMNS = 0
End If

NROWS = UBound(DATA_MATRIX, 1)
ReDim TEMP_MATRIX(1 To 2 * NROWS, 1 To 3) 'trace of swaps

kk = 0 'swap COUNTER
BCOMPLEX_OBJ.reel = 1
BCOMPLEX_OBJ.imag = 0

'convert the original matrices in complex matrices

ReDim ACOMPLEX_MATRIX(1 To NROWS, 1 To NROWS)
If NCOLUMNS > 0 Then ReDim BCOMPLEX_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For i = 1 To NROWS
    For j = 1 To NROWS
        ACOMPLEX_MATRIX(i, j).reel = DATA_MATRIX(i, j)
        ACOMPLEX_MATRIX(i, j).imag = DATA_MATRIX(i, j + NROWS)
    Next j
Next i

For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        BCOMPLEX_MATRIX(i, j).reel = DATA_VECTOR(i, j)
        BCOMPLEX_MATRIX(i, j).imag = DATA_VECTOR(i, j + NCOLUMNS)
    Next i
Next j

For k = 1 To NROWS 'search max pivot
    ii = k
    jj = k
    
    MAX_VAL = 0
    For i = k To NROWS
        For j = k To NROWS
            If COMPLEX_MODULUS_OBJ_FUNC(ACOMPLEX_MATRIX(i, j)) > MAX_VAL Then
                ii = i
                jj = j
                MAX_VAL = COMPLEX_MODULUS_OBJ_FUNC(ACOMPLEX_MATRIX(i, j))
            End If
        Next j
    Next i
    
    ' swap rows and columns
    If ii > k Then
        Call CPLX_MATRIX_SWAP_ROW_FUNC(ACOMPLEX_MATRIX, k, ii)
        If NCOLUMNS > 0 Then Call CPLX_MATRIX_SWAP_ROW_FUNC(ACOMPLEX_MATRIX, k, ii)
        If DETERM_FLAG Then BCOMPLEX_OBJ = COMPLEX_NEGATIVE_OBJ_FUNC(BCOMPLEX_OBJ)
            kk = kk + 1
            TEMP_MATRIX(kk, 1) = k
            TEMP_MATRIX(kk, 2) = ii
            TEMP_MATRIX(kk, 3) = 1
    End If
    If jj > k Then
        Call CPLX_MATRIX_SWAP_COLUMN_FUNC(ACOMPLEX_MATRIX, k, jj)
        If DETERM_FLAG Then BCOMPLEX_OBJ = COMPLEX_NEGATIVE_OBJ_FUNC(BCOMPLEX_OBJ)
            kk = kk + 1
            TEMP_MATRIX(kk, 1) = k
            TEMP_MATRIX(kk, 2) = jj
            TEMP_MATRIX(kk, 3) = 2
    End If
    ' check pivot 0
    
    If COMPLEX_MODULUS_OBJ_FUNC(ACOMPLEX_MATRIX(k, k)) <= epsilon Then
            ACOMPLEX_MATRIX(k, k).reel = 0
            ACOMPLEX_MATRIX(k, k).imag = 0
            
            If DETERM_FLAG Then
                DETERM_VECTOR(1, 1) = 0
                DETERM_VECTOR(2, 1) = 0 '"singular"
            End If
            ERROR_STR = "SINGULAR"
        GoTo ERROR_LABEL
    End If
    
    'normalization
    ACOMPLEX_OBJ = ACOMPLEX_MATRIX(k, k)
    If DETERM_FLAG Then BCOMPLEX_OBJ = COMPLEX_PRODUCT_OBJ_FUNC(BCOMPLEX_OBJ, ACOMPLEX_OBJ)
       ACOMPLEX_MATRIX(k, k).reel = 1
        ACOMPLEX_MATRIX(k, k).imag = 0
    
    For j = 1 To NROWS
        ACOMPLEX_MATRIX(k, j) = COMPLEX_QUOTIENT_OBJ_FUNC(ACOMPLEX_MATRIX(k, j), ACOMPLEX_OBJ)
    Next j
    For j = 1 To NCOLUMNS
        BCOMPLEX_MATRIX(k, j) = COMPLEX_QUOTIENT_OBJ_FUNC(BCOMPLEX_MATRIX(k, j), ACOMPLEX_OBJ)
    Next j
    
    'linear reduction
    For i = 1 To NROWS
        If i <> k And Not COMPLEX_ZERO_OBJ_FUNC(ACOMPLEX_MATRIX(i, k), epsilon) Then
            ACOMPLEX_OBJ = ACOMPLEX_MATRIX(i, k)
            ACOMPLEX_MATRIX(i, k).reel = 0
            ACOMPLEX_MATRIX(i, k).imag = 0
            
            For j = 1 To NROWS
                ACOMPLEX_MATRIX(i, j) = COMPLEX_SUBTRACTION_OBJ_FUNC(ACOMPLEX_MATRIX(i, j), _
                COMPLEX_PRODUCT_OBJ_FUNC(ACOMPLEX_OBJ, ACOMPLEX_MATRIX(k, j)))
            Next j
            For j = 1 To NCOLUMNS
                BCOMPLEX_MATRIX(i, j) = COMPLEX_SUBTRACTION_OBJ_FUNC(BCOMPLEX_MATRIX(i, j), _
                COMPLEX_PRODUCT_OBJ_FUNC(ACOMPLEX_OBJ, BCOMPLEX_MATRIX(k, j)))
            Next j
        End If
    Next i
Next k

'convert the matrices from complex to double
For i = 1 To NROWS
    For j = 1 To NROWS
       DATA_MATRIX(i, j) = ACOMPLEX_MATRIX(i, j).reel
       DATA_MATRIX(i, j + NROWS) = ACOMPLEX_MATRIX(i, j).imag
   Next j
Next i
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        DATA_VECTOR(i, j) = BCOMPLEX_MATRIX(i, j).reel
        DATA_VECTOR(i, j + NCOLUMNS) = BCOMPLEX_MATRIX(i, j).imag
    Next i
Next j

'scramble rows
For i = kk To 1 Step -1
    If TEMP_MATRIX(i, 3) = 1 Then
        DATA_MATRIX = COMPLEX_MATRIX_SWAP_COLUMN_FUNC(DATA_MATRIX, _
            TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2))
    Else
        DATA_MATRIX = MATRIX_SWAP_ROW_FUNC(DATA_MATRIX, TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2))
        If NCOLUMNS > 0 Then DATA_VECTOR = MATRIX_SWAP_ROW_FUNC(DATA_VECTOR, _
            TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2))
    End If
Next i

If DETERM_FLAG Then
    DETERM_VECTOR(1, 1) = BCOMPLEX_OBJ.reel
    DETERM_VECTOR(2, 1) = BCOMPLEX_OBJ.imag
End If

COMPLEX_MATRIX_GJ_SOLVER_FUNC = True

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_GJ_SOLVER_FUNC = False
End Function

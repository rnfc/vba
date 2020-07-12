Attribute VB_Name = "MATRIX_ARITHM_MULT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_ELEMENTS_3D_PRODUCT_FUNC
'DESCRIPTION   : Return vector product (only 3 dimension)
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_ELEMENTS_3D_PRODUCT_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant)

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then
    DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)
End If
DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then
    DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
End If
If UBound(DATA1_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Or UBound(DATA1_VECTOR, 1) <> 3 Then: GoTo ERROR_LABEL
ReDim TEMP_MATRIX(1 To 3, 1 To 1)
TEMP_MATRIX(1, 1) = DATA1_VECTOR(2, 1) * DATA2_VECTOR(3, 1) - DATA2_VECTOR(2, 1) * DATA1_VECTOR(3, 1)
TEMP_MATRIX(2, 1) = DATA1_VECTOR(3, 1) * DATA2_VECTOR(1, 1) - DATA1_VECTOR(1, 1) * DATA2_VECTOR(3, 1)
TEMP_MATRIX(3, 1) = DATA1_VECTOR(1, 1) * DATA2_VECTOR(2, 1) - DATA2_VECTOR(1, 1) * DATA1_VECTOR(2, 1)

VECTOR_ELEMENTS_3D_PRODUCT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
VECTOR_ELEMENTS_3D_PRODUCT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_ELEMENTS_DOT_PRODUCT_FUNC
'DESCRIPTION   : Dot Product Mult
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_ELEMENTS_DOT_PRODUCT_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant)
    
Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant
        
On Error GoTo ERROR_LABEL
           
DATA1_VECTOR = DATA1_RNG
DATA2_VECTOR = DATA2_RNG

If IS_2D_ARRAY_FUNC(DATA1_VECTOR) And IS_2D_ARRAY_FUNC(DATA2_VECTOR) Then
    NROWS = UBound(DATA1_VECTOR, 1)
    TEMP_SUM = 0
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + DATA1_VECTOR(i, 1) * DATA2_VECTOR(i, 1)
    Next i
ElseIf IS_1D_ARRAY_FUNC(DATA1_VECTOR) And IS_1D_ARRAY_FUNC(DATA2_VECTOR) Then
    NROWS = UBound(DATA1_VECTOR)
    TEMP_SUM = 0
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + DATA1_VECTOR(i) * DATA2_VECTOR(i)
    Next i
Else
    GoTo ERROR_LABEL
End If

VECTOR_ELEMENTS_DOT_PRODUCT_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
VECTOR_ELEMENTS_DOT_PRODUCT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_ELEMENTS_MULT_SCALAR_FUNC
'DESCRIPTION   : Multiplies all the numbers in a vector by a given SCALAR_VAL
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_ELEMENTS_MULT_SCALAR_FUNC(ByRef DATA_RNG As Variant, _
ByVal SCALAR_VAL As Double, _
Optional ByVal VERSION As Integer = 0)
    
Dim i As Long
Dim NCOLUMNS As Long
Dim DATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
Select Case VERSION
Case 0
    NCOLUMNS = UBound(DATA_VECTOR, 2)
    ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)
    For i = 1 To NCOLUMNS
        TEMP_VECTOR(1, i) = DATA_VECTOR(1, i) * SCALAR_VAL
    Next i
Case 1
    NCOLUMNS = UBound(DATA_VECTOR, 2)
    ReDim TEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        TEMP_VECTOR(i, 1) = DATA_VECTOR(1, i) * SCALAR_VAL
    Next i
Case 2
    NCOLUMNS = UBound(DATA_VECTOR, 1)
    ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)
    For i = 1 To NCOLUMNS
        TEMP_VECTOR(1, i) = DATA_VECTOR(i, 1) * SCALAR_VAL
    Next i
Case Else
    NCOLUMNS = UBound(DATA_VECTOR, 1)
    ReDim TEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        TEMP_VECTOR(i, 1) = DATA_VECTOR(i, 1) * SCALAR_VAL
    Next i
End Select

VECTOR_ELEMENTS_MULT_SCALAR_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_ELEMENTS_MULT_SCALAR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_ELEMENTS_MULT_FUNC
'DESCRIPTION   : This routine multiplies two vectors
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_ELEMENTS_MULT_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant)
    
Dim i As Long
Dim NROWS As Long

Dim TEMP_VECTOR As Variant
Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

On Error GoTo ERROR_LABEL
           
DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then
    DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)
End If

DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then
    DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
End If
If UBound(DATA1_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(DATA1_VECTOR, 1)
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    TEMP_VECTOR(i, 1) = DATA1_VECTOR(i, 1) * DATA2_VECTOR(i, 1)
Next i
    
VECTOR_ELEMENTS_MULT_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_ELEMENTS_MULT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_ELEMENTS_MATRIX_MULT_FUNC
'DESCRIPTION   : n x 1 Vector into (n x n) mult matrix
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_ELEMENTS_MATRIX_MULT_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)

For i = 1 To NROWS
    TEMP_MATRIX(i, i) = DATA_VECTOR(i, 1) * DATA_VECTOR(i, 1)
    For j = 1 To i - 1
        TEMP_MATRIX(i, j) = DATA_VECTOR(i, 1) * DATA_VECTOR(j, 1)
        TEMP_MATRIX(j, i) = TEMP_MATRIX(i, j)
    Next j
Next i
VECTOR_ELEMENTS_MATRIX_MULT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
VECTOR_ELEMENTS_MATRIX_MULT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_PRODUCT_SUM_FUNC
'DESCRIPTION   : Matrix scalar product
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_PRODUCT_SUM_FUNC(ByRef DATA_RNG As Variant, _
ByVal j As Long, _
ByVal k As Long)

Dim i As Long
Dim NSIZE As Long
Dim TEMP_SUM As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL
    
DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX)
TEMP_SUM = 0
For i = 1 To NSIZE
    TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j) * DATA_MATRIX(i, k)
Next i
MATRIX_ELEMENTS_PRODUCT_SUM_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_PRODUCT_SUM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_MULT_SCALAR_FUNC
'DESCRIPTION   : Matrix SCALAR_VAL multiplication
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_MULT_SCALAR_FUNC(ByRef DATA_RNG As Variant, _
ByVal SCALAR_VAL As Double)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(i, j) = SCALAR_VAL * DATA_MATRIX(i, j)
    Next j
Next i

MATRIX_ELEMENTS_MULT_SCALAR_FUNC = TEMP_MATRIX
Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_MULT_SCALAR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_MULT_FUNC
'DESCRIPTION   : Returns the M = aA x bB
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_MULT_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal SCALAR1_VAL As Double = 1, _
Optional ByVal SCALAR2_VAL As Double = 1)
  
Dim i As Long
Dim j As Long

Dim NROWS1 As Long
Dim NCOLUMNS1 As Long

Dim NROWS2 As Long
Dim NCOLUMNS2 As Long

Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = DATA1_RNG
DATA2_MATRIX = DATA2_RNG

NROWS1 = UBound(DATA1_MATRIX, 1)
NROWS2 = UBound(DATA2_MATRIX, 1)

NCOLUMNS1 = UBound(DATA1_MATRIX, 2)
NCOLUMNS2 = UBound(DATA2_MATRIX, 2)

'  If (NROWS1 <> NROWS2) Or (NCOLUMNS1 <> NCOLUMNS2) Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NROWS1, 1 To NCOLUMNS1)
For i = 1 To NROWS1
    For j = 1 To NCOLUMNS1
        TEMP_MATRIX(i, j) = (SCALAR1_VAL * DATA1_MATRIX(i, j)) * (SCALAR2_VAL * DATA2_MATRIX(i, j))
    Next j
Next i
MATRIX_ELEMENTS_MULT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_MULT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MMULT_FUNC
'DESCRIPTION   : Fast matrix multiplication without size limitation
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function MMULT_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal THRESHOLD As Single = 70)

Dim i As Single
Dim j As Single
Dim k As Single

Dim ii As Single
Dim jj As Single
Dim kk As Single

Dim iii As Single
Dim jjj As Single
Dim kkk As Single

Dim IMAX_VAL As Single
Dim IMIN_VAL As Single

Dim JMAX_VAL As Single
Dim JMIN_VAL As Single

Dim KMAX_VAL As Single
Dim KMIN_VAL As Single

Dim NSIZE As Single
Dim MSIZE As Single
Dim PSIZE As Single

Dim NROWS1 As Single
Dim NCOLUMNS1 As Single
Dim NCOLUMNS2 As Single

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant
Dim TEMP3_VECTOR As Variant

Dim TEMP_MATRIX As Variant

Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = DATA1_RNG
DATA2_MATRIX = DATA2_RNG
  
If UBound(DATA1_MATRIX, 2) <> UBound(DATA2_MATRIX, 1) Then: GoTo ERROR_LABEL
  
NROWS1 = UBound(DATA1_MATRIX, 1)  'rows of DATA1_MATRIX
NCOLUMNS1 = UBound(DATA1_MATRIX, 2)
  'columns of DATA1_MATRIX = rows of DATA2_MATRIX
NCOLUMNS2 = UBound(DATA2_MATRIX, 2)  'columns of DATA2_MATRIX

If NROWS1 <= THRESHOLD And NCOLUMNS2 <= THRESHOLD Then
    'fast multiplication
    MMULT_FUNC = MMULT2_FUNC(DATA1_MATRIX, DATA2_MATRIX)
    Exit Function
End If

'sub-matrix multiplication begins
NSIZE = Int(NROWS1 / THRESHOLD)   'row-blocks of DATA1_MATRIX
PSIZE = Int(NCOLUMNS1 / THRESHOLD)
'column-blocks of DATA1_MATRIX = row-blocks of DATA2_MATRIX
MSIZE = Int(NCOLUMNS2 / THRESHOLD)   'column-blocks of DATA2_MATRIX

If NSIZE * THRESHOLD < NROWS1 Then NSIZE = NSIZE + 1
If PSIZE * THRESHOLD < NCOLUMNS1 Then PSIZE = PSIZE + 1
If MSIZE * THRESHOLD < NCOLUMNS2 Then MSIZE = MSIZE + 1

ReDim TEMP_MATRIX(1 To NROWS1, 1 To NCOLUMNS2)

For ii = 1 To NSIZE
    For jj = 1 To MSIZE
        For kk = 1 To PSIZE
            'extract the sub-matrix DATA1_MATRIX(ii, kk) -> TEMP1_VECTOR
            IMIN_VAL = THRESHOLD * (ii - 1) + 1
            IMAX_VAL = THRESHOLD * ii
    
            If IMAX_VAL > NROWS1 Then IMAX_VAL = NROWS1
            KMIN_VAL = THRESHOLD * (kk - 1) + 1
            KMAX_VAL = THRESHOLD * kk
    
            If KMAX_VAL > NCOLUMNS1 Then KMAX_VAL = NCOLUMNS1
            iii = IMAX_VAL - IMIN_VAL + 1
            jjj = KMAX_VAL - KMIN_VAL + 1
            ReDim TEMP1_VECTOR(1 To iii, 1 To jjj)
            For i = 1 To UBound(TEMP1_VECTOR, 1)
                For k = 1 To UBound(TEMP1_VECTOR, 2)
                    TEMP1_VECTOR(i, k) = DATA1_MATRIX(i + IMIN_VAL - 1, k + KMIN_VAL - 1)
                Next k
            Next i
            'extract the sub-matrix DATA2_MATRIX(kk, jj) -> TEMP2_VECTOR
            KMIN_VAL = THRESHOLD * (kk - 1) + 1
            KMAX_VAL = THRESHOLD * kk
            If KMAX_VAL > NCOLUMNS1 Then KMAX_VAL = NCOLUMNS1
            JMIN_VAL = THRESHOLD * (jj - 1) + 1
            JMAX_VAL = THRESHOLD * jj
            If JMAX_VAL > NCOLUMNS2 Then JMAX_VAL = NCOLUMNS2
            jjj = KMAX_VAL - KMIN_VAL + 1
            kkk = JMAX_VAL - JMIN_VAL + 1
    
            ReDim TEMP2_VECTOR(1 To jjj, 1 To kkk)
            For k = 1 To UBound(TEMP2_VECTOR, 1)
                For j = 1 To UBound(TEMP2_VECTOR, 2)
                    TEMP2_VECTOR(k, j) = DATA2_MATRIX(k + KMIN_VAL - 1, j + JMIN_VAL - 1)
                Next j
            Next k
    
            'performs the multiplication of the sub-matrices
            TEMP3_VECTOR = MMULT2_FUNC(TEMP1_VECTOR, TEMP2_VECTOR)
            IMIN_VAL = THRESHOLD * (ii - 1) + 1
            JMIN_VAL = THRESHOLD * (jj - 1) + 1
    
            'accumulate the sub-matrix result
            For i = 1 To UBound(TEMP3_VECTOR, 1)
                For j = 1 To UBound(TEMP3_VECTOR, 2)
                    iii = i + IMIN_VAL - 1
                    jjj = j + JMIN_VAL - 1
                    TEMP_MATRIX(iii, jjj) = TEMP_MATRIX(iii, jjj) + TEMP3_VECTOR(i, j)
                Next j
            Next i
    
        Next kk
    Next jj
Next ii

MMULT_FUNC = TEMP_MATRIX
Exit Function
ERROR_LABEL:
MMULT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MMULT2_FUNC
'DESCRIPTION   : Returns the matrix product of two arrays. The result is an array
'with the same number of rows as array1 and the same number of columns as array.
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MMULT2_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS1 As Long
Dim NROWS2 As Long

Dim NCOLUMNS1 As Long
Dim NCOLUMNS2 As Long

Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = DATA1_RNG
DATA2_MATRIX = DATA2_RNG
  
NROWS1 = UBound(DATA1_MATRIX, 1)
NROWS2 = UBound(DATA2_MATRIX, 1)
  
NCOLUMNS1 = UBound(DATA1_MATRIX, 2)
NCOLUMNS2 = UBound(DATA2_MATRIX, 2)
  
If NCOLUMNS1 <> NROWS2 Then: GoTo ERROR_LABEL
    
ReDim TEMP_MATRIX(1 To NROWS1, 1 To NCOLUMNS2)

For i = 1 To NROWS1
    For j = 1 To NCOLUMNS2
        TEMP_MATRIX(i, j) = 0
        For k = 1 To NCOLUMNS1
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + DATA1_MATRIX(i, k) * DATA2_MATRIX(k, j)
        Next k
    Next j
Next i
    
MMULT2_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MMULT2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MMULT2_FUNC
'DESCRIPTION   : Returns matrix Z where Y=X*Z; inputs are matrix Y and X - must be
'of same no of rows this is equivalent to matlab code Z=X\Y

'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function IMULT_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim Y_VECTOR As Variant
Dim X_MATRIX As Variant
Dim XT_MATRIX As Variant

On Error GoTo ERROR_LABEL

Y_VECTOR = YDATA_RNG: X_MATRIX = XDATA_RNG
XT_MATRIX = MATRIX_TRANSPOSE_FUNC(X_MATRIX)
Select Case VERSION
Case 0
    IMULT_FUNC = MMULT_FUNC(MMULT_FUNC(MATRIX_INVERSE_FUNC(MMULT_FUNC(XT_MATRIX, X_MATRIX, 70), 2), XT_MATRIX, 70), Y_VECTOR, 70)
Case Else 'If UBOUND(Y,1) <> UBOUND(X,1)
    IMULT_FUNC = MMULT_FUNC(Y_VECTOR, MMULT_FUNC(MATRIX_INVERSE_FUNC(MMULT_FUNC(X_MATRIX, XT_MATRIX, 70), 0), X_MATRIX, 70), 70)
End Select

Exit Function
ERROR_LABEL:
IMULT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_VANDERMONDE_MATRIX_FUNC
'DESCRIPTION   : Returns the Vandermonde's matrix for a given vector:
'x = (x1, x2, ...xn)
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MULT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_VANDERMONDE_MATRIX_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)
NCOLUMNS = UBound(DATA_VECTOR, 2)
'If NROWS > 1 And NCOLUMNS > 1 Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = 1
    For j = 2 To NROWS
        TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j - 1) * DATA_VECTOR(i, 1)
    Next j
Next i

VECTOR_VANDERMONDE_MATRIX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
VECTOR_VANDERMONDE_MATRIX_FUNC = Err.number
End Function

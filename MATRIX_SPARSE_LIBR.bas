Attribute VB_Name = "MATRIX_SPARSE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_CONVERT_FUNC

'DESCRIPTION   : Converts a matrix in sparse coordinates format and vice versa.
'Input accepts both standard or sparse format

'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_CONVERT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim NSIZE As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

'--------------------------------------------------------------------------------
Select Case VERSION
'--------------------------------------------------------------------------------
Case 0 'RECT_MATRIX <- SPAR_MATRIX
'--------------------------------------------------------------------------------
    NSIZE = UBound(DATA_MATRIX, 1)
    'search for the max row and the max column
    NROWS = 0
    NCOLUMNS = 0
    For k = 1 To NSIZE
        If DATA_MATRIX(k, 1) > NROWS Then NROWS = DATA_MATRIX(k, 1)
        If DATA_MATRIX(k, 2) > NCOLUMNS Then NCOLUMNS = DATA_MATRIX(k, 2)
    Next k
    
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            TEMP_MATRIX(i, j) = 0
        Next j
    Next i
    For k = 1 To NSIZE
        i = DATA_MATRIX(k, 1)
        j = DATA_MATRIX(k, 2)
        TEMP_MATRIX(i, j) = DATA_MATRIX(k, 3)
    Next k
    MATRIX_SPARSE_CONVERT_FUNC = TEMP_MATRIX
'--------------------------------------------------------------------------------
Case Else 'RECT_MATRIX -> SPAR_MATRIX
'--------------------------------------------------------------------------------
    NROWS = UBound(DATA_MATRIX, 1)
    NCOLUMNS = UBound(DATA_MATRIX, 2)
    'search for the max number of not empty cells
    NSIZE = 0
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            If Abs(DATA_MATRIX(i, j)) > 0 Then NSIZE = NSIZE + 1
        Next j
    Next i
    ReDim TEMP_MATRIX(1 To NSIZE, 1 To 3)
    k = 0
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            If Abs(DATA_MATRIX(i, j)) > 0 Then
                k = k + 1
                TEMP_MATRIX(k, 1) = i
                TEMP_MATRIX(k, 2) = j
                TEMP_MATRIX(k, 3) = DATA_MATRIX(i, j)
            End If
        Next j
    Next i
    
    MATRIX_SPARSE_CONVERT_FUNC = TEMP_MATRIX
'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_CONVERT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_DIMENSION_FUNC

'DESCRIPTION   : Returns the true dimensions (n x m) and the filling
'factor of a sparse matrix

'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_DIMENSION_FUNC(ByRef DATA_RNG As Variant)

Dim j As Long
Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX, 1)

NROWS = 0
NCOLUMNS = 0

For j = 1 To NSIZE
    If DATA_MATRIX(j, 1) > NROWS Then NROWS = DATA_MATRIX(j, 1)
    If DATA_MATRIX(j, 2) > NCOLUMNS Then NCOLUMNS = DATA_MATRIX(j, 2)
Next j

ReDim TEMP_VECTOR(1 To 1, 1 To 2)
TEMP_VECTOR(1, 1) = NROWS
TEMP_VECTOR(1, 2) = NCOLUMNS
    
MATRIX_SPARSE_DIMENSION_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_DIMENSION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_MULT_VECTOR_ELEMENTS_FUNC

'DESCRIPTION   : Returns the product of a sparse matrix A for a vector b.
'performs A*x = y, where A is in sparse format
'A = (k x 3) equivalent to a (n x m) matrix; x = vector (m)

'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_MULT_VECTOR_ELEMENTS_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant, _
Optional ByVal NROWS As Long = 0)

'Input DATA_RNG is in sparse format

Dim i As Long
Dim j As Long
Dim k As Long
Dim NSIZE As Long
'Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant

Dim DIMEN_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_RNG
DATA_VECTOR = VECTOR_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

DIMEN_VECTOR = MATRIX_SPARSE_DIMENSION_FUNC(DATA_MATRIX)
NCOLUMNS = DIMEN_VECTOR(1, 2)
If NCOLUMNS > UBound(DATA_VECTOR, 1) Then: GoTo ERROR_LABEL

If NROWS = 0 Then: NROWS = DIMEN_VECTOR(1, 1)
If NROWS < DIMEN_VECTOR(1, 1) Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NROWS, 1 To 1)

NSIZE = UBound(DATA_MATRIX, 1)
For k = 1 To NSIZE 'multiplication begins
    
    i = DATA_MATRIX(k, 1)
    j = DATA_MATRIX(k, 2)
    
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 1) + _
                        DATA_MATRIX(k, 3) * DATA_VECTOR(j, 1)
Next k

MATRIX_SPARSE_MULT_VECTOR_ELEMENTS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_MULT_VECTOR_ELEMENTS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_GET_ROW_FUNC
'DESCRIPTION   : Extract a row in a sparse matrix
'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_SPARSE_GET_ROW_FUNC(ByRef DATA_RNG As Variant, _
ByVal AROW As Long, _
Optional ByVal SROW As Long = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim TEMP_FLAG As Boolean

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim DIMEN_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
DIMEN_VECTOR = MATRIX_SPARSE_DIMENSION_FUNC(DATA_MATRIX)
If AROW > DIMEN_VECTOR(1, 1) Then: GoTo ERROR_LABEL

k = DIMEN_VECTOR(1, 2)
ReDim TEMP_VECTOR(1 To 1, 1 To k)

NROWS = UBound(DATA_MATRIX, 1)
If SROW < 1 Then SROW = 1
    
TEMP_FLAG = False
    
For j = SROW To NROWS
    If DATA_MATRIX(j, 1) = AROW Then
        i = DATA_MATRIX(j, 2)
        TEMP_VECTOR(1, i) = DATA_MATRIX(j, 3)
        TEMP_FLAG = True
    Else
        If TEMP_FLAG Then Exit For
    End If
    TEMP_FLAG = False
Next j

MATRIX_SPARSE_GET_ROW_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_GET_ROW_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_GET_COLUMN_FUNC
'DESCRIPTION   : Extract a column in a sparse matrix
'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_GET_COLUMN_FUNC(ByRef DATA_RNG As Variant, _
ByVal ACOLUMN As Long, _
Optional ByVal SROW As Long = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim TEMP_FLAG As Boolean

Dim TEMP_VECTOR As Variant
Dim DIMEN_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
DIMEN_VECTOR = MATRIX_SPARSE_DIMENSION_FUNC(DATA_MATRIX)
If ACOLUMN > DIMEN_VECTOR(1, 2) Then: GoTo ERROR_LABEL

k = DIMEN_VECTOR(1, 1)
ReDim TEMP_VECTOR(1 To k, 1 To 1)

NROWS = UBound(DATA_MATRIX, 1)
If SROW < 1 Then SROW = 1

TEMP_FLAG = False
For j = SROW To NROWS
    If DATA_MATRIX(j, 2) = ACOLUMN Then
        i = DATA_MATRIX(j, 1)
        TEMP_VECTOR(i, 1) = DATA_MATRIX(j, 3)
        TEMP_FLAG = True
    Else
        If TEMP_FLAG Then Exit For
    End If
    TEMP_FLAG = False
Next j

MATRIX_SPARSE_GET_COLUMN_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_GET_COLUMN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_ADD_ELEMENTS_FUNC

'DESCRIPTION   : Returns the addition of two sparse matrices. If SCALAR = -1 Then
'it returns the subtraction of two sparse matrices. Input A and B
'are in sparse format.

'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_ADD_ELEMENTS_FUNC(ByRef ADATA_RNG As Variant, _
ByRef BDATA_RNG As Variant, _
Optional ByVal SCALAR As Double = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Double

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim DIMEN_VECTOR As Variant

Dim ADATA_MATRIX As Variant
Dim BDATA_MATRIX As Variant

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ADATA_MATRIX = ADATA_RNG
BDATA_MATRIX = BDATA_RNG

If UBound(ADATA_MATRIX, 1) <> UBound(BDATA_MATRIX, 1) Then: GoTo ERROR_LABEL
If UBound(ADATA_MATRIX, 2) <> UBound(BDATA_MATRIX, 2) Then: GoTo ERROR_LABEL

DIMEN_VECTOR = MATRIX_SPARSE_DIMENSION_FUNC(ADATA_MATRIX)

NROWS = DIMEN_VECTOR(1, 1)
NCOLUMNS = DIMEN_VECTOR(1, 2)

ReDim ATEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
ReDim BTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)

NSIZE = UBound(ADATA_MATRIX, 1) + UBound(BDATA_MATRIX, 1) 'max dimension
ReDim BTEMP_MATRIX(1 To NSIZE, 1 To 3)

k = 0
For i = 1 To NROWS
    ATEMP_VECTOR = MATRIX_SPARSE_GET_ROW_FUNC(ADATA_MATRIX, i)
    BTEMP_VECTOR = MATRIX_SPARSE_GET_ROW_FUNC(BDATA_MATRIX, i)
    For j = 1 To NCOLUMNS
        TEMP_VAL = ATEMP_VECTOR(1, j) + SCALAR * BTEMP_VECTOR(1, j)
'        If TEMP_VAL <> 0 Then
            k = k + 1
            BTEMP_MATRIX(k, 1) = i
            BTEMP_MATRIX(k, 2) = j
            BTEMP_MATRIX(k, 3) = TEMP_VAL
'        End If
    Next j
Next i

NSIZE = k 'reload out matrix
ReDim ATEMP_MATRIX(1 To NSIZE, 1 To 3)
For k = 1 To NSIZE
    ATEMP_MATRIX(k, 1) = BTEMP_MATRIX(k, 1)
    ATEMP_MATRIX(k, 2) = BTEMP_MATRIX(k, 2)
    ATEMP_MATRIX(k, 3) = BTEMP_MATRIX(k, 3)
Next k

MATRIX_SPARSE_ADD_ELEMENTS_FUNC = ATEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_ADD_ELEMENTS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_MMULT_FUNC

'DESCRIPTION   : Returns the matrix product of two sparse matrix. The result
'is a sparse matrix with the same number of rows as ADATA_MATRIX and the
'same number of columns BDATA_MATRIX.

'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_MMULT_FUNC(ByRef ADATA_RNG As Variant, _
ByRef BDATA_RNG As Variant)

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ATEMP_MATRIX = MATRIX_SPARSE_CONVERT_FUNC(ADATA_RNG, 0)
BTEMP_MATRIX = MATRIX_SPARSE_CONVERT_FUNC(BDATA_RNG, 0)
CTEMP_MATRIX = MMULT_FUNC(ATEMP_MATRIX, BTEMP_MATRIX, 70)
CTEMP_MATRIX = MATRIX_SPARSE_CONVERT_FUNC(CTEMP_MATRIX, 1)

MATRIX_SPARSE_MMULT_FUNC = CTEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_MMULT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_TRANSPOSE_FUNC
'DESCRIPTION   : Returns the transposed of a sparse matrix
'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_TRANSPOSE_FUNC(ByRef DATA_RNG As Variant)
'RECT_MATRIX

Dim i As Long
Dim j As Long
Dim NSIZE As Long
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NSIZE = UBound(DATA_MATRIX, 1)
For j = 1 To NSIZE
    i = DATA_MATRIX(j, 1)
    DATA_MATRIX(j, 1) = DATA_MATRIX(j, 2)
    DATA_MATRIX(j, 2) = i
Next j

MATRIX_SPARSE_TRANSPOSE_FUNC = MATRIX_DOUBLE_SORT_FUNC(DATA_MATRIX)

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_TRANSPOSE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_DOMINANCE_FUNC

'DESCRIPTION   : Compute the dominance factor for a square sparse matrix (n x n)
' df = |aii|/sum(|aij|),   0 <= df <= 1

'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_SPARSE_DOMINANCE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim TEMP_MIN As Double
Dim TEMP_MAX As Double
Dim TEMP_MEAN As Double

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim DIMEN_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX, 1)
DIMEN_VECTOR = MATRIX_SPARSE_DIMENSION_FUNC(DATA_MATRIX)

NROWS = DIMEN_VECTOR(1, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim DIAG_VECTOR(1 To NROWS, 1 To 1)

For k = 1 To NSIZE
    i = DATA_MATRIX(k, 1)
    TEMP_VECTOR(i, 1) = TEMP_VECTOR(i, 1) + Abs(DATA_MATRIX(k, 3))
    If i = DATA_MATRIX(k, 2) Then DIAG_VECTOR(i, 1) = DATA_MATRIX(k, 3) 'diagonal
Next k

TEMP_MEAN = 0
TEMP_MAX = 0
TEMP_MIN = 10 ^ 300
For i = 1 To NROWS
    TEMP_VECTOR(i, 1) = Abs(DIAG_VECTOR(i, 1)) / TEMP_VECTOR(i, 1)
    TEMP_MEAN = TEMP_MEAN + TEMP_VECTOR(i, 1)
    If TEMP_VECTOR(i, 1) < TEMP_MIN Then TEMP_MIN = TEMP_VECTOR(i, 1)
    If TEMP_VECTOR(i, 1) > TEMP_MAX Then TEMP_MAX = TEMP_VECTOR(i, 1)
Next i

Select Case OUTPUT
    Case 0
        MATRIX_SPARSE_DOMINANCE_FUNC = TEMP_MEAN / NROWS
    Case 1
        MATRIX_SPARSE_DOMINANCE_FUNC = TEMP_VECTOR
    Case Else
        MATRIX_SPARSE_DOMINANCE_FUNC = Array(TEMP_MEAN / NROWS, TEMP_VECTOR)
End Select

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_DOMINANCE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_SWAP_ROW_FUNC
'DESCRIPTION   : Swap Row in a Sparse matrix
'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_SPARSE_SWAP_ROW_FUNC(ByRef DATA_RNG As Variant, _
ByVal j As Long, _
ByVal i As Long, _
Optional ByVal VERSION As Integer = 0)

Dim k As Long
Dim TEMP_VAL As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

Select Case VERSION
Case 0
    For k = LBound(DATA_MATRIX, 1) To UBound(DATA_MATRIX, 1)
        If DATA_MATRIX(k, 1) = i Then
            DATA_MATRIX(k, 1) = j
        ElseIf DATA_MATRIX(k, 1) = j Then
            DATA_MATRIX(k, 1) = i
        End If
    Next k
Case Else
    For k = 1 To 2
        TEMP_VAL = DATA_MATRIX(i, k)
        DATA_MATRIX(i, k) = DATA_MATRIX(j, k)
        DATA_MATRIX(j, k) = TEMP_VAL
    Next k
End Select

MATRIX_SPARSE_SWAP_ROW_FUNC = DATA_MATRIX
    
Exit Function
ERROR_LABEL:
MATRIX_SPARSE_SWAP_ROW_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_SEARCH_COLUMN_FUNC
'DESCRIPTION   : Search for column in a sparse matrix using the binary algorithm
'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_SEARCH_COLUMN_FUNC(ByRef DATA_RNG As Variant, _
ByVal ACOLUMN As Long, _
ByVal MIN_VAL As Long, _
ByVal MAX_VAL As Long)

Dim i As Long
Dim j As Long

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

i = 0
Do
    j = Int((MIN_VAL + MAX_VAL) / 2)
    If DATA_MATRIX(j, 2) > ACOLUMN Then
        MAX_VAL = j - 1
    ElseIf DATA_MATRIX(j, 2) < ACOLUMN Then
        MIN_VAL = j + 1
    Else
        i = j
        Exit Do
    End If
Loop Until MIN_VAL + 1 >= MAX_VAL

If i = 0 Then
    If DATA_MATRIX(MAX_VAL, 2) = ACOLUMN Then
        i = MAX_VAL
    ElseIf DATA_MATRIX(MIN_VAL, 2) = ACOLUMN Then
        i = MIN_VAL
    End If
End If

MATRIX_SPARSE_SEARCH_COLUMN_FUNC = i

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_SEARCH_COLUMN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_DOMINANCE_TRANSFORMATION_FUNC

'DESCRIPTION   : For sparse matrices, this algorithm addensate the biggest
'values around the first diagonal in order to make diagonal
'dominant the system matrix A*x = B. It also tries to improve
'the dominance of a sparse matrix by rows exchanging.

'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_DOMINANCE_TRANSFORMATION_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim hh As Double
Dim ii As Double
Dim jj As Double
Dim kk As Double
Dim ll As Double

Dim NSIZE As Long
Dim nLOOPS As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DIAG_VECTOR As Variant
Dim INDEX_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

Dim DATA_VECTOR As Variant
Dim DIMEN_VECTOR As Variant

Dim TEMP_FLAG As Boolean

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_RNG
DATA_VECTOR = VECTOR_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

NSIZE = UBound(DATA_MATRIX, 1)
DIMEN_VECTOR = MATRIX_SPARSE_DIMENSION_FUNC(DATA_MATRIX)

NROWS = DIMEN_VECTOR(1, 1)
NCOLUMNS = DIMEN_VECTOR(1, 2)

ReDim TEMP_MATRIX(1 To NSIZE, 1 To 3)
For k = 1 To NSIZE 'load the absolute auxiliary matrix |aij|
    TEMP_MATRIX(k, 1) = DATA_MATRIX(k, 1)
    TEMP_MATRIX(k, 2) = DATA_MATRIX(k, 2)
    TEMP_MATRIX(k, 3) = Abs(DATA_MATRIX(k, 3))
Next k

TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 3, 0)
DIAG_VECTOR = MATRIX_SPARSE_DIAGONAL_FUNC(DATA_MATRIX)

ReDim INDEX_VECTOR(1 To NROWS, 1 To 2) 'create the row-index for a sparse matrix

INDEX_VECTOR(DATA_MATRIX(1, 1), 1) = 1
For k = 2 To UBound(DATA_MATRIX, 1)
    'i = DATA_MATRIX(k, 1)
    If DATA_MATRIX(k - 1, 1) <> DATA_MATRIX(k, 1) Then
        INDEX_VECTOR(DATA_MATRIX(k, 1), 1) = k
        INDEX_VECTOR(DATA_MATRIX(k - 1, 1), 2) = k - 1
    End If
Next k
INDEX_VECTOR(DATA_MATRIX(UBound(DATA_MATRIX, 1), 1), 2) = UBound(DATA_MATRIX, 1)
nLOOPS = 0
Do
    TEMP_FLAG = False
    k = 0
    Do
        k = k + 1
        i = TEMP_MATRIX(k, 1)
        j = TEMP_MATRIX(k, 2)
        kk = TEMP_MATRIX(k, 3)
        ii = DIAG_VECTOR(i, 1)
        jj = DIAG_VECTOR(j, 1)
        'get the symmetric element
        hh = 0

        h = MATRIX_SPARSE_SEARCH_COLUMN_FUNC(DATA_MATRIX, i, INDEX_VECTOR(j, 1), INDEX_VECTOR(j, 2))
        
        If h <> 0 Then hh = DATA_MATRIX(h, 3)

        ll = Abs(kk) * Abs(hh) - Abs(ii) * Abs(jj)
        If ll <= 0 Then
            ll = Abs(kk) + Abs(hh) - Abs(ii) - Abs(jj)
        End If
        If ll > 0 Then 'swap the rows i, j
            DATA_VECTOR = MATRIX_SWAP_ROW_FUNC(DATA_VECTOR, j, i)
            DATA_MATRIX = MATRIX_SPARSE_SWAP_ROW_FUNC(DATA_MATRIX, j, i, 0)
            INDEX_VECTOR = MATRIX_SPARSE_SWAP_ROW_FUNC(INDEX_VECTOR, j, i, 1)
            TEMP_MATRIX = MATRIX_SPARSE_SWAP_ROW_FUNC(TEMP_MATRIX, j, i, 0)
            
            DIAG_VECTOR(i, 1) = hh
            DIAG_VECTOR(j, 1) = kk
            nLOOPS = nLOOPS + 1
            TEMP_FLAG = True
        End If
    Loop Until k = NSIZE
Loop Until TEMP_FLAG = False Or nLOOPS > NROWS

DATA_VECTOR = MATRIX_DOUBLE_SORT_FUNC(DATA_VECTOR)
DATA_MATRIX = MATRIX_DOUBLE_SORT_FUNC(DATA_MATRIX)

Select Case OUTPUT
    Case 0
        MATRIX_SPARSE_DOMINANCE_TRANSFORMATION_FUNC = DATA_MATRIX
    Case 1
        MATRIX_SPARSE_DOMINANCE_TRANSFORMATION_FUNC = DATA_VECTOR
    Case Else
        MATRIX_SPARSE_DOMINANCE_TRANSFORMATION_FUNC = Array(DATA_MATRIX, DATA_VECTOR)
End Select

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_DOMINANCE_TRANSFORMATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_DIAGONAL_FUNC
'DESCRIPTION   : Extract the first diagonal from a sparse matrix
'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_DIAGONAL_FUNC(ByRef DATA_RNG As Variant)

Dim k As Long
Dim NROWS As Long
Dim NSIZE As Long

Dim DIMEN_VECTOR As Variant
Dim DIAG_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
DIMEN_VECTOR = MATRIX_SPARSE_DIMENSION_FUNC(DATA_MATRIX)

NROWS = DIMEN_VECTOR(1, 1)

ReDim DIAG_VECTOR(1 To NROWS, 1 To 1)
NSIZE = UBound(DATA_MATRIX, 1)
For k = 1 To NSIZE
    If DATA_MATRIX(k, 1) = DATA_MATRIX(k, 2) Then
        DIAG_VECTOR(DATA_MATRIX(k, 1), 1) = DATA_MATRIX(k, 3)
    End If
Next k

MATRIX_SPARSE_DIAGONAL_FUNC = DIAG_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_DIAGONAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_CHOLESKY_FUNC
'DESCRIPTION   : Performs the Cholesky decomposition of a sparse matrix
'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_CHOLESKY_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NSIZE As Long
Dim MSIZE As Long

Dim NROWS As Long

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant
Dim DTEMP_VECTOR As Variant
Dim ETEMP_VECTOR As Variant

Dim DATA_MATRIX As Variant
Dim DIMEN_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX, 1)
MSIZE = NSIZE

DIMEN_VECTOR = MATRIX_SPARSE_DIMENSION_FUNC(DATA_MATRIX)
NROWS = DIMEN_VECTOR(1, 1)

jj = 2 * NSIZE

ReDim ATEMP_ARR(1 To jj, 1 To 3)
ReDim ATEMP_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    ReDim BTEMP_VECTOR(1 To NROWS, 1 To 1)
    ReDim CTEMP_VECTOR(1 To 1, 1 To NROWS)
    'extract the i-row from DATA_MATRIX
    CTEMP_VECTOR = MATRIX_SPARSE_GET_ROW_FUNC(DATA_MATRIX, i, 1)
    
    BTEMP_VECTOR(i, 1) = CTEMP_VECTOR(1, i) - ATEMP_VECTOR(i, 1)
    If BTEMP_VECTOR(i, 1) <= 0 Then
        MATRIX_SPARSE_CHOLESKY_FUNC = 1
        Exit Function 'decomposition fails
    End If
    BTEMP_VECTOR(i, 1) = Sqr(BTEMP_VECTOR(i, 1))
    
    ReDim DTEMP_VECTOR(1 To NROWS, 1 To 1)
    ReDim ETEMP_VECTOR(1 To 1, 1 To NROWS)
    
    kk = 0
    For k = 1 To i - 1
        ETEMP_VECTOR = MATRIX_SPARSE_GET_ROW_FUNC(ATEMP_ARR, k, kk)
        For j = i + 1 To NROWS
            DTEMP_VECTOR(j, 1) = DTEMP_VECTOR(j, 1) + _
                ETEMP_VECTOR(1, j) * ETEMP_VECTOR(1, i)
        Next j
    Next k
    
    For j = i + 1 To NROWS
        BTEMP_VECTOR(j, 1) = (CTEMP_VECTOR(1, j) - _
            DTEMP_VECTOR(j, 1)) / BTEMP_VECTOR(i, 1)
        ATEMP_VECTOR(j, 1) = ATEMP_VECTOR(j, 1) + _
            BTEMP_VECTOR(j, 1) ^ 2
    Next j
    
    'the i-row of BTEMP_ARR is computed. Save the row into the matrix
    For j = 1 To NROWS
        If BTEMP_VECTOR(j, 1) <> 0 Then
            ii = ii + 1
            If ii > jj Then
                jj = UBound(ATEMP_ARR, 1) + MSIZE
                ATEMP_ARR = MATRIX_REDIM_FUNC(ATEMP_ARR, MSIZE, 0)
            End If
            ATEMP_ARR(ii, 1) = i
            ATEMP_ARR(ii, 2) = j
            ATEMP_ARR(ii, 3) = BTEMP_VECTOR(j, 1)
        End If
    Next j
Next i

'load the output matrix
ReDim BTEMP_ARR(1 To ii, 1 To 3)

For k = 1 To ii
    BTEMP_ARR(k, 1) = ATEMP_ARR(k, 1)
    BTEMP_ARR(k, 2) = ATEMP_ARR(k, 2)
    BTEMP_ARR(k, 3) = ATEMP_ARR(k, 3)
Next k

MATRIX_SPARSE_CHOLESKY_FUNC = MATRIX_SPARSE_TRANSPOSE_FUNC(BTEMP_ARR)

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_CHOLESKY_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_GJ_LINEAR_SYSTEM_FUNC

'DESCRIPTION   : Solves a sparse linear system using the iterative Gauss
'algorithm with partial pivot and back-substitution. Input
'A accepts both standard or sparse format.

'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_GJ_LINEAR_SYSTEM_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant, _
Optional ByVal epsilon As Double = 10 ^ -11, _
Optional ByVal OUTPUT As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim iii As Long
Dim jjj As Long
Dim kkk As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ANSIZE As Long
Dim BNSIZE As Long
Dim CNSIZE As Long
Dim DNSIZE As Long

Dim TEMP_DET As Double
Dim TEMP_MULT As Double
Dim TEMP_VAL As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim ATEMP_ABS As Double
Dim BTEMP_ABS As Double

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim DATA_VECTOR As Variant
Dim INDEX_VECTOR As Variant
Dim DIMEN_VECTOR As Variant
Dim DATA_MATRIX As Variant

Dim NULL_FLAG As Boolean
'Dim SWAP_FLAG As Boolean

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_RNG
DATA_VECTOR = VECTOR_RNG

TEMP_MULT = 1
CNSIZE = UBound(DATA_MATRIX)
DIMEN_VECTOR = MATRIX_SPARSE_DIMENSION_FUNC(DATA_MATRIX)
NROWS = DIMEN_VECTOR(1, 1)
NCOLUMNS = DIMEN_VECTOR(1, 2)

ANSIZE = Int(2 * CNSIZE)
BNSIZE = Int(2 * CNSIZE)

ReDim ATEMP_ARR(1 To ANSIZE, 1 To 3) 'contains the final matrix
DNSIZE = ANSIZE
kk = 0 'Triangolarized matrix COUNTER
'ii = 1 'pivot
'SWAP_FLAG = False
For ii = 1 To NROWS

    kkk = UBound(DATA_MATRIX)
    ReDim INDEX_VECTOR(1 To NROWS, 1 To 2)
    INDEX_VECTOR(DATA_MATRIX(1, 1), 1) = 1
    For jjj = 2 To kkk
        iii = DATA_MATRIX(jjj, 1)
        If DATA_MATRIX(jjj - 1, 1) <> DATA_MATRIX(jjj, 1) Then
            INDEX_VECTOR(DATA_MATRIX(jjj, 1), 1) = jjj
            INDEX_VECTOR(DATA_MATRIX(jjj - 1, 1), 2) = jjj - 1
        End If
    Next jjj
    INDEX_VECTOR(DATA_MATRIX(kkk, 1), 2) = kkk
    
    k = 1  'original matrix COUNTER
    ReDim BTEMP_ARR(1 To BNSIZE, 1 To 3)  'contains the reduce matrix
    ReDim ATEMP_VECTOR(1 To 1, 1 To NROWS) 'contains the pivot-row
    'load the pivot-row
    i = ii
    Do
        k = INDEX_VECTOR(i, 1)
        ATEMP_VECTOR = MATRIX_SPARSE_GET_ROW_FUNC(DATA_MATRIX, i, k)
        ATEMP_SUM = 0
        For h = 1 To UBound(ATEMP_VECTOR, 2)
            ATEMP_SUM = ATEMP_SUM + (ATEMP_VECTOR(1, h)) ^ 2
        Next h
        ATEMP_ABS = Sqr(ATEMP_SUM)
        BTEMP_ABS = Abs(ATEMP_VECTOR(1, ii))
        If BTEMP_ABS > 10 ^ -8 Then BTEMP_ABS = BTEMP_ABS / ATEMP_ABS
        If BTEMP_ABS > epsilon Then
            If i <> ii Then  'swap row ii <-> i
                DATA_MATRIX = MATRIX_SPARSE_SWAP_ROW_FUNC(DATA_MATRIX, i, ii, 0)
                INDEX_VECTOR = MATRIX_SPARSE_SWAP_ROW_FUNC(INDEX_VECTOR, i, ii, 1)
                DATA_VECTOR = MATRIX_SWAP_ROW_FUNC(DATA_VECTOR, i, ii)
            End If
            Exit Do
        End If
        i = i + 1
    Loop Until i > NROWS
    If i > NROWS Then
        NULL_FLAG = True  'matrix singular
        Exit For
    End If

    TEMP_VAL = ATEMP_VECTOR(1, ii)
    TEMP_MULT = TEMP_MULT * TEMP_VAL
    For j = ii To NROWS
        ATEMP_VECTOR(1, j) = ATEMP_VECTOR(1, j) / TEMP_VAL
    Next j
    DATA_VECTOR(ii, 1) = DATA_VECTOR(ii, 1) / TEMP_VAL
    'store into triangolarized matrix
    For j = ii To NROWS
        If ATEMP_VECTOR(1, j) <> 0 Then
            kk = kk + 1
            If kk > ANSIZE Then
                ANSIZE = UBound(ATEMP_ARR, 1) + DNSIZE
                ATEMP_ARR = MATRIX_REDIM_FUNC(ATEMP_ARR, DNSIZE, 0)
            End If
            ATEMP_ARR(kk, 3) = ATEMP_VECTOR(1, j)
            ATEMP_ARR(kk, 2) = j
            ATEMP_ARR(kk, 1) = ii
        End If
    Next j
    
    jj = 0  'reduce matrix COUNTER
    'reduce all rows under the pivot-row
    For i = ii + 1 To NROWS
        'load the i-row
        ReDim BTEMP_VECTOR(1 To 1, 1 To NROWS) 'contains the i-row
        k = INDEX_VECTOR(i, 1)
        BTEMP_VECTOR = MATRIX_SPARSE_GET_ROW_FUNC(DATA_MATRIX, i, k)

        TEMP_VAL = BTEMP_VECTOR(1, ii)
        If TEMP_VAL <> 0 Then
            For j = ii To NROWS
                BTEMP_VECTOR(1, j) = BTEMP_VECTOR(1, j) - TEMP_VAL _
                    * ATEMP_VECTOR(1, j)
                If Abs(BTEMP_VECTOR(1, j)) < epsilon Then _
                    BTEMP_VECTOR(1, j) = 0 'mop-up
            Next j
            BTEMP_VECTOR(1, ii) = 0
        End If
        DATA_VECTOR(i, 1) = DATA_VECTOR(i, 1) - TEMP_VAL * DATA_VECTOR(ii, 1)
        'store into BTEMP_ARR
        NULL_FLAG = True
        For j = ii To NROWS
            If BTEMP_VECTOR(1, j) <> 0 Then
                NULL_FLAG = False
                jj = jj + 1
                If jj > BNSIZE Then
                    BNSIZE = UBound(BTEMP_ARR, 1) + DNSIZE
                    BTEMP_ARR = MATRIX_REDIM_FUNC(BTEMP_ARR, DNSIZE, 0)
                End If
                BTEMP_ARR(jj, 3) = BTEMP_VECTOR(1, j)
                BTEMP_ARR(jj, 2) = j
                BTEMP_ARR(jj, 1) = i
            End If
        Next j
    Next i
    
    If NULL_FLAG Then Exit For  'singular matrix
        
    'reload BTEMP_ARR --> DATA_MATRIX
    CNSIZE = jj
    If CNSIZE > 0 Then
        ReDim DATA_MATRIX(1 To CNSIZE, 1 To 3)
        For k = 1 To CNSIZE
            DATA_MATRIX(k, 1) = BTEMP_ARR(k, 1)
            DATA_MATRIX(k, 2) = BTEMP_ARR(k, 2)
            DATA_MATRIX(k, 3) = BTEMP_ARR(k, 3)
        Next k
        BNSIZE = Int(2 * CNSIZE)
    End If
    'come back
Next ii

If NULL_FLAG Then
    TEMP_DET = 0
    MATRIX_SPARSE_GJ_LINEAR_SYSTEM_FUNC = TEMP_DET
    Exit Function
End If

ReDim DATA_MATRIX(1 To kk, 1 To 3)

For k = 1 To kk
    DATA_MATRIX(k, 1) = ATEMP_ARR(k, 1)
    DATA_MATRIX(k, 2) = ATEMP_ARR(k, 2)
    DATA_MATRIX(k, 3) = ATEMP_ARR(k, 3)
Next k
Erase ATEMP_ARR, BTEMP_ARR
'DATA_MATRIX is DATA_MATRIX triangolarized matrix with diagonal = 1
'back-substitution begins
CNSIZE = UBound(DATA_MATRIX)

ReDim XTEMP_VECTOR(1 To NROWS, 1 To 1)
BTEMP_SUM = 0
For k = CNSIZE To 1 Step -1
    i = DATA_MATRIX(k, 1)
    j = DATA_MATRIX(k, 2)
    If i = j Then
        XTEMP_VECTOR(i, 1) = DATA_VECTOR(i, 1) - BTEMP_SUM
        BTEMP_SUM = 0
    Else
        BTEMP_SUM = BTEMP_SUM + DATA_MATRIX(k, 3) * XTEMP_VECTOR(j, 1)
    End If
Next k
TEMP_DET = TEMP_MULT

Select Case OUTPUT
    Case 0
        MATRIX_SPARSE_GJ_LINEAR_SYSTEM_FUNC = XTEMP_VECTOR
    Case 1
        MATRIX_SPARSE_GJ_LINEAR_SYSTEM_FUNC = DATA_MATRIX
    Case 2
        MATRIX_SPARSE_GJ_LINEAR_SYSTEM_FUNC = TEMP_DET
    Case 3
        MATRIX_SPARSE_GJ_LINEAR_SYSTEM_FUNC = DATA_VECTOR
    Case Else
        MATRIX_SPARSE_GJ_LINEAR_SYSTEM_FUNC = Array(XTEMP_VECTOR, DATA_MATRIX, _
                            TEMP_DET, DATA_VECTOR)
End Select

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_GJ_LINEAR_SYSTEM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SPARSE_ADSOR_SOLVER_FUNC

'DESCRIPTION   : Adaptive-SOR algorithm for sparse linear system
'solving (Ax = b (ADSOR); Solves a sparse linear system using the iterative
'ADSOR algorithm (Adaptive-SOR). Input A accepts both
'standard or sparse format. The Iterative algorithm always
'converges if the matrix is diagonal dominant (Di > 0.5 for each i).
'It is usually faster than other deterministic methods (Gauss. LR, LL).
'The max iterations (default 400) can be modified by the top-right field

'LIBRARY       : MATRIX
'GROUP         : SPARSE
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SPARSE_ADSOR_SOLVER_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant, _
Optional ByVal epsilon As Double = 2 * 10 ^ -15, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal VERSION As Integer = 3, _
Optional ByVal omega As Double = 1)

'  Adaptive-SOR algorithm for sparse linear system solving
'  This program solves an inhomogeneous linear system AX = B of
'  equations with a nonsingular system matrix A. The method of
'  Gauss-Seidel is used jointly with relaxation parameter OMEGA
'  adjusted during the iteration.
'  Under certain conditions the acceleration delta^2 di Aitken
'  is performed to improve the convergence

'  A      : 2-dimensional array A, containing the
'           system matrix for the linear equations
'  B      : the right hand side of the system

'  epsilon    : desired accuracy; the iteration is stopped when the
'           norm-2 of the relative error does not exceed  epsilon

'  nLOOPS  : Maximal number of iterations allowed

'  VERSION  : determines which the method used:
'           = 0, adaptive SOR method
'           = 1, SOR method for a given relaxation parameter
'           = 2, Gauﬂ-Seidel method
'           = 3, adaptive SOR method + extrapolation

'  OMEGA  : in case Imeth=1, the optimal relaxation parameter


Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim iii As Long
Dim kkk As Long
Dim jjj As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_MULT As Double
Dim BTEMP_MULT As Double
Dim CTEMP_MULT As Double

Dim TEMP_ABS As Double
Dim TEMP_ERR As Double
Dim TEMP_DIV As Double
Dim TEMP_SUM As Double
Dim TEMP_BETA As Double
Dim TEMP_ALPHA As Double
Dim TEMP_MEAN As Double
Dim TEMP_NORM As Double
Dim TEMP_RESID As Double
Dim TEMP_SIGMA As Double

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim CTEMP_ARR As Variant
Dim DTEMP_ARR As Variant

Dim COEF_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim DATA_VECTOR As Variant
Dim DIMEN_VECTOR As Variant

Dim ERROR_STR As String
Dim SWITCH_FLAG As Boolean

On Error GoTo ERROR_LABEL

ERROR_STR = 0
DATA_MATRIX = MATRIX_RNG
DATA_VECTOR = VECTOR_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

If VERSION = 1 Then ATEMP_MULT = omega Else ATEMP_MULT = 1

DIMEN_VECTOR = MATRIX_SPARSE_DIMENSION_FUNC(DATA_MATRIX)
NROWS = DIMEN_VECTOR(1, 1)
NCOLUMNS = DIMEN_VECTOR(1, 2)

If NROWS <> NCOLUMNS Or NROWS <> UBound(DATA_VECTOR, 1) _
Then GoTo ERROR_LABEL 'matrix not allowed

ReDim BTEMP_ARR(1 To NROWS)
ReDim ATEMP_ARR(1 To NROWS, 1 To 2)
ReDim RESULT_VECTOR(1 To NROWS, 1 To 1)

NSIZE = UBound(DATA_MATRIX) 'Normalize an mop-up system System
ReDim COEF_VECTOR(1 To NROWS)

For k = 1 To NSIZE
    i = DATA_MATRIX(k, 1)
    If Abs(DATA_MATRIX(k, 3)) < epsilon Then DATA_MATRIX(k, 3) = 0
    If i = DATA_MATRIX(k, 2) Then
        COEF_VECTOR(i) = DATA_MATRIX(k, 3)
        If COEF_VECTOR(i) = 0 Then GoTo ERROR_LABEL 'matrix not allowed
    End If
Next k

For k = 1 To NSIZE
    i = DATA_MATRIX(k, 1)
    DATA_MATRIX(k, 3) = DATA_MATRIX(k, 3) / COEF_VECTOR(i)
Next k

For i = 1 To NROWS
    DATA_VECTOR(i, 1) = DATA_VECTOR(i, 1) / COEF_VECTOR(i)
Next i
'Initial starting vector
For i = 1 To NROWS
    RESULT_VECTOR(i, 1) = 0
Next i
' Successive Over Relaxation algorithm begin
TEMP_ABS = 10 ^ 50
TEMP_NORM = 1
kkk = 0
jj = 9  'must be odd > 3
ReDim DTEMP_ARR(1 To jj + 1) 'error queue
BTEMP_MULT = -0.1

Do
    kk = 0
    k = 1
    'save starting value
    For i = 1 To NROWS
        BTEMP_ARR(i) = RESULT_VECTOR(i, 1)
    Next i
    Do
        'save queue
        For i = 1 To NROWS
            ATEMP_ARR(i, k) = RESULT_VECTOR(i, 1)
        Next i
        k = k + 1
        If k > 2 Then k = 1 'initializing Gauss-Seidel loop
        TEMP_SUM = 0
        TEMP_NORM = 0
        TEMP_RESID = 0
        SWITCH_FLAG = False
        For ii = 1 To NSIZE
            j = DATA_MATRIX(ii, 2)
            TEMP_RESID = TEMP_RESID + DATA_MATRIX(ii, 3) * RESULT_VECTOR(j, 1)
            'check end-of-row
            If ii < NSIZE Then
                If DATA_MATRIX(ii, 1) <> _
                    DATA_MATRIX(ii + 1, 1) Then SWITCH_FLAG = True
            Else
                SWITCH_FLAG = True   'the last row
            End If
            If SWITCH_FLAG Then
                'end-of row
                i = DATA_MATRIX(ii, 1)
                TEMP_RESID = DATA_VECTOR(i, 1) - TEMP_RESID       'residual
                CTEMP_MULT = ATEMP_MULT * TEMP_RESID            'new increment
                RESULT_VECTOR(i, 1) = RESULT_VECTOR(i, 1) + CTEMP_MULT
                TEMP_NORM = TEMP_NORM + Abs(CTEMP_MULT)
                TEMP_SUM = TEMP_SUM + Abs(RESULT_VECTOR(i, 1))
                TEMP_RESID = 0
                SWITCH_FLAG = False
            End If
        Next ii
        '
        TEMP_NORM = TEMP_NORM / NROWS
        TEMP_SUM = TEMP_SUM / NROWS
        If TEMP_SUM < 1 Then TEMP_SUM = 1
        If kk = 0 Then TEMP_ERR = TEMP_NORM
        kk = kk + 1
        DTEMP_ARR(kk) = TEMP_NORM
        If TEMP_NORM > TEMP_ABS Then
            'diverge too fast. stop
             ERROR_STR = "convergence has not been met"
             GoTo ERROR_LABEL
        End If
    Loop Until kk > jj 'Or TEMP_NORM < epsilon * TEMP_SUM
    kkk = kkk + kk
    'check stop
    If TEMP_NORM < epsilon * TEMP_SUM Then Exit Do
    'convergence constant TEMP_BETA : > 1 converges , < 1 diverges
    TEMP_BETA = (TEMP_ERR / TEMP_NORM) ^ (1 / kk)
    
     'check for delta extrapolation; by computing the puntual convercence
     'factor of the last 4 values

    ReDim CTEMP_ARR(1 To 4)
    For iii = 1 To 4
        jjj = UBound(DTEMP_ARR) - 4 + iii
        If DTEMP_ARR(jjj) Then: CTEMP_ARR(iii) = _
            DTEMP_ARR(jjj - 1) / DTEMP_ARR(jjj)
    Next iii
    'compute the statistic
    TEMP_MEAN = 0
    For iii = 1 To 4
        TEMP_MEAN = TEMP_MEAN + CTEMP_ARR(iii)
    Next iii
    TEMP_MEAN = TEMP_MEAN / 4
    TEMP_SIGMA = 0
    For iii = 1 To 4
        TEMP_SIGMA = TEMP_SIGMA + (TEMP_MEAN - CTEMP_ARR(iii)) ^ 2
    Next iii
    TEMP_SIGMA = Sqr(TEMP_SIGMA / 4)
   
    If TEMP_SIGMA < 10 ^ -3 Then
        'perform the delta^2 extrapolation without moving OMEGA
        For iii = 1 To UBound(ATEMP_ARR, 1)
            TEMP_DIV = RESULT_VECTOR(iii, 1) - 2 * _
                ATEMP_ARR(iii, 2) + ATEMP_ARR(iii, 1)
            If TEMP_DIV <> 0 Then
                RESULT_VECTOR(iii, 1) = RESULT_VECTOR(iii, 1) - _
                    (RESULT_VECTOR(iii, 1) - ATEMP_ARR(iii, 2)) ^ 2 / TEMP_DIV
            End If
        Next iii
    ElseIf TEMP_BETA < 1.7 And TEMP_NORM > 10 ^ -5 Then
        'try to increse the convergence
        If TEMP_BETA < 1 Then BTEMP_MULT = -0.1
        If (TEMP_BETA - TEMP_ALPHA) > 0 Then
            ATEMP_MULT = ATEMP_MULT + BTEMP_MULT
        Else
            BTEMP_MULT = -BTEMP_MULT / 2
            ATEMP_MULT = ATEMP_MULT + BTEMP_MULT
        End If
        If ATEMP_MULT < 0.1 Then ATEMP_MULT = 0.1
        If ATEMP_MULT > 1.5 Then ATEMP_MULT = 1.5
        TEMP_ALPHA = TEMP_BETA
    End If
    
    If TEMP_BETA < 1 And TEMP_SIGMA > 0.1 Then
        'discharge RESULT_VECTOR value, restore starting value
        For i = 1 To NROWS
            RESULT_VECTOR(i, 1) = BTEMP_ARR(i)
        Next i
    End If
    
Loop Until kkk > nLOOPS

If kkk > nLOOPS Then
    If TEMP_NORM < 10 ^ -3 Then
        ERROR_STR = "tolerance not reached"
    Else
        ERROR_STR = "convergence has not been met"
    End If
    GoTo ERROR_LABEL
Else
    MATRIX_SPARSE_ADSOR_SOLVER_FUNC = RESULT_VECTOR
End If

Exit Function
ERROR_LABEL:
MATRIX_SPARSE_ADSOR_SOLVER_FUNC = ERROR_STR
End Function

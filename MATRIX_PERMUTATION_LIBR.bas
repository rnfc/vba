Attribute VB_Name = "MATRIX_PERMUTATION_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_PERMUTATION_FUNC

'DESCRIPTION   : Matrix Block - Score-algorithm. Block partitioning with
'score-algorithm. It returns the permutations matrix. It consists of sequence of
'n integer unitary vectors: The parameter is a vector indicating the sequence.

'LIBRARY       : MATRIX
'GROUP         : PERMUTATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_PERMUTATION_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim NSIZE As Long

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: _
DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

NSIZE = UBound(DATA_VECTOR, 1)
ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

For i = 1 To NSIZE
    If 1 <= DATA_VECTOR(i, 1) And DATA_VECTOR(i, 1) <= NSIZE Then
        TEMP_MATRIX(DATA_VECTOR(i, 1), i) = 1
    End If
Next i

MATRIX_PERMUTATION_FUNC = TEMP_MATRIX
Exit Function
ERROR_LABEL:
MATRIX_PERMUTATION_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_BLOCK_PARTITIONED_FUNC

'DESCRIPTION   : Return the block-partitioned matrix using the Score-Algorithm
'Transforms a sparse square matrix (n x n) into a block-partitioned
'matrix by orthogonal permutation matrices. From the theory we know
'that, under certain conditions, a square matrix can be transformed
'into a block-partitioned form (also called block-triangular form)
'by similarity transformation: B = P^t * A * P; where P is a (n x n)
'permutation matrix. Note that not all matrices can be transformed in
'block-triangular form. It can be done if, and only if, the graph
'associated to the matrix is not strongly connected . On the contrary,
'if the graph is strong connected, we say that the matrix is irreducible.
'A dense matrix without zero elements, for example, is always strongly
'connected.
'LIBRARY       : MATRIX
'GROUP         : PERMUTATION
'ID            : 002



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_BLOCK_PARTITIONED_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim nLOOPS As Long

Dim TEMP_VALUE As Double
Dim FIRST_VALUE As Double
Dim SECOND_VALUE As Double

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
If NROWS <> UBound(DATA_MATRIX, 2) Then GoTo ERROR_LABEL   '"Matrix not square"
nLOOPS = 2 * NROWS

'------------------------------------------------------------------------
'      Routine for matrix reduction to the block-partitioned form
'------------------------------------------------------------------------

ReDim TEMP_VECTOR(1 To 1, 1 To NROWS)
For j = 1 To NROWS
    TEMP_VECTOR(1, j) = j
Next j

FIRST_VALUE = 0
l = 0
Do
    FIRST_VALUE = SECOND_VALUE
    For i = 1 To NROWS
        For j = NROWS To i + 1 Step -1
            If DATA_MATRIX(i, j) <> 0 Then
                For k = 1 To j - 1
                    If DATA_MATRIX(i, k) = 0 Then
                        TEMP_VALUE = MATRIX_BLOCK_REDUCTION_SCORE_FUNC(DATA_MATRIX, k, j)
                        If TEMP_VALUE > 0 Then
                            DATA_MATRIX = MATRIX_SWAP_ROW_FUNC(DATA_MATRIX, k, j)
                            DATA_MATRIX = MATRIX_SWAP_COLUMN_FUNC(DATA_MATRIX, k, j)
                            TEMP_VECTOR = MATRIX_SWAP_COLUMN_FUNC(TEMP_VECTOR, k, j)
                            SECOND_VALUE = SECOND_VALUE + TEMP_VALUE
                            Exit For
                        End If
                    End If
                Next k
            End If
        Next j
    Next i
    
    'butterfly scan begins
    For i = 1 To NROWS - 1
        If DATA_MATRIX(i, i + 1) <> 0 And DATA_MATRIX(i + 1, i) = 0 Then
            TEMP_VALUE = MATRIX_BLOCK_REDUCTION_SCORE_FUNC(DATA_MATRIX, i, i + 1)
            If TEMP_VALUE > 0 Then
                DATA_MATRIX = MATRIX_SWAP_ROW_FUNC(DATA_MATRIX, i, i + 1)
                DATA_MATRIX = MATRIX_SWAP_COLUMN_FUNC(DATA_MATRIX, i, i + 1)
                TEMP_VECTOR = MATRIX_SWAP_COLUMN_FUNC(TEMP_VECTOR, i, i + 1)
                SECOND_VALUE = SECOND_VALUE + TEMP_VALUE
            End If
        End If
    Next
    l = l + 1
Loop While SECOND_VALUE > FIRST_VALUE And l < nLOOPS

If l >= nLOOPS Then GoTo ERROR_LABEL   '"Iteration overflow"

MATRIX_BLOCK_PARTITIONED_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_BLOCK_PARTITIONED_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_BLOCK_PERMUTATION_FUNC

'DESCRIPTION   : Returns the permutation matrix that transforms a sparse square
'matrix (n x n) into a block-partitioned matrix.
'B = P^t * A * P; where P is a permutation matrix (n x n)
'This function returns the permutation vector (n)
'Note that not all matrices can be transformed in block-triangular
'form.  This usually happens if the matrix is irreducible.
'Example. Find the permutation matrix that transforms the given
'matrix into block triangular form.

'LIBRARY       : MATRIX
'GROUP         : PERMUTATION
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_BLOCK_PERMUTATION_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim nLOOPS As Long

Dim TEMP_MAX As Double
Dim TEMP_SUM As Double
Dim TEMP_VALUE As Double

Dim FIRST_VALUE As Double
Dim SECOND_VALUE As Double

Dim EDGE_VECTOR As Variant
Dim BLOCK_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
If NROWS <> UBound(DATA_MATRIX, 2) Then GoTo ERROR_LABEL   '"Matrix not square"
nLOOPS = 2 * NROWS

'------------------------------------------------------------------------
'      Routine for matrix reduction to the block-partitioned form
'------------------------------------------------------------------------

ReDim TEMP_VECTOR(1 To 1, 1 To NROWS)
For j = 1 To NROWS
    TEMP_VECTOR(1, j) = j
Next j

FIRST_VALUE = 0
l = 0
Do
    FIRST_VALUE = SECOND_VALUE
    For i = 1 To NROWS
        For j = NROWS To i + 1 Step -1
            If DATA_MATRIX(i, j) <> 0 Then
                For k = 1 To j - 1
                    If DATA_MATRIX(i, k) = 0 Then
                        TEMP_VALUE = MATRIX_BLOCK_REDUCTION_SCORE_FUNC(DATA_MATRIX, k, j)
                        If TEMP_VALUE > 0 Then
                            DATA_MATRIX = MATRIX_SWAP_ROW_FUNC(DATA_MATRIX, k, j)
                            DATA_MATRIX = MATRIX_SWAP_COLUMN_FUNC(DATA_MATRIX, k, j)
                            TEMP_VECTOR = MATRIX_SWAP_COLUMN_FUNC(TEMP_VECTOR, k, j)
                            SECOND_VALUE = SECOND_VALUE + TEMP_VALUE
                            Exit For
                        End If
                    End If
                Next k
            End If
        Next j
    Next i
    
    'butterfly scan begins
    For i = 1 To NROWS - 1
        If DATA_MATRIX(i, i + 1) <> 0 And DATA_MATRIX(i + 1, i) = 0 Then
            TEMP_VALUE = MATRIX_BLOCK_REDUCTION_SCORE_FUNC(DATA_MATRIX, i, i + 1)
            If TEMP_VALUE > 0 Then
                DATA_MATRIX = MATRIX_SWAP_ROW_FUNC(DATA_MATRIX, i, i + 1)
                DATA_MATRIX = MATRIX_SWAP_COLUMN_FUNC(DATA_MATRIX, i, i + 1)
                TEMP_VECTOR = MATRIX_SWAP_COLUMN_FUNC(TEMP_VECTOR, i, i + 1)
                SECOND_VALUE = SECOND_VALUE + TEMP_VALUE
            End If
        End If
    Next
    l = l + 1
Loop While SECOND_VALUE > FIRST_VALUE And l < nLOOPS

If l >= nLOOPS Then GoTo ERROR_LABEL   '"Iteration overflow"


'---------Check if the matrix is in the Block Vector-Jordan's form----------------

NROWS = UBound(DATA_MATRIX, 1)
ReDim BLOCK_VECTOR(1 To NROWS, 1 To 1)
ReDim EDGE_VECTOR(1 To NROWS, 1 To 1) ' build the right-border vector

For i = 1 To NROWS
    For j = NROWS To 1 Step -1
        If DATA_MATRIX(i, j) <> 0 Then Exit For
    Next j
    EDGE_VECTOR(i, 1) = j
Next i

k = 0 'max blocks found
NSIZE = 0 'dimension of the largest block
TEMP_SUM = 0

For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + 1
    If EDGE_VECTOR(i, 1) > TEMP_MAX Then TEMP_MAX = EDGE_VECTOR(i, 1)
    If i >= TEMP_MAX Then   'one BLOCK_VECTOR found
        k = k + 1
        BLOCK_VECTOR(k, 1) = TEMP_SUM
        If TEMP_SUM > NSIZE Then NSIZE = TEMP_SUM
        TEMP_SUM = 0
    End If
Next i

If k = 1 Then GoTo ERROR_LABEL         'irriducible matrix

MATRIX_BLOCK_PERMUTATION_FUNC = MATRIX_TRANSPOSE_FUNC(TEMP_VECTOR) 'Permutation Vector

Exit Function
ERROR_LABEL:
MATRIX_BLOCK_PERMUTATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_BLOCK_REDUCTION_SCORE_FUNC
'DESCRIPTION   : Score function for block-reduction routine
'LIBRARY       : MATRIX
'GROUP         : PERMUTATION
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_BLOCK_REDUCTION_SCORE_FUNC(ByRef DATA_RNG As Variant, _
ByVal SROW As Long, _
ByVal SCOLUMN As Long)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim FIRST_VALUE As Double
Dim SECOND_VALUE As Double

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
FIRST_VALUE = 0

'--------------first: Weight-function for block-reduction routine-----------------
For j = SROW + 1 To NROWS
    If DATA_MATRIX(SROW, j) = 0 Then _
        FIRST_VALUE = FIRST_VALUE + (NROWS - SROW + 1) ^ 2 * j ^ 2
Next j
For j = SCOLUMN + 1 To NROWS
    If DATA_MATRIX(SCOLUMN, j) = 0 Then _
        FIRST_VALUE = FIRST_VALUE + (NROWS - SCOLUMN + 1) ^ 2 * j ^ 2
Next j
For i = 1 To SROW - 1
    If DATA_MATRIX(i, SROW) = 0 Then _
        FIRST_VALUE = FIRST_VALUE + (NROWS - i + 1) ^ 2 * SROW ^ 2
Next i
For i = 1 To SCOLUMN - 1
    If DATA_MATRIX(i, SCOLUMN) = 0 Then _
        FIRST_VALUE = FIRST_VALUE + (NROWS - i + 1) ^ 2 * SCOLUMN ^ 2
Next i

If DATA_MATRIX(SROW, SCOLUMN) = 0 Then _
    FIRST_VALUE = FIRST_VALUE - (NROWS - SROW + 1) ^ 2 * SCOLUMN ^ 2

'after --------------Weight-function for block-reduction routine-----------------
SECOND_VALUE = 0
For j = SROW + 1 To NROWS
    If DATA_MATRIX(SCOLUMN, j) = 0 Then _
        SECOND_VALUE = SECOND_VALUE + (NROWS - SROW + 1) ^ 2 * j ^ 2
Next j
For j = SCOLUMN + 1 To NROWS
    If DATA_MATRIX(SROW, j) = 0 Then _
        SECOND_VALUE = SECOND_VALUE + (NROWS - SCOLUMN + 1) ^ 2 * j ^ 2
Next j
For i = 1 To SROW - 1
    If DATA_MATRIX(i, SCOLUMN) = 0 Then _
        SECOND_VALUE = SECOND_VALUE + (NROWS - i + 1) ^ 2 * SROW ^ 2
Next i
For i = 1 To SCOLUMN - 1
    If DATA_MATRIX(i, SROW) = 0 Then _
        SECOND_VALUE = SECOND_VALUE + (NROWS - i + 1) ^ 2 * SCOLUMN ^ 2
Next i
If DATA_MATRIX(SCOLUMN, SROW) = 0 Then _
    SECOND_VALUE = SECOND_VALUE + (NROWS - SROW + 1) ^ 2 * SCOLUMN ^ 2

MATRIX_BLOCK_REDUCTION_SCORE_FUNC = SECOND_VALUE - FIRST_VALUE

Exit Function
ERROR_LABEL:
MATRIX_BLOCK_REDUCTION_SCORE_FUNC = Err.number
End Function

Attribute VB_Name = "MATRIX_RANDOM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : GENERATE_RANDOM_MATRIX_FUNC
'DESCRIPTION   : Generates a random matrix
'LIBRARY       : MATRIX
'GROUP         : RANDOM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function GENERATE_RANDOM_MATRIX_FUNC(ByVal NROWS As Single, _
ByVal NCOLUMNS As Single, _
Optional ByVal TYPE_STR_VAL As Variant = 0, _
Optional ByVal INT_FLAG As Boolean = True, _
Optional ByVal MAX_VAL As Double = 10, _
Optional ByVal MIN_VAL As Double = -10, _
Optional ByVal SPARSE_VAL As Double = 0, _
Optional ByVal DEC_VAL As Double = 15, _
Optional ByVal FLIP_FLAG As Boolean = False, _
Optional ByVal SYMM_FLAG As Boolean = False, _
Optional ByVal THRESHOLD As Double = 15)

'-------------------------------------------------------------------
'TYPE_STR_VAL
'-------------------------------------------------------------------
'   0,"ALL" (default) - fills all cells
'   1, "SYM" - Symmetrical
'   2, "TRD" -     Tridiagonal
'   3, "DIA" -     Diagonal
'   4, "TLW" -     Triangular lower
'   5, "TUP" -     Triangular upper
'   6, "SYMTRD"    Symmetrical tridiagonal
'-------------------------------------------------------------------

'INT_FLAG = True (default) for Integer matrix, False for decimal
'MAX_VAL   = max number allowed
'MIN_VAL   = min number allowed
'SPARSE_VAL = decimal, from 0 to 1; 0 means no sparse, 1 means very sparse


Dim i As Single
Dim j As Single
Dim k As Single

Dim ii As Single
Dim jj As Single

Dim LOWER_BOUND As Double
Dim UPPER_BOUND As Double

Dim TRIANG_UP_FLAG As Boolean
Dim TRIANG_LOW_FLAG As Boolean
Dim MAT_DIAG_FLAG As Boolean
Dim MAT_TRI_DIAG_FLAG As Boolean

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

TRIANG_UP_FLAG = False
TRIANG_LOW_FLAG = False
MAT_DIAG_FLAG = False
MAT_TRI_DIAG_FLAG = False

Select Case UCase(TYPE_STR_VAL)
    Case 0, "ALLS" 'Fills all cells
    Case 1, "SYM"
        SYMM_FLAG = True
    Case 2, "TRD"
        MAT_TRI_DIAG_FLAG = True
    Case 3, "DIA"
        MAT_DIAG_FLAG = True
    Case 4, "TLW"
        TRIANG_LOW_FLAG = True
    Case 5, "TUP"
        TRIANG_UP_FLAG = True
    Case 6, "SYMTRD": SYMM_FLAG = True: MAT_TRI_DIAG_FLAG = True
End Select
'--------------------------------------

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

UPPER_BOUND = MAX_VAL
LOWER_BOUND = MIN_VAL

Randomize

For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(i, j) = (UPPER_BOUND - LOWER_BOUND + 1) * _
            CDbl(Rnd) + LOWER_BOUND
        If DEC_VAL < THRESHOLD Then TEMP_MATRIX(i, j) = _
            Round(TEMP_MATRIX(i, j), DEC_VAL)
        If INT_FLAG Then TEMP_MATRIX(i, j) = Int(TEMP_MATRIX(i, j))
    Next j
Next i

If SYMM_FLAG Then
    For i = 1 To NROWS
        For j = i + 1 To NCOLUMNS
            TEMP_MATRIX(j, i) = TEMP_MATRIX(i, j)
        Next j
    Next i
End If

For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        If FLIP_FLAG Then
            jj = NCOLUMNS + 1 - j
            If i > jj And TRIANG_UP_FLAG Then TEMP_MATRIX(i, j) = 0
            If i < jj And TRIANG_LOW_FLAG Then TEMP_MATRIX(i, j) = 0
            If i <> jj And MAT_DIAG_FLAG Then TEMP_MATRIX(i, j) = 0
            If Abs(jj - i) > 1 And MAT_TRI_DIAG_FLAG Then TEMP_MATRIX(i, j) = 0
        Else
            If i > j And TRIANG_UP_FLAG Then TEMP_MATRIX(i, j) = 0
            If i < j And TRIANG_LOW_FLAG Then TEMP_MATRIX(i, j) = 0
            If i <> j And MAT_DIAG_FLAG Then TEMP_MATRIX(i, j) = 0
            If Abs(j - i) > 1 And MAT_TRI_DIAG_FLAG Then TEMP_MATRIX(i, j) = 0
        End If
    Next j
Next i

If SPARSE_VAL > 0 Then
    If Not SYMM_FLAG Then
        ii = Int(2 * SPARSE_VAL * NROWS * NCOLUMNS)
        For k = 1 To ii
            i = Int(NROWS * Rnd + 1)
            j = Int(NCOLUMNS * Rnd + 1)
            TEMP_MATRIX(i, j) = 0
        Next k
    Else
        ii = Int(0.5 * SPARSE_VAL * NROWS * NCOLUMNS)
        For k = 1 To ii
            i = Int(NROWS * Rnd + 1)
            j = Int(NCOLUMNS * Rnd + 1)
            TEMP_MATRIX(i, j) = 0
            TEMP_MATRIX(j, i) = 0
        Next k
    End If
End If

GENERATE_RANDOM_MATRIX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
GENERATE_RANDOM_MATRIX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GENERATE_RANDOM_RANK_MATRIX_FUNC
'DESCRIPTION   : Returns a matrix with a given Rank or Determinant
'LIBRARY       : MATRIX
'GROUP         : GENERATION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function GENERATE_RANDOM_RANK_MATRIX_FUNC(ByVal NROWS As Single, _
ByVal NCOLUMNS As Single, _
Optional ByVal INT_FLAG As Boolean = True, _
Optional ByVal TRIANG_UP_FLAG As Boolean = True, _
Optional ByVal RANK_VAL As Double = 0, _
Optional ByVal DETERM_VAL As Double = 0, _
Optional ByVal SHAFFER_VAL As Double = 3, _
Optional ByVal MAX_VAL As Double = 10, _
Optional ByVal MIN_VAL As Double = -10, _
Optional ByVal DEC_VAL As Double = 15)

'INT_FLAG = True (default) for Integer matrix, False for decimal matrix
'Note: if Rank < max dimension then always Det=0

Dim h As Single
Dim i As Single
Dim j As Single
Dim k As Single

Dim hh As Single
Dim ii As Single
Dim jj As Single
Dim kk As Single

Dim TEMP_VALUE As Double
Dim UPPER_BOUND As Double
Dim LOWER_BOUND As Double

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If RANK_VAL = 0 Then: RANK_VAL = MINIMUM_FUNC(NROWS, NCOLUMNS)
If DETERM_VAL = 0 Then: DETERM_VAL = 1
If SHAFFER_VAL = 0 Then: SHAFFER_VAL = 3

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

SHAFFER_VAL = 3

UPPER_BOUND = Int(MAX_VAL / SHAFFER_VAL)
LOWER_BOUND = Int(MIN_VAL / SHAFFER_VAL)

Randomize

For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(i, j) = _
        (UPPER_BOUND - LOWER_BOUND + 1) * (Rnd) + LOWER_BOUND
        If DEC_VAL < 15 Then TEMP_MATRIX(i, j) = _
        Round(TEMP_MATRIX(i, j), DEC_VAL)
        If INT_FLAG Then TEMP_MATRIX(i, j) = Int(TEMP_MATRIX(i, j))
        If i > j And TRIANG_UP_FLAG Then TEMP_MATRIX(i, j) = 0
    Next j
Next i

hh = MINIMUM_FUNC(NROWS, NCOLUMNS)
For i = 1 To hh
    TEMP_MATRIX(i, i) = 1
Next i

TEMP_MATRIX(hh, hh) = DETERM_VAL

For i = 1 To hh - RANK_VAL 'set rank
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(hh - i + 1, j) = 0
    Next j
Next i

k = UPPER_BOUND
kk = LOWER_BOUND 'shaffer
For h = 1 To SHAFFER_VAL
    ii = 1
    For i = 2 To NROWS
        TEMP_VALUE = Int((k - kk + 1) * (Rnd) + kk)
        TEMP_MATRIX = _
        MATRIX_LINEAR_ROWS_COMBINATION_FUNC(TEMP_MATRIX, i, ii, TEMP_VALUE)
    Next i
    jj = NCOLUMNS
    For j = 1 To NCOLUMNS - 1
        TEMP_VALUE = Int((k - kk + 1) * (Rnd) + kk)
        TEMP_MATRIX = _
        MATRIX_LINEAR_COLUMNS_COMBINATION_FUNC(TEMP_MATRIX, j, jj, TEMP_VALUE)
    Next j
Next h

GENERATE_RANDOM_RANK_MATRIX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
GENERATE_RANDOM_RANK_MATRIX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GENERATE_RANDOM_SYMMETRIC_MATRIX_FUNC
'DESCRIPTION   : Returns a symmetric matrix with a given Rank or Determinant
'LIBRARY       : MATRIX
'GROUP         : GENERATION
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function GENERATE_RANDOM_SYMMETRIC_MATRIX_FUNC(ByVal NSIZE As Single, _
Optional ByVal INT_FLAG As Boolean = True, _
Optional ByVal TRIANG_UP_FLAG As Boolean = True, _
Optional ByVal RANK_VAL As Double = 0, _
Optional ByVal DETERM_VAL As Double = 0, _
Optional ByVal DEC_VAL As Double = 15, _
Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal DEVIATION_VAL As Double = 5)

'INT_FLAG = True (default) for Integer matrix, False for decimal matrix
'Note: if Rank < max dimension then always Det=0

Dim i As Single
Dim j As Single
Dim k As Single

Dim UPPER_BOUND As Double
Dim LOWER_BOUND As Double

Dim CTEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim ATEMP_MATRIX As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If RANK_VAL = 0 Then RANK_VAL = MINIMUM_FUNC(NSIZE, NSIZE)
If DETERM_VAL = 0 Then DETERM_VAL = 1

ReDim CTEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
ReDim ATEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
ReDim BTEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

UPPER_BOUND = (MEAN_VAL + DEVIATION_VAL)
LOWER_BOUND = (MEAN_VAL - DEVIATION_VAL)

Randomize
For i = 1 To NSIZE
    For j = 1 To NSIZE
        BTEMP_MATRIX(i, j) = (UPPER_BOUND - LOWER_BOUND + 1) * _
            (Rnd) + LOWER_BOUND
        If DEC_VAL < 15 Then CTEMP_MATRIX(i, j) = _
        Round(CTEMP_MATRIX(i, j), DEC_VAL)
        If INT_FLAG Then BTEMP_MATRIX(i, j) = Int(BTEMP_MATRIX(i, j))
        If i > j And TRIANG_UP_FLAG Then BTEMP_MATRIX(i, j) = 0
    Next j
Next i

k = MINIMUM_FUNC(NSIZE, NSIZE)
For i = 1 To k
    BTEMP_MATRIX(i, i) = 1
Next i

k = MINIMUM_FUNC(NSIZE, NSIZE)
For i = 1 To k
    ATEMP_MATRIX(i, i) = 1
Next
ATEMP_MATRIX(k, k) = DETERM_VAL

For i = 1 To k - RANK_VAL
    ATEMP_MATRIX(k - i + 1, k - i + 1) = 0
Next i


CTEMP_MATRIX = MATRIX_TRANSPOSE_FUNC(BTEMP_MATRIX)

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

For i = 1 To NSIZE
    For j = 1 To NSIZE
        For k = 1 To NSIZE
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + CTEMP_MATRIX(i, k) _
                * ATEMP_MATRIX(k, j)
        Next k
    Next j
Next i

CTEMP_MATRIX = TEMP_MATRIX
ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

For i = 1 To NSIZE
    For j = 1 To NSIZE
        For k = 1 To NSIZE
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + CTEMP_MATRIX(i, k) _
                * BTEMP_MATRIX(k, j)
        Next k
    Next j
Next i

GENERATE_RANDOM_SYMMETRIC_MATRIX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
GENERATE_RANDOM_SYMMETRIC_MATRIX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GENERATE_RANDOM_SPARSE_MATRIX_FUNC
'DESCRIPTION   : Generates random sparse (n x m) matrix with
'LOWER_BOUND <= values <= UPPER_BOUND

'LIBRARY       : MATRIX
'GROUP         : RANDOM
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function GENERATE_RANDOM_SPARSE_MATRIX_FUNC(ByVal NROWS As Single, _
ByVal NCOLUMNS As Single, _
Optional ByVal UPPER_BOUND As Double = 10, _
Optional ByVal LOWER_BOUND As Double = -10, _
Optional ByVal SPREAD_FACTOR As Double = 0.2, _
Optional ByVal FILL_FACTOR As Double = 0.3, _
Optional ByVal DOM_FACTOR As Double = 0.66, _
Optional ByVal SYMM_FLAG As Boolean = False, _
Optional ByVal INT_FLAG As Boolean = False, _
Optional ByVal OUTPUT As Integer = 0)

'SPREAD_FACTOR = spreading factor around the first diagonal (0 - 1)
'FILL_FACTOR = filling factor (No-zero elements)/ (total element) <= 1
'DOM_FACTOR = dominance factor < 1

'SYMM_FLAG = true/false simmetric matrix
'INT_FLAG = true/false integer matrix

Dim i As Single
Dim j As Single
Dim k As Single

Dim ii As Single 'decimals
Dim jj As Single
Dim kk As Single

Dim NSIZE As Single
Dim MSIZE As Single

Dim TEMP_SUM As Double
Dim TEMP_VAL As Double
Dim TEMP_FACTOR As Double
Dim TEMP_SCALE As Double
Dim TEMP_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

'REMEMBER:
'Dominance factor must be 0 < D < 1
'Filling factor must be 0 < F < 1
'Spreading factor must be 0 < S < 1

If SPREAD_FACTOR < 0 Then SPREAD_FACTOR = 0 'check parameters bounding
If SPREAD_FACTOR > 1 Then SPREAD_FACTOR = 1
If FILL_FACTOR < 0 Then FILL_FACTOR = 0
If FILL_FACTOR > 1 Then FILL_FACTOR = 1
If FILL_FACTOR < 0 Then DOM_FACTOR = 0
If INT_FLAG Then ii = 0 Else ii = 3

NSIZE = 2 * NROWS * NCOLUMNS * FILL_FACTOR 'total non-zero elements
If FILL_FACTOR < 0.4 Then
      TEMP_FACTOR = NCOLUMNS / 2 * FILL_FACTOR
Else: TEMP_FACTOR = NCOLUMNS / 2 * (0.6 + FILL_FACTOR) * FILL_FACTOR
End If

ReDim TEMP_MATRIX(1 To NSIZE, 1 To 3)
MSIZE = NSIZE
Randomize
k = 0
TEMP_SUM = 0

Do
    TEMP_SUM = TEMP_SUM + 1
    For i = 1 To NROWS
        ReDim TEMP_VECTOR(1 To NCOLUMNS)
        'generate TEMP_FACTOR random elements on the row i
        'centred around the pivot column jj-th
        jj = Round(NROWS / NCOLUMNS * i, 0)
        'add the pivot element
        TEMP_VECTOR(jj) = 1
        For kk = 1 To TEMP_FACTOR
            TEMP_SCALE = Rnd ^ (25 / (1 + 24 * SPREAD_FACTOR))
            If SPREAD_FACTOR = 0 Then TEMP_SCALE = 0
            TEMP_VAL = NCOLUMNS * TEMP_SCALE + jj  'random column
            Do While TEMP_VAL <= NCOLUMNS
                If TEMP_VECTOR(TEMP_VAL) = 0 Then
                    'add e new element and exit
                    TEMP_VECTOR(TEMP_VAL) = Round((UPPER_BOUND - LOWER_BOUND) * _
                        Rnd + LOWER_BOUND, ii)
                    Exit Do
                End If
                TEMP_VAL = TEMP_VAL + 1
            Loop
        Next kk
        
        'save the row into the sparse matrix
        For j = 1 To NCOLUMNS
            If TEMP_VECTOR(j) <> 0 And i <> j Then
                'check if exists enough space
                If k > NSIZE - 2 Then
                    NSIZE = UBound(TEMP_MATRIX, 1) + MSIZE
                    TEMP_MATRIX = MATRIX_REDIM_FUNC(TEMP_MATRIX, MSIZE, 0)
                End If
                If TEMP_SUM = 1 Then
                    k = k + 1
                    TEMP_MATRIX(k, 1) = i
                    TEMP_MATRIX(k, 2) = j
                    TEMP_MATRIX(k, 3) = TEMP_VECTOR(j)
                End If
                If TEMP_SUM = 2 Or SYMM_FLAG Then
                    k = k + 1
                    TEMP_MATRIX(k, 1) = j
                    TEMP_MATRIX(k, 2) = i
                    TEMP_MATRIX(k, 3) = TEMP_VECTOR(j)
                End If
            End If
        Next j
    Next i
    If TEMP_SUM = 2 Or (TEMP_SUM = 1 And SYMM_FLAG) Then Exit Do
Loop

'add the diagonal elements
ReDim TEMP_VECTOR(1 To NROWS)
For j = 1 To NSIZE
    i = TEMP_MATRIX(j, 1)
    If i = 0 Then Exit For
    TEMP_VECTOR(i) = TEMP_VECTOR(i) + Abs(TEMP_MATRIX(j, 3))
Next j
'
If k + NROWS > NSIZE Then
    NSIZE = UBound(TEMP_MATRIX, 1) + MSIZE
    TEMP_MATRIX = MATRIX_REDIM_FUNC(TEMP_MATRIX, MSIZE, 0)
End If

TEMP_SCALE = DOM_FACTOR / (1 - DOM_FACTOR)
For i = 1 To NROWS
    k = k + 1
    TEMP_MATRIX(k, 1) = i
    TEMP_MATRIX(k, 2) = i
    TEMP_MATRIX(k, 3) = Round(TEMP_VECTOR(i) * TEMP_SCALE, ii)
Next i
'
If k < NSIZE Then NSIZE = k 'reset the true diemsion

DATA_MATRIX = TEMP_MATRIX
DATA_MATRIX = MATRIX_TRIM_FUNC(DATA_MATRIX, 1, "")
DATA_MATRIX = MATRIX_DOUBLE_SORT_FUNC(DATA_MATRIX)

Select Case OUTPUT
    Case 0
        GENERATE_RANDOM_SPARSE_MATRIX_FUNC = DATA_MATRIX
    Case 1
        GENERATE_RANDOM_SPARSE_MATRIX_FUNC = MATRIX_SPARSE_CONVERT_FUNC(DATA_MATRIX, 0)
    Case Else
        GENERATE_RANDOM_SPARSE_MATRIX_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
GENERATE_RANDOM_SPARSE_MATRIX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GENERATE_RANDOM_TOEPLITZ_MATRIX_FUNC

'DESCRIPTION   : Generate a random Toeplitz matrix
'LIBRARY       : MATRIX
'GROUP         : RANDOM
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function GENERATE_RANDOM_TOEPLITZ_MATRIX_FUNC(ByVal NROWS As Single, _
Optional ByVal INT_FLAG As Boolean = True, _
Optional ByVal MAX_VAL As Double = 10, _
Optional ByVal MIN_VAL As Double = -10, _
Optional ByVal SPARSE_VAL As Double = 0, _
Optional ByVal DEC_VAL As Double = 15, _
Optional ByVal SYMM_FLAG As Boolean = False)

Dim h As Single
Dim i As Single
Dim j As Single
Dim k As Single

Dim NCOLUMNS As Single

Dim UPPER_BOUND As Double
Dim LOWER_BOUND As Double

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

NCOLUMNS = 2 * NROWS - 1
ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)
ReDim DATA_VECTOR(1 To NCOLUMNS, 1 To 1)

UPPER_BOUND = MAX_VAL
LOWER_BOUND = MIN_VAL

Randomize

For j = 1 To NCOLUMNS
    DATA_VECTOR(j, 1) = (UPPER_BOUND - _
        LOWER_BOUND + 1) * (Rnd) + LOWER_BOUND
    If DEC_VAL < 15 Then DATA_VECTOR(j, 1) = _
        Round(DATA_VECTOR(j, 1), DEC_VAL)
    If INT_FLAG Then DATA_VECTOR(j, 1) = Int(DATA_VECTOR(j, 1))
Next j

'-------------------------------Make Sparse----------------------------
If SPARSE_VAL > 0 Then
    h = Int(SPARSE_VAL * NCOLUMNS)
    For k = 1 To h
        i = Int(NCOLUMNS * Rnd + 1)
        DATA_VECTOR(i, 1) = 0
    Next k
End If

'--------------------------------Symmetrize----------------------------
If SYMM_FLAG Then
    For j = 1 To NROWS - 1
        DATA_VECTOR(2 * NROWS - j, 1) = DATA_VECTOR(j, 1)
    Next j
End If

For i = 1 To NROWS
    TEMP_MATRIX(1, i) = DATA_VECTOR(NROWS - i + 1, 1)
    TEMP_MATRIX(i, 1) = DATA_VECTOR(NROWS + i - 1, 1)
Next i

For i = 2 To NROWS
    For j = 2 To NROWS
        TEMP_MATRIX(i, j) = TEMP_MATRIX(i - 1, j - 1)
    Next j
Next i

GENERATE_RANDOM_TOEPLITZ_MATRIX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
GENERATE_RANDOM_TOEPLITZ_MATRIX_FUNC = Err.number
End Function

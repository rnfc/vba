Attribute VB_Name = "MATRIX_GAUSS_JORDAN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SQUARE_GJ_DETERM_FUNC
'DESCRIPTION   : Returns the determinant of a square matrix (n x n)
'LIBRARY       : MATRIX
'GROUP         : GAUSS_JORDAN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SQUARE_GJ_DETERM_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal epsilon As Double = 10 ^ -15)

'VERSION switch (default 0) sets the floating point or
'integer arithmetic (1). Integer arithmetic is intrinsically
'more accurate for integer matrices but it may easily reaches
'the overflow. Use it only with integer matrices of low dimension.
'Epsilon (default is 0) sets the minimum round-off error; any value
'with an absolute value less than epsilon will be set to zero.


Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 1)

    If NROWS <> NCOLUMNS Then: GoTo ERROR_LABEL
    
    If VERSION = 0 Then
        MATRIX_SQUARE_GJ_DETERM_FUNC = MATRIX_GS_REDUCTION_PIVOT_FUNC(DATA_MATRIX, , epsilon, 2)
    Else
        MATRIX_SQUARE_GJ_DETERM_FUNC = MATRIX_GS_INTEGER_REDUCTION_FUNC(DATA_MATRIX, , epsilon, 1)
    End If

Exit Function
ERROR_LABEL:
MATRIX_SQUARE_GJ_DETERM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GS_DIAGONAL_TRIANGULAR_REDUCTION_FUNC

'DESCRIPTION   : This function traces, step by step, the diagonal reduction
'(Gauss-Jordan algorithm ) or triangular reduction (Gauss algorithm) of a matrix.

'LIBRARY       : MATRIX
'GROUP         : GAUSS_JORDAN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_GS_DIAGONAL_TRIANGULAR_REDUCTION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 1, _
Optional ByVal INT_FLAG = True, _
Optional ByVal epsilon As Double = 2 * 10 ^ -15)

'Optional parameter VERSION can be 0 for Diagonal or 1 for Triangular.
'Optional parameter INT_FLAG = TRUE forces the
'function to conserve integer values through all steps. Default is FALSE.
'The argument DATA_RNG is the complete matrix (n x m) of the linear system


Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long 'pivot

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim P_VAL As Double
Dim Q_VAL As Double

Dim MAX_VAL As Double
Dim MCM_VAL As Double
Dim DETERM_VAL As Double

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
'------inizio algoritmo-----------
k = 1

1983:

i = 1
For jj = k To NROWS
    'search for first non zero element in column k; returns the first non
    'zero element out of the diagonal searching for row
    
    If VERSION = 1 Then ii = jj + 1 Else ii = 1
    'VERSION 0: searchs for all matrix
    'VERSION 1: serchs for lower-triangular
        
        For i = ii To NROWS
            If Abs(DATA_MATRIX(i, jj)) <= epsilon Then _
                DATA_MATRIX(i, jj) = 0
            If i <> jj And DATA_MATRIX(i, jj) <> 0 Then
                k = jj
                GoTo jumper
            End If
        Next i
Next jj
jumper:
    If k > NROWS Or i > NROWS Then GoTo 1984 'search pivot
    kk = k
    MAX_VAL = 0
    For ii = k To NROWS
        If Abs(DATA_MATRIX(ii, k)) > MAX_VAL Then
            kk = ii
            MAX_VAL = Abs(DATA_MATRIX(ii, k))
        End If
    Next ii
    
    ' swap row if need
    If kk > k Then
        DATA_MATRIX = MATRIX_SWAP_ROW_FUNC(DATA_MATRIX, k, kk)
        GoTo 1985
    End If
    ' pivot nullo
    If kk = k And DATA_MATRIX(k, k) = 0 Then
        If k = NROWS Then GoTo 1984  'last iteration
        k = k + 1  'continue with next colum
        GoTo 1983
    End If
    
    If Not INT_FLAG Then
        'linear reduction with decimal value. Original Gauss-Jordan Algorith
        If i <> k And DATA_MATRIX(i, k) <> 0 Then
            Q_VAL = -DATA_MATRIX(i, k) / DATA_MATRIX(k, k)
            For j = 1 To NCOLUMNS
                DATA_MATRIX(i, j) = DATA_MATRIX(i, j) + Q_VAL * DATA_MATRIX(k, j)
                If Abs(DATA_MATRIX(i, j)) < epsilon Then _
                    DATA_MATRIX(i, j) = 0 'mop-up
            Next j
            
        End If
    Else
        'linear reduction with integer value. For didactic scope
        If i <> k And DATA_MATRIX(i, k) <> 0 Then
            MCM_VAL = PAIR_MCM_FUNC(Abs(DATA_MATRIX(k, k)), Abs(DATA_MATRIX(i, k)))
            Q_VAL = -MCM_VAL / DATA_MATRIX(k, k)
            P_VAL = MCM_VAL / DATA_MATRIX(i, k)
            DETERM_VAL = DETERM_VAL * P_VAL
            For j = 1 To NCOLUMNS
                DATA_MATRIX(i, j) = P_VAL * DATA_MATRIX(i, j) + Q_VAL * DATA_MATRIX(k, j)
            Next j
        End If
    End If
    GoTo 1985
1984:
    For i = 1 To NROWS
        jj = DATA_MATRIX(i, i)
        If jj <> 0 Then
            For j = i To NCOLUMNS
                DATA_MATRIX(i, j) = DATA_MATRIX(i, j) / jj
            Next j
        End If
    Next i
1985:
    MATRIX_GS_DIAGONAL_TRIANGULAR_REDUCTION_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_GS_DIAGONAL_TRIANGULAR_REDUCTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PRINT_MATRIX_GS_DIAGONAL_TRIANGULAR_REDUCTION_FUNC

'DESCRIPTION   : This function has been developed for its didactic scope.
'It traces, step by step, the diagonal reduction (Gauss-Jordan
'algorithm ) or triangular reduction (Gauss algorithm) of a matrix.

'LIBRARY       : MATRIX
'GROUP         : GAUSS_JORDAN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function PRINT_MATRIX_GS_DIAGONAL_TRIANGULAR_REDUCTION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal PIVOT_FLAG As Boolean = False, _
Optional ByVal DIAG_FLAG As Boolean = True)

'If DIAG_FLAG = TRUE; Then DIAG Else: TRIDIAG

'Optional parameter DIAG_FLAG can be True for Diagonal or False for
'Triangular. the default is D

'The argument DATA_RNG is the complete matrix (n x m) of the linear system

'Remember that for a linear system:
'DATA_RNG  is the system square matrix (n x n)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim AROWS As Long
Dim ACOLUMNS As Long

Dim MCM_VAL As Double
Dim MAX_VAL As Double
Dim DETERM_VAL As Double

Dim TEMP_STR As String
Dim INT_FLAG As Boolean

Dim ATEMP_VAL As Variant
Dim BTEMP_VAL As Variant
Dim CTEMP_VAL As Variant
Dim DTEMP_VAL As Variant

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

If UBound(DATA_MATRIX, 1) < 3 Then: GoTo ERROR_LABEL
If UBound(DATA_MATRIX, 2) < 3 Then: GoTo ERROR_LABEL

SROW = LBound(DATA_MATRIX, 1)
SCOLUMN = LBound(DATA_MATRIX, 2)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_VECTOR(1 To NROWS)
ReDim TEMP_MATRIX(1 To 100000, 1 To NCOLUMNS + 1)

ACOLUMNS = SCOLUMN + NCOLUMNS
AROWS = SROW + (NROWS + 1)

GoSub 1985


'-------------------------check if the matrix is integer-------------------------
'This routing forces the function to conserve integer values through all steps

INT_FLAG = True
For i = 1 To NROWS
    For j = 1 To NROWS
        If Not IS_INTEGER_FUNC(DATA_MATRIX(i, j)) Then
            INT_FLAG = False: Exit For
        End If
    Next j
Next i
'--------------------------------------------------------------------------------

l = 1
DETERM_VAL = 1
DTEMP_VAL = 1
For k = 1 To NROWS
    kk = k
    If PIVOT_FLAG = True Or Abs(DATA_MATRIX(k, k)) = 0 Then
        'search max pivot in column k
        MAX_VAL = 0
        For i = k To NROWS
            If Abs(DATA_MATRIX(i, k)) > MAX_VAL Then
                kk = i
                MAX_VAL = Abs(DATA_MATRIX(i, k))
            End If
        Next i
    End If
    ' swap row
    If kk > k Then
        DATA_MATRIX = MATRIX_SWAP_ROW_FUNC(DATA_MATRIX, k, kk)
        DETERM_VAL = -DETERM_VAL
        i = kk
        ATEMP_VAL = "< swap"
        BTEMP_VAL = "< swap"
        GoSub 1983
    End If
    ' pivot nullo
    If kk = k And DATA_MATRIX(k, k) = 0 Then
        DETERM_VAL = 0
        l = l - 1
        TEMP_MATRIX(AROWS + NROWS + 1, SCOLUMN) = _
           "DETERM(A" & Trim(CStr(l)) & ") = 0"
        TEMP_MATRIX(AROWS + NROWS + 2, SCOLUMN) = _
            "DETERM(A) = 0"
        GoTo 1982
    End If
    
    If DIAG_FLAG = False Then
          h = k + 1 ' Triangolarizza
    Else
          h = 1     ' Diagonalizza
    End If
    
    For i = h To NROWS
        If INT_FLAG Then
'linear reduction with integer values. Modified Gauss-Jordan algorithm
            If i <> k And DATA_MATRIX(i, k) <> 0 Then
                MCM_VAL = _
                    PAIR_MCM_FUNC(Abs(DATA_MATRIX(k, k)), Abs(DATA_MATRIX(i, k)))
                ATEMP_VAL = MCM_VAL / DATA_MATRIX(k, k)
                BTEMP_VAL = -MCM_VAL / DATA_MATRIX(i, k)
                DTEMP_VAL = DTEMP_VAL * BTEMP_VAL
                For j = 1 To NCOLUMNS
                    DATA_MATRIX(i, j) = BTEMP_VAL * _
                        DATA_MATRIX(i, j) + ATEMP_VAL * DATA_MATRIX(k, j)
                Next j
                GoSub 1983
            End If
        Else
'linear reduction with decimal values. Original Gauss-Jordan algorithm
            If i <> k And DATA_MATRIX(i, k) <> 0 Then
                ATEMP_VAL = -DATA_MATRIX(i, k) / DATA_MATRIX(k, k)
                BTEMP_VAL = 1
                For j = 1 To NCOLUMNS
                    DATA_MATRIX(i, j) = DATA_MATRIX(i, j) + _
                        ATEMP_VAL * DATA_MATRIX(k, j)
                Next j
                GoSub 1983
            End If
        End If
    Next i
Next k


If (DETERM_VAL / DTEMP_VAL) <> 0 Then
    For k = 1 To NROWS     'Normalize
        TEMP_VECTOR(k) = DATA_MATRIX(k, k)
        DETERM_VAL = DETERM_VAL * TEMP_VECTOR(k)
        For j = k To NCOLUMNS
            DATA_MATRIX(k, j) = DATA_MATRIX(k, j) / TEMP_VECTOR(k)
        Next j
    Next k
    CTEMP_VAL = DETERM_VAL / DTEMP_VAL
    GoSub 1984:
Else
    GoTo ERROR_LABEL
End If


'------------------------SOME HOUSE KEEPING-----------------------------
1982:
TEMP_MATRIX = MATRIX_TRIM_FUNC(TEMP_MATRIX, 1, "")
For j = LBound(TEMP_MATRIX, 2) To UBound(TEMP_MATRIX, 2)
    For i = LBound(TEMP_MATRIX, 1) To UBound(TEMP_MATRIX, 1)
        If IsEmpty(TEMP_MATRIX(i, j)) = True Then: TEMP_MATRIX(i, j) = ""
    Next i
Next j
PRINT_MATRIX_GS_DIAGONAL_TRIANGULAR_REDUCTION_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------

Exit Function

1983: 'write_step_results
    TEMP_MATRIX(AROWS + k - 1, SCOLUMN + NCOLUMNS) = ATEMP_VAL
    TEMP_MATRIX(AROWS + i - 1, SCOLUMN + NCOLUMNS) = BTEMP_VAL
        
    TEMP_STR = NUMBER_FRACTION_STRING1_FUNC(DTEMP_VAL, DETERM_VAL)
    
    TEMP_MATRIX(AROWS + NROWS + 1, SCOLUMN) = _
        "DETERM(A" & Trim(CStr(l)) & ") = " & TEMP_STR & " DETERM(A)"
    
    AROWS = AROWS + NROWS + 3
    GoSub 1985
'PrintMatrix DATA_MATRIX, AROWS, SCOLUMN
    l = l + 1
Return
'-----------------------------------------------------------------------

1984: 'write_Normalize_Matrix
    If INT_FLAG Then
        For k = 1 To NROWS
            TEMP_MATRIX(AROWS + k - 1, _
                SCOLUMN + NCOLUMNS) = "'" & NUMBER_FRACTION_STRING1_FUNC(1, TEMP_VECTOR(k))
        Next k
    Else
        For k = 1 To NROWS
            TEMP_MATRIX(AROWS + k - 1, _
                SCOLUMN + NCOLUMNS) = 1 / TEMP_VECTOR(k)
        Next k
    End If
    TEMP_STR = NUMBER_FRACTION_STRING1_FUNC(DTEMP_VAL, DETERM_VAL)
    TEMP_MATRIX(AROWS + NROWS + 1, SCOLUMN) = _
        "DETERM(A" & Trim(CStr(l)) & _
            ") = " & TEMP_STR & " DETERM(A)"
    AROWS = AROWS + NROWS + 3
    'printMatrix DATA_MATRIX, AROWS, SCOLUMN
    GoSub 1985
    
    TEMP_MATRIX(AROWS + NROWS + 1, SCOLUMN) = _
        "DETERM(A" & Trim(CStr(l)) & ") = 1"
    TEMP_MATRIX(AROWS + NROWS + 2, SCOLUMN) = _
        "DETERM(A) = " & NUMBER_FRACTION_STRING1_FUNC(CTEMP_VAL, 1)
Return
'-----------------------------------------------------------------------

1985: 'Print_matrix
    For ii = 1 To UBound(DATA_MATRIX, 1)
        For jj = 1 To UBound(DATA_MATRIX, 2)
            TEMP_MATRIX(AROWS - 1 + ii, _
                SCOLUMN - 1 + jj) = DATA_MATRIX(ii, jj)
        Next jj
    Next ii
Return

'----------------------------------------------------------------------------
ERROR_LABEL:
PRINT_MATRIX_GS_DIAGONAL_TRIANGULAR_REDUCTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GS_REDUCTION_PIVOT_FUNC
'DESCRIPTION   : Gauss-Jordan algorithm for matrix reduction with full
'pivot method (solve with floating arithmetic)

'LIBRARY       : MATRIX
'GROUP         : GAUSS_JORDAN
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_GS_REDUCTION_PIVOT_FUNC(ByRef XDATA_RNG As Variant, _
Optional ByRef YDATA_RNG As Variant, _
Optional ByVal epsilon As Double = 2E-16, _
Optional ByVal OUTPUT As Integer = 0)

'XDATA_MATRIX is XDATA_MATRIX matrix (NROWS x NROWS); at the end contains the
'inverse of XDATA_MATRIX
'YDATA_VECTOR is XDATA_MATRIX matrix (NROWS x NCOLUMNS); at the end
'cotains the solution of AX=B. This version apply the check for too small
'elements: |aij|<Tiny


Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Double

Dim MAX_VAL As Double
Dim DETERM_VAL As Double

Dim DETERM_FLAG As Boolean

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL


DETERM_FLAG = True

XDATA_MATRIX = XDATA_RNG

If IsArray(YDATA_RNG) = True Then
    YDATA_VECTOR = YDATA_RNG
    NCOLUMNS = UBound(YDATA_VECTOR, 2)
Else
    NCOLUMNS = 0
End If

NROWS = UBound(XDATA_MATRIX, 1)
ReDim TEMP_MATRIX(1 To 2 * NROWS, 1 To 3) 'trace of swaps
kk = 0 'swap
DETERM_VAL = 1


For k = 1 To NROWS
    'search max pivot
    ii = k
    jj = k
    
    MAX_VAL = 0
    For i = k To NROWS
        For j = k To NROWS
            If Abs(XDATA_MATRIX(i, j)) > MAX_VAL Then
                ii = i
                jj = j
                MAX_VAL = Abs(XDATA_MATRIX(i, j))
            End If
        Next j
    Next i
    
    If ii = jj And Abs(XDATA_MATRIX(k, k)) <> 0 Then
        ii = k
        jj = k
    End If

    ' swap rows and columns
    If ii > k Then
        XDATA_MATRIX = MATRIX_SWAP_ROW_FUNC(XDATA_MATRIX, k, ii)
        If NCOLUMNS > 0 Then YDATA_VECTOR = MATRIX_SWAP_ROW_FUNC(YDATA_VECTOR, k, ii)
        If DETERM_FLAG Then DETERM_VAL = -DETERM_VAL
        kk = kk + 1
        TEMP_MATRIX(kk, 1) = k
        TEMP_MATRIX(kk, 2) = ii
        TEMP_MATRIX(kk, 3) = 1
    End If
    
    If jj > k Then
        XDATA_MATRIX = MATRIX_SWAP_COLUMN_FUNC(XDATA_MATRIX, k, jj)
        If DETERM_FLAG Then DETERM_VAL = -DETERM_VAL
        kk = kk + 1
        TEMP_MATRIX(kk, 1) = k
        TEMP_MATRIX(kk, 2) = jj
        TEMP_MATRIX(kk, 3) = 2
    End If
    
    ' check pivot 0
    If Abs(XDATA_MATRIX(k, k)) <= epsilon Then
        XDATA_MATRIX(k, k) = 0
        DETERM_VAL = 0
        GoTo 1983
    End If
    
    'normalization
    TEMP_VAL = XDATA_MATRIX(k, k)
    If DETERM_FLAG Then DETERM_VAL = DETERM_VAL * TEMP_VAL
    XDATA_MATRIX(k, k) = 1
    For j = 1 To NROWS
        XDATA_MATRIX(k, j) = XDATA_MATRIX(k, j) / TEMP_VAL
    Next j
    For j = 1 To NCOLUMNS
        YDATA_VECTOR(k, j) = YDATA_VECTOR(k, j) / TEMP_VAL
    Next j
    'linear reduction
    For i = 1 To NROWS
        If i <> k And XDATA_MATRIX(i, k) <> 0 Then
            TEMP_VAL = XDATA_MATRIX(i, k)
            XDATA_MATRIX(i, k) = 0
            For j = 1 To NROWS
                XDATA_MATRIX(i, j) = XDATA_MATRIX(i, j) - TEMP_VAL * XDATA_MATRIX(k, j)
            Next j
            For j = 1 To NCOLUMNS
                YDATA_VECTOR(i, j) = YDATA_VECTOR(i, j) - TEMP_VAL * YDATA_VECTOR(k, j)
            Next j
        End If
    Next i
Next k

'scramble rows
For i = kk To 1 Step -1
    If TEMP_MATRIX(i, 3) = 1 Then
        XDATA_MATRIX = _
            MATRIX_SWAP_COLUMN_FUNC(XDATA_MATRIX, TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2))
    Else
        XDATA_MATRIX = _
            MATRIX_SWAP_ROW_FUNC(XDATA_MATRIX, TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2))
        If NCOLUMNS > 0 Then YDATA_VECTOR = _
            MATRIX_SWAP_ROW_FUNC(YDATA_VECTOR, TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2))
    End If
Next i

1983:
            
Select Case OUTPUT
Case 0
    MATRIX_GS_REDUCTION_PIVOT_FUNC = YDATA_VECTOR 'Solution
Case 1
    MATRIX_GS_REDUCTION_PIVOT_FUNC = XDATA_MATRIX 'Inverse
Case 2
    MATRIX_GS_REDUCTION_PIVOT_FUNC = DETERM_VAL 'Determinant
Case Else
    MATRIX_GS_REDUCTION_PIVOT_FUNC = Array(YDATA_VECTOR, XDATA_MATRIX, DETERM_VAL)
End Select

Exit Function
ERROR_LABEL:
MATRIX_GS_REDUCTION_PIVOT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GS_SINGULAR_LINEAR_SYSTEM_FUNC

'DESCRIPTION   : A singular linear system can have infinitely many
'solutions or even none.

'This happens when DET(A) = 0. In that case the following equations:
'Ax=b; Ax=0
'define both an implicit Linear Function - also called Linear Transformation
'- between vector spaces, which can be put in the following explicit form
'y = Cx + d

'where C is the transformation matrix and d is the known vector; C is a
'square matrix having the same columns of A, and d the same dimension of  b
'This function returns the matrix C in the first n columns; eventually,
'a last column contains the vector d (only if b is not missing). If the
'system has no solution, this function returns an error

'The optional parameter tolerance sets the relative precision level; elements
'lower than this level are forced to zero. Default is 1E-13.

'LIBRARY       : MATRIX
'GROUP         : GAUSS_JORDAN
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_GS_SINGULAR_LINEAR_SYSTEM_FUNC(ByRef XDATA_RNG As Variant, _
Optional ByRef YDATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal tolerance As Double = 10 ^ -15, _
Optional ByVal OUTPUT As Integer = 0)

'VERSION = 0 --> diagonal
'VERSION = 1 --> triangle

'This version solves also systems where the number of equations is less
'than the number of variables; In other words, A is a rectangular matrix
'(n x m) where n < m

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long 'pivot

Dim RM_VAL As Long
Dim CM_VAL As Long
Dim RV_VAL As Long
Dim CV_VAL As Long

Dim NROWS As Long
Dim NCOLUMNS As Long


Dim MAX_VAL As Double
Dim TEMP_VAL As Double
Dim SCALE_VAL As Double
Dim DETERM_VAL As Variant

Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim LAMBDA As Double
Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 0
LAMBDA = 0
If tolerance = 0 Then: tolerance = 10 ^ -15

XDATA_MATRIX = XDATA_RNG

RM_VAL = UBound(XDATA_MATRIX, 1)
CM_VAL = UBound(XDATA_MATRIX, 2)

If IsArray(YDATA_RNG) = False Then
    RV_VAL = RM_VAL
    CV_VAL = 1
Else
    YDATA_VECTOR = YDATA_RNG
    RV_VAL = UBound(YDATA_VECTOR, 1)
    CV_VAL = UBound(YDATA_VECTOR, 2)
End If

If RM_VAL <> RV_VAL Or CV_VAL <> 1 Then: GoTo ERROR_LABEL

NROWS = MAXIMUM_FUNC(RM_VAL, CM_VAL)
NCOLUMNS = 1
ReDim TEMP_MATRIX(1 To NROWS, 1 To CM_VAL + NCOLUMNS)

For i = 1 To RM_VAL
    For j = 1 To CM_VAL
        TEMP_MATRIX(i, j) = XDATA_MATRIX(i, j)
        If Abs(TEMP_MATRIX(i, j)) > epsilon Then _
            epsilon = Abs(TEMP_MATRIX(i, j))
    Next j
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(i, j + CM_VAL) = 0
        If IsArray(YDATA_VECTOR) = True Then _
            TEMP_MATRIX(i, j + CM_VAL) = YDATA_VECTOR(i, j)
    Next j
Next i

If epsilon > 1 Then
    LAMBDA = tolerance * epsilon
Else
    LAMBDA = tolerance
End If

NCOLUMNS = NROWS + NCOLUMNS

DETERM_VAL = 1
For k = 1 To NROWS
    Select Case VERSION
        Case 0 'Diagonalizza
            h = 1
        Case Else '"T" Triangolarizza
            h = k + 1
    End Select
    
    n = k     'search max pivot in column k
    MAX_VAL = Abs(TEMP_MATRIX(k, k))
    For i = k + 1 To NROWS
        If Abs(TEMP_MATRIX(i, k)) > MAX_VAL Then
            n = i
            MAX_VAL = Abs(TEMP_MATRIX(i, k))
        End If
    Next i
    
    If n > k Then     ' swap row
        TEMP_MATRIX = MATRIX_SWAP_ROW_FUNC(TEMP_MATRIX, k, n)
        DETERM_VAL = -DETERM_VAL
    End If
    
    If Abs(TEMP_MATRIX(k, k)) <= LAMBDA Then     ' check pivot 0
        TEMP_MATRIX(k, k) = 0
        DETERM_VAL = 0
        GoTo 1983
    End If

    SCALE_VAL = TEMP_MATRIX(k, k)     'normalization
    DETERM_VAL = DETERM_VAL * SCALE_VAL
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(k, j) = TEMP_MATRIX(k, j) / SCALE_VAL
    Next j

    For i = h To NROWS     'linear reduction
        If i <> k And TEMP_MATRIX(i, k) <> 0 Then
            SCALE_VAL = TEMP_MATRIX(i, k)
            For j = 1 To NCOLUMNS
                TEMP_MATRIX(i, j) = _
                    TEMP_MATRIX(i, j) - SCALE_VAL * TEMP_MATRIX(k, j)
            Next j
        End If
    Next i
Next k

1983:
For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        If Abs(TEMP_MATRIX(i, j)) < LAMBDA Then TEMP_MATRIX(i, j) = 0
    Next j
Next i
'-----------------
For i = 1 To NROWS 'search for row null except one element (if exist)
    m = 0
    l = 0
    For j = 1 To NROWS
        If TEMP_MATRIX(i, j) <> 0 Then
            m = m + 1
            l = j
        End If
    Next j
    If m = 1 And l <> i Then: TEMP_MATRIX = MATRIX_SWAP_ROW_FUNC(TEMP_MATRIX, i, l)
    If m = 0 Then 'check if the problem is impossible
        For j = NROWS + 1 To NCOLUMNS
            If TEMP_MATRIX(i, j) <> 0 Then: GoTo ERROR_LABEL
        Next j
    End If
Next i

For k = 1 To NROWS
    If TEMP_MATRIX(k, k) <> 0 Then 'cerca un altro elemento non zero sopra
        For i = k - 1 To 1 Step -1
            If TEMP_MATRIX(i, k) <> 0 And i <> k Then
                TEMP_VAL = -TEMP_MATRIX(i, k)
                SCALE_VAL = TEMP_MATRIX(k, k)
                For j = 1 To NCOLUMNS 'linear combination between l and i rows
                    TEMP_MATRIX(i, j) = SCALE_VAL * _
                        TEMP_MATRIX(i, j) + TEMP_VAL * TEMP_MATRIX(k, j)
                Next j
            End If
        Next i
    End If
Next k

For i = 1 To NROWS 'normalize
    If TEMP_MATRIX(i, i) <> 0 And TEMP_MATRIX(i, i) <> 1 Then
        TEMP_VAL = TEMP_MATRIX(i, i)
        For j = 1 To NCOLUMNS
                TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) / TEMP_VAL
        Next j
    End If
Next i

For j = 1 To NROWS
    If TEMP_MATRIX(j, j) = 0 Then
        For i = 1 To NROWS
            TEMP_MATRIX(i, j) = -TEMP_MATRIX(i, j)
        Next i
        TEMP_MATRIX(j, j) = 1
    Else
        TEMP_MATRIX(j, j) = 0
    End If
Next j

Select Case OUTPUT
Case 0
    MATRIX_GS_SINGULAR_LINEAR_SYSTEM_FUNC = TEMP_MATRIX
Case 1
    MATRIX_GS_SINGULAR_LINEAR_SYSTEM_FUNC = DETERM_VAL 'Determinant
Case Else
    MATRIX_GS_SINGULAR_LINEAR_SYSTEM_FUNC = Array(TEMP_MATRIX, DETERM_VAL)
End Select

Exit Function
ERROR_LABEL:
MATRIX_GS_SINGULAR_LINEAR_SYSTEM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GS_INTEGER_REDUCTION_FUNC
'DESCRIPTION   : Gauss-Jordan integer algorithm for matrix reduction
'LIBRARY       : MATRIX
'GROUP         : GAUSS_JORDAN
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_GS_INTEGER_REDUCTION_FUNC(ByRef XDATA_RNG As Variant, _
Optional ByRef YDATA_RNG As Variant, _
Optional ByVal epsilon As Double = 2E-16, _
Optional ByVal OUTPUT As Integer = 2)

'DATA_RNG: is a matrix (n x n); at the end contains the inverse of A
'DATA_VECTOR: is a matrix (n x m); at the end cotains the solution of AX=B

'Integer arithmetic is intrinsically more accurate for integer
'matrices but it may easily reaches the overflow. Use it only
'with integer matrices of low dimension.

'Epsilon sets the minimum round-off error; any
'value with an absolute value less than epsilon will
'be set to zero.


Dim i As Long
Dim j As Long
Dim k As Long

Dim AROWS As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim P_VAL As Double
Dim Q_VAL As Double

Dim MAX_VAL As Double
Dim MCM_VAL As Double
Dim DETERM_VAL As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim DETERM_FLAG As Boolean

On Error GoTo ERROR_LABEL

DETERM_FLAG = True
XDATA_MATRIX = XDATA_RNG

If IsArray(YDATA_RNG) = True Then
    YDATA_VECTOR = YDATA_RNG
    NCOLUMNS = UBound(YDATA_VECTOR, 2)
Else
    NCOLUMNS = 0
End If

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)
ReDim TEMP_VECTOR(1 To NROWS)

For i = 1 To NROWS 'initialization
    TEMP_MATRIX(i, i) = 1
    TEMP_VECTOR(i) = 1
Next i

For k = 1 To NROWS 'search max pivot
    AROWS = k
    MAX_VAL = 0
    For i = k To NROWS
        If Abs(XDATA_MATRIX(i, k)) > MAX_VAL Then
            AROWS = i
            MAX_VAL = Abs(XDATA_MATRIX(i, k))
        End If
    Next i
    
    ' swap rows
    If AROWS > k Then
        XDATA_MATRIX = MATRIX_SWAP_ROW_FUNC(XDATA_MATRIX, k, AROWS)
        TEMP_MATRIX = MATRIX_SWAP_ROW_FUNC(TEMP_MATRIX, k, AROWS)
        If NCOLUMNS > 0 Then _
            YDATA_VECTOR = MATRIX_SWAP_ROW_FUNC(YDATA_VECTOR, k, AROWS)
        If DETERM_FLAG Then TEMP_VECTOR(k) = -TEMP_VECTOR(k)
    End If
    
    ' check pivot 0
    If Abs(XDATA_MATRIX(k, k)) <= epsilon Then
        XDATA_MATRIX(k, k) = 0
        DETERM_VAL = 0
        GoTo 1983
    End If
    
    'integer linear reduction
    For i = 1 To NROWS
        If Abs(XDATA_MATRIX(i, k)) <= epsilon Then _
            XDATA_MATRIX(i, k) = 0 'mop-up Aik
        If i <> k And XDATA_MATRIX(i, k) <> 0 Then
            MCM_VAL = PAIR_MCM_FUNC(Abs(XDATA_MATRIX(k, k)), _
                Abs(XDATA_MATRIX(i, k)))
            
            Q_VAL = MCM_VAL / XDATA_MATRIX(k, k)
            P_VAL = -MCM_VAL / XDATA_MATRIX(i, k)
            
            If DETERM_FLAG Then TEMP_VECTOR(k) = TEMP_VECTOR(k) * P_VAL
            For j = 1 To NROWS
                XDATA_MATRIX(i, j) = P_VAL * XDATA_MATRIX(i, j) + _
                    Q_VAL * XDATA_MATRIX(k, j)
                TEMP_MATRIX(i, j) = P_VAL * TEMP_MATRIX(i, j) + Q_VAL * TEMP_MATRIX(k, j)
            Next j
            For j = 1 To NCOLUMNS
                YDATA_VECTOR(i, j) = P_VAL * YDATA_VECTOR(i, j) + _
                    Q_VAL * YDATA_VECTOR(k, j)
            Next j
        End If
    Next i
Next k

If DETERM_FLAG Then 'determinant computing
    DETERM_VAL = 1
    For i = 1 To NROWS
        DETERM_VAL = DETERM_VAL * (XDATA_MATRIX(i, i) / TEMP_VECTOR(i))
    Next i
End If

For i = 1 To NROWS 'normalization
    For j = 1 To NROWS
        TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) / XDATA_MATRIX(i, i)
    Next j
    For j = 1 To NCOLUMNS
        YDATA_VECTOR(i, j) = YDATA_VECTOR(i, j) / XDATA_MATRIX(i, i)
    Next j
Next i
    
1983:

Select Case OUTPUT
Case 0
    MATRIX_GS_INTEGER_REDUCTION_FUNC = TEMP_MATRIX ' INVERSE
Case 1
    MATRIX_GS_INTEGER_REDUCTION_FUNC = DETERM_VAL ' DETERMINANT
Case 2
    MATRIX_GS_INTEGER_REDUCTION_FUNC = YDATA_VECTOR
Case 3
    MATRIX_GS_INTEGER_REDUCTION_FUNC = XDATA_MATRIX
Case Else
    MATRIX_GS_INTEGER_REDUCTION_FUNC = Array(TEMP_MATRIX, DETERM_VAL, YDATA_VECTOR, XDATA_MATRIX)
End Select

Exit Function
ERROR_LABEL:
MATRIX_GS_INTEGER_REDUCTION_FUNC = Err.number
End Function

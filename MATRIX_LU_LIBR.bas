Attribute VB_Name = "MATRIX_LU_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_LU_FACT_CROUT_FUNC

'DESCRIPTION   : Perform the LU decomposition with Crout's algorithm with
'partial pivot.

'Note: LU decomposition without pivoting does not work if the
'first diagonal element is zero

'Note: when partial pivot is active (TRUE), the LU decomposition
'can refers to a permutation of A, That is, in generally, A <> LU
'In that case, the right decomposition formula is   A = P L U
'where P is a permutation matrix. This function returns also the
'permutation matrix in the last n columns

'LIBRARY       : MATRIX
'GROUP         : LU FACTORIZATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_LU_FACT_CROUT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal PIVOT_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 1)

'LU decomposition without pivoting does not work if the first
'diagonal element of A is zero

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
If NROWS <> NCOLUMNS Then GoTo ERROR_LABEL

'----------------------------------------------------------------

ReDim TEMP2_MATRIX(1 To NROWS, 1 To 1) 'permutation vector
For i = 1 To NROWS
    TEMP2_MATRIX(i, 1) = i
Next i

'------------------------start Crout's algorithm------------------------

For j = 1 To NROWS
    For i = 1 To j - 1
        TEMP1_VAL = DATA_MATRIX(i, j)
        For k = 1 To i - 1
            TEMP1_VAL = TEMP1_VAL - DATA_MATRIX(i, k) * DATA_MATRIX(k, j)
        Next k
        DATA_MATRIX(i, j) = TEMP1_VAL
    Next i
    
    TEMP3_VAL = 0
    For i = j To NROWS
        TEMP1_VAL = DATA_MATRIX(i, j)
        For k = 1 To j - 1
            TEMP1_VAL = TEMP1_VAL - DATA_MATRIX(i, k) * DATA_MATRIX(k, j)
        Next k
        DATA_MATRIX(i, j) = TEMP1_VAL
        TEMP2_VAL = Abs(TEMP1_VAL)
        If TEMP2_VAL >= TEMP3_VAL Then
            TEMP3_VAL = TEMP2_VAL
            h = i
        End If
    Next i
    
    If PIVOT_FLAG = True Then
        If j <> h Then
            For k = 1 To NROWS
                TEMP2_VAL = DATA_MATRIX(h, k)
                DATA_MATRIX(h, k) = DATA_MATRIX(j, k)
                DATA_MATRIX(j, k) = TEMP2_VAL
            Next k
            TEMP2_VAL = TEMP2_MATRIX(j, 1)
            TEMP2_MATRIX(j, 1) = TEMP2_MATRIX(h, 1)
            TEMP2_MATRIX(h, 1) = TEMP2_VAL
        End If
    End If
    
    If j <> NROWS And DATA_MATRIX(j, j) <> 0 Then
        TEMP2_VAL = 1 / DATA_MATRIX(j, j)
        For i = j + 1 To NROWS
            DATA_MATRIX(i, j) = DATA_MATRIX(i, j) * TEMP2_VAL
        Next i
    End If
Next j
        
        
'----------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------
Case 0
'----------------------------------------------------------------
    MATRIX_LU_FACT_CROUT_FUNC = DATA_MATRIX
'----------------------------------------------------------------
Case Else
'----------------------------------------------------------------
    ReDim TEMP1_MATRIX(1 To NROWS, 1 To 3 * NROWS)
    'stores L side and U side and TEMP2_MATRIX
    For i = 1 To NROWS
        For j = 1 To NROWS
            If j >= i Then
                TEMP1_MATRIX(i, NROWS + j) = DATA_MATRIX(i, j)
                TEMP1_MATRIX(i, j) = 0
            Else
                TEMP1_MATRIX(i, NROWS + j) = 0
                TEMP1_MATRIX(i, j) = DATA_MATRIX(i, j)
            End If
        Next j
        TEMP1_MATRIX(i, i) = 1
    Next i

    'store the permutation matrix
    TEMP2_MATRIX = MATRIX_PERMUTATION_FUNC(TEMP2_MATRIX)
    For i = 1 To NROWS
        For j = 1 To NROWS
            TEMP1_MATRIX(i, 2 * NROWS + j) = TEMP2_MATRIX(i, j)
        Next j
    Next i
    MATRIX_LU_FACT_CROUT_FUNC = TEMP1_MATRIX
'----------------------------------------------------------------
End Select
'----------------------------------------------------------------

Exit Function
ERROR_LABEL:
MATRIX_LU_FACT_CROUT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_LU_PRINT_FACT_FUNC
'DESCRIPTION   : Return the LU decomposition by segments (A = L, A = U, A = P)
'GROUP         : LU FACTORIZATION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_LU_PRINT_FACT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

DATA_MATRIX = MATRIX_LU_FACT_CROUT_FUNC(DATA_MATRIX, True, 1)

'-----------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------
Case 0 'matrix L
'-----------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)
    For i = 1 To NROWS
        For j = 1 To NROWS
            TEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
        Next j
    Next i
'-----------------------------------------------------------------------------
Case 1 'matrix U
'-----------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)
    For i = 1 To NROWS
        For j = 1 To NROWS
            TEMP_MATRIX(i, j) = DATA_MATRIX(i, NROWS + j)
        Next j
    Next i
'-----------------------------------------------------------------------------
Case Else 'matrix P
'-----------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)
    For i = 1 To NROWS
        For j = 1 To NROWS
            TEMP_MATRIX(i, j) = DATA_MATRIX(i, 2 * NCOLUMNS + j)
        Next j
    Next i
'-----------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------

MATRIX_LU_PRINT_FACT_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
MATRIX_LU_PRINT_FACT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_LU_FACT_FUNC
'DESCRIPTION   : Performs a procedure for decomposing a matrix into a product of a
'lower triangular matrix and an upper triangular matrix.

'GROUP         : LU FACTORIZATION
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_LU_FACT_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NSIZE As Long

Dim DD_VAL As Double
Dim TEMP_SUM As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim INDEX_ARR As Variant
Dim DETERM_ARR As Variant

Dim DATA_MATRIX As Variant
Dim LTEMP_MATRIX As Variant
Dim UTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL
  
DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX, 1)

ReDim LTEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
ReDim UTEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

ReDim INDEX_ARR(1 To NSIZE)
ReDim DETERM_ARR(1 To NSIZE)
  
'NR ludcmp
DD_VAL = 1
For i = 1 To NSIZE
    TEMP1_VAL = 0
    For j = 1 To NSIZE
        If Abs(DATA_MATRIX(i, j)) > TEMP1_VAL Then TEMP1_VAL = Abs(DATA_MATRIX(i, j))
    Next j
    'inverse of the absolute value of the largest element of the row
    DETERM_ARR(i) = 1 / TEMP1_VAL
Next i
For j = 1 To NSIZE
    For i = 1 To j - 1
        TEMP_SUM = DATA_MATRIX(i, j)
        For k = 1 To i - 1
            TEMP_SUM = TEMP_SUM - DATA_MATRIX(i, k) * DATA_MATRIX(k, j)
        Next k
        DATA_MATRIX(i, j) = TEMP_SUM
    Next i
    TEMP1_VAL = 0
    For i = j To NSIZE
        TEMP_SUM = DATA_MATRIX(i, j)
        For k = 1 To j - 1
            TEMP_SUM = TEMP_SUM - DATA_MATRIX(i, k) * DATA_MATRIX(k, j)
        Next k
        DATA_MATRIX(i, j) = TEMP_SUM
        TEMP2_VAL = DETERM_ARR(i) * Abs(TEMP_SUM)
        If TEMP2_VAL >= TEMP1_VAL Then
            l = i
            TEMP1_VAL = TEMP2_VAL
        End If
    Next i
    If j <> l Then
        'interchange row j and row l
        For k = 1 To NSIZE
            TEMP2_VAL = DATA_MATRIX(l, k)
            DATA_MATRIX(l, k) = DATA_MATRIX(j, k)
            DATA_MATRIX(j, k) = TEMP2_VAL
        Next k
        'change the parity of interchanges
        DD_VAL = -DD_VAL
        DETERM_ARR(l) = DETERM_ARR(j)
    End If
    INDEX_ARR(j) = l
    'If DATA_MATRIX(j, j) = 0 Then DATA_MATRIX(j, j) = tolerance
    If j <> NSIZE Then
        TEMP2_VAL = 1 / DATA_MATRIX(j, j)
        For i = j + 1 To NSIZE
            DATA_MATRIX(i, j) = DATA_MATRIX(i, j) * TEMP2_VAL
        Next i
    End If
Next j
For i = 1 To NSIZE
    For j = 1 To i - 1
        LTEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
    Next j
    LTEMP_MATRIX(i, i) = 1
Next i
For i = 1 To NSIZE
    For j = i To NSIZE
        UTEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
    Next j
Next i
MATRIX_LU_FACT_FUNC = UTEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_LU_FACT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_LU_LINEAR_SYSTEM_FUNC
'DESCRIPTION   : Solve a linear system of equations using LU Factorization
'LIBRARY       : MATRIX
'GROUP         : LU FACTORIZATION
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_LU_LINEAR_SYSTEM_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim DD_VAL As Variant
Dim TEMP_SUM As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim INDEX_ARR As Variant
Dim DETERM_ARR As Variant
Dim COEF_VECTOR As Variant

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR  As Variant

Dim NSIZE As Variant
Dim tolerance As Variant

On Error GoTo ERROR_LABEL
  
tolerance = 0.0000000000001

XDATA_MATRIX = XDATA_RNG
NSIZE = UBound(XDATA_MATRIX, 1)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

ReDim INDEX_ARR(1 To NSIZE)
ReDim DETERM_ARR(1 To NSIZE)
ReDim COEF_VECTOR(1 To NSIZE, 1 To 1)
  
DD_VAL = 1
For i = 1 To NSIZE
    TEMP1_VAL = 0
    For j = 1 To NSIZE
    If Abs(XDATA_MATRIX(i, j)) > TEMP1_VAL Then TEMP1_VAL = Abs(XDATA_MATRIX(i, j))
    Next j
    'inverse of the absolute value of the largest element of the row
    DETERM_ARR(i) = 1 / TEMP1_VAL
Next i
For j = 1 To NSIZE
    For i = 1 To j - 1
        TEMP_SUM = XDATA_MATRIX(i, j)
        For k = 1 To i - 1
            TEMP_SUM = TEMP_SUM - XDATA_MATRIX(i, k) * XDATA_MATRIX(k, j)
        Next k
        XDATA_MATRIX(i, j) = TEMP_SUM
    Next i
    TEMP1_VAL = 0
    For i = j To NSIZE
        TEMP_SUM = XDATA_MATRIX(i, j)
        For k = 1 To j - 1
            TEMP_SUM = TEMP_SUM - XDATA_MATRIX(i, k) * XDATA_MATRIX(k, j)
        Next k
        XDATA_MATRIX(i, j) = TEMP_SUM
        TEMP2_VAL = DETERM_ARR(i) * Abs(TEMP_SUM)
        If TEMP2_VAL >= TEMP1_VAL Then
            l = i
            TEMP1_VAL = TEMP2_VAL
        End If
    Next i
    If j <> l Then
        'interchange row j and row l
        For k = 1 To NSIZE
            TEMP2_VAL = XDATA_MATRIX(l, k)
            XDATA_MATRIX(l, k) = XDATA_MATRIX(j, k)
            XDATA_MATRIX(j, k) = TEMP2_VAL
        Next k
        'change the parity of interchanges
        DD_VAL = -DD_VAL
        DETERM_ARR(l) = DETERM_ARR(j)
    End If
    INDEX_ARR(j) = l
    If XDATA_MATRIX(j, j) = 0 Then XDATA_MATRIX(j, j) = tolerance
    If j <> NSIZE Then
        TEMP2_VAL = 1 / XDATA_MATRIX(j, j)
        For i = j + 1 To NSIZE
            XDATA_MATRIX(i, j) = XDATA_MATRIX(i, j) * TEMP2_VAL
        Next i
    End If
Next j

TEMP_SUM = 0
For i = 1 To NSIZE
    COEF_VECTOR(i, 1) = YDATA_VECTOR(i, 1)
Next i
ii = 0
For i = 1 To NSIZE
    jj = INDEX_ARR(i)
    TEMP_SUM = COEF_VECTOR(jj, 1)
    COEF_VECTOR(jj, 1) = COEF_VECTOR(i, 1)
    If ii <> 0 Then
        For j = ii To i - 1
            TEMP_SUM = TEMP_SUM - XDATA_MATRIX(i, j) * COEF_VECTOR(j, 1)
        Next j
    ElseIf TEMP_SUM <> 0 Then
        ii = i
    End If
    COEF_VECTOR(i, 1) = TEMP_SUM
Next i
For i = NSIZE To 1 Step -1
    TEMP_SUM = COEF_VECTOR(i, 1)
    For j = i + 1 To NSIZE
        TEMP_SUM = TEMP_SUM - XDATA_MATRIX(i, j) * COEF_VECTOR(j, 1)
    Next j
    COEF_VECTOR(i, 1) = TEMP_SUM / XDATA_MATRIX(i, i)
Next i

MATRIX_LU_LINEAR_SYSTEM_FUNC = COEF_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_LU_LINEAR_SYSTEM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_LU_INVERSE_FUNC
'DESCRIPTION   : CALCULATES THE INVERSE OF A MATRIX
'LIBRARY       : MATRIX
'GROUP         : LU FACTORIZATION
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_LU_INVERSE_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim NSIZE As Long

Dim DD_VAL As Double
Dim TEMP_SUM As Double
  
Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim INDEX_ARR As Variant
Dim DETERM_ARR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 0.0000000000001

DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX, 1)
  
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL

ReDim INDEX_ARR(1 To NSIZE)
ReDim DETERM_ARR(1 To NSIZE)
ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
  
DD_VAL = 1
For i = 1 To NSIZE
    TEMP1_VAL = 0
    For j = 1 To NSIZE
        If Abs(DATA_MATRIX(i, j)) > TEMP1_VAL Then TEMP1_VAL = Abs(DATA_MATRIX(i, j))
    Next j 'inverse of the absolute value of the largest element of the row
    DETERM_ARR(i) = 1 / TEMP1_VAL
Next i

For j = 1 To NSIZE
    For i = 1 To j - 1
        TEMP_SUM = DATA_MATRIX(i, j)
        For k = 1 To i - 1
            TEMP_SUM = TEMP_SUM - DATA_MATRIX(i, k) * DATA_MATRIX(k, j)
        Next k
        DATA_MATRIX(i, j) = TEMP_SUM
    Next i
    TEMP1_VAL = 0
    For i = j To NSIZE
        TEMP_SUM = DATA_MATRIX(i, j)
        For k = 1 To j - 1
            TEMP_SUM = TEMP_SUM - DATA_MATRIX(i, k) * DATA_MATRIX(k, j)
        Next k
        DATA_MATRIX(i, j) = TEMP_SUM
        TEMP2_VAL = DETERM_ARR(i) * Abs(TEMP_SUM)
        If TEMP2_VAL >= TEMP1_VAL Then
            l = i
            TEMP1_VAL = TEMP2_VAL
        End If
    Next i

    If j <> l Then
        'interchange row j and row l
        For k = 1 To NSIZE
            TEMP2_VAL = DATA_MATRIX(l, k)
            DATA_MATRIX(l, k) = DATA_MATRIX(j, k)
            DATA_MATRIX(j, k) = TEMP2_VAL
        Next k
        'change the parity of interchanges
        DD_VAL = -DD_VAL
        DETERM_ARR(l) = DETERM_ARR(j)
    End If

    INDEX_ARR(j) = l
    If DATA_MATRIX(j, j) = 0 Then DATA_MATRIX(j, j) = tolerance
    If j <> NSIZE Then
        TEMP2_VAL = 1 / DATA_MATRIX(j, j)
        For i = j + 1 To NSIZE
            DATA_MATRIX(i, j) = DATA_MATRIX(i, j) * TEMP2_VAL
        Next i
    End If
Next j

TEMP_SUM = 0
For i = 1 To NSIZE
    For j = 1 To NSIZE
        TEMP_MATRIX(i, j) = 0
    Next j
    TEMP_MATRIX(i, i) = 1
Next i

For k = 1 To NSIZE
    ii = 0
    For i = 1 To NSIZE
        jj = INDEX_ARR(i)
        TEMP_SUM = TEMP_MATRIX(jj, k)
        TEMP_MATRIX(jj, k) = TEMP_MATRIX(i, k)
        If ii <> 0 Then
            For j = ii To i - 1
                TEMP_SUM = TEMP_SUM - DATA_MATRIX(i, j) * TEMP_MATRIX(j, k)
            Next j
        ElseIf TEMP_SUM <> 0 Then
            ii = i
        End If
        TEMP_MATRIX(i, k) = TEMP_SUM
    Next i
    For i = NSIZE To 1 Step -1
        TEMP_SUM = TEMP_MATRIX(i, k)
        For j = i + 1 To NSIZE
            TEMP_SUM = TEMP_SUM - DATA_MATRIX(i, j) * TEMP_MATRIX(j, k)
        Next j
        TEMP_MATRIX(i, k) = TEMP_SUM / DATA_MATRIX(i, i)
    Next i
Next k

MATRIX_LU_INVERSE_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_LU_INVERSE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_LU_DETERM_FUNC
'DESCRIPTION   : Returns the matrix determinant of a symmetric matrix; using
'procedure for decomposing a matrix into a product of a lower
'triangular matrix and an upper triangular matrix.
'LIBRARY       : MATRIX
'GROUP         : LU FACTORIZATION
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_LU_DETERM_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NSIZE As Long

Dim DD_VAL As Double
Dim TEMP_SUM As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim INDEX_ARR As Variant
Dim DETERM_ARR As Variant

Dim LTEMP_MATRIX As Variant
Dim UTEMP_MATRIX As Variant

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL
  
DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX, 1)

ReDim INDEX_ARR(1 To NSIZE)
ReDim DETERM_ARR(1 To NSIZE)

ReDim LTEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
ReDim UTEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
    
'NR ludcmp
DD_VAL = 1
For i = 1 To NSIZE
    TEMP1_VAL = 0
    For j = 1 To NSIZE
        If Abs(DATA_MATRIX(i, j)) > TEMP1_VAL Then TEMP1_VAL = Abs(DATA_MATRIX(i, j))
    Next j
    'inverse of the absolute value of the largest element of the row
    DETERM_ARR(i) = 1 / TEMP1_VAL
Next i
For j = 1 To NSIZE
    For i = 1 To j - 1
        TEMP_SUM = DATA_MATRIX(i, j)
        For k = 1 To i - 1
            TEMP_SUM = TEMP_SUM - DATA_MATRIX(i, k) * DATA_MATRIX(k, j)
        Next k
        DATA_MATRIX(i, j) = TEMP_SUM
    Next i
    TEMP1_VAL = 0
    For i = j To NSIZE
        TEMP_SUM = DATA_MATRIX(i, j)
        For k = 1 To j - 1
            TEMP_SUM = TEMP_SUM - DATA_MATRIX(i, k) * DATA_MATRIX(k, j)
        Next k
        DATA_MATRIX(i, j) = TEMP_SUM
        TEMP2_VAL = DETERM_ARR(i) * Abs(TEMP_SUM)
        If TEMP2_VAL >= TEMP1_VAL Then
            l = i
            TEMP1_VAL = TEMP2_VAL
        End If
    Next i
    If j <> l Then
        'interchange row j and row l
        For k = 1 To NSIZE
            TEMP2_VAL = DATA_MATRIX(l, k)
            DATA_MATRIX(l, k) = DATA_MATRIX(j, k)
            DATA_MATRIX(j, k) = TEMP2_VAL
        Next k
        'change the parity of interchanges
        DD_VAL = -DD_VAL
        DETERM_ARR(l) = DETERM_ARR(j)
    End If
    INDEX_ARR(j) = l
    'If DATA_MATRIX(j, j) = 0 Then DATA_MATRIX(j, j) = tolerance
    If j <> NSIZE Then
        TEMP2_VAL = 1 / DATA_MATRIX(j, j)
        For i = j + 1 To NSIZE
            DATA_MATRIX(i, j) = DATA_MATRIX(i, j) * TEMP2_VAL
        Next i
    End If
Next j
For i = 1 To NSIZE
    For j = 1 To i - 1
        LTEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
    Next j
    LTEMP_MATRIX(i, i) = 1
Next i

For i = 1 To NSIZE
    For j = i To NSIZE
        UTEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
    Next j
Next i
For i = 1 To NSIZE
    DD_VAL = DD_VAL * UTEMP_MATRIX(i, i)
Next i

MATRIX_LU_DETERM_FUNC = DD_VAL
  
Exit Function
ERROR_LABEL:
MATRIX_LU_DETERM_FUNC = Err.number
End Function

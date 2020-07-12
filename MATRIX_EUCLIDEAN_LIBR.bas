Attribute VB_Name = "MATRIX_EUCLIDEAN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_EUCLIDEAN_NORM_FUNC

'DESCRIPTION   : Returns the euclidean norm of a vector
'|v| = (v1^2 + v2^2 + ...vn^2) ^ 1/2

'LIBRARY       : MATRIX
'GROUP         : EUCLIDEAN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_EUCLIDEAN_NORM_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim TEMP_SUM As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG

If IS_1D_ARRAY_FUNC(DATA_VECTOR) Then
    TEMP_SUM = 0
    For i = LBound(DATA_VECTOR) To UBound(DATA_VECTOR)
        TEMP_SUM = TEMP_SUM + DATA_VECTOR(i) ^ 2
    Next i
ElseIf IS_2D_ARRAY_FUNC(DATA_VECTOR) Then
    TEMP_SUM = 0
    For i = LBound(DATA_VECTOR, 1) To UBound(DATA_VECTOR, 1)
        TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1) ^ 2
    Next i
Else
    GoTo ERROR_LABEL
End If

VECTOR_EUCLIDEAN_NORM_FUNC = Sqr(TEMP_SUM)

Exit Function
ERROR_LABEL:
VECTOR_EUCLIDEAN_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_ABSOLUTE_NORM_FUNC
'DESCRIPTION   : Returns the norm of a vector (Absolute value)
'LIBRARY       : MATRIX
'GROUP         : EUCLIDEAN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_ABSOLUTE_NORM_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

'VERSION = 1      Absolute sum
'VERSION = Else  ( also infinite)  Maximum absolute

Dim i As Long
Dim SROW As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

SROW = LBound(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

Select Case VERSION
    Case 0
        TEMP_SUM = 0 '
        For i = SROW To NROWS
             TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1) ^ 2
        Next i
        TEMP_SUM = Sqr(TEMP_SUM)
    Case Else
        TEMP_SUM = 0 '
        For i = SROW To NROWS
            If Abs(DATA_VECTOR(i, 1)) > TEMP_SUM Then _
                TEMP_SUM = Abs(DATA_VECTOR(i, 1))
        Next i
End Select

For i = SROW To NROWS
    DATA_VECTOR(i, 1) = DATA_VECTOR(i, 1) / TEMP_SUM
Next i

VECTOR_ABSOLUTE_NORM_FUNC = DATA_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_ABSOLUTE_NORM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_NORMALIZED_VECTOR_FUNC

'DESCRIPTION   : Returns the normalized vectors of the matrix. The optional
'parameter VERSION indicates what normalization is performed. The optional
'parameter epsilon  sets the minimum error level (default 2E-14). Values
'under this level will be reset to zero.

'LIBRARY       : MATRIX
'GROUP         : EUCLIDEAN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_NORMALIZED_VECTOR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal epsilon As Double = 2 * 10 ^ -14)

'VERSION = 1.  All vector’s components are scaled to the min
'of the absolute values

'VERSION = 2   (default). All non-zero vectors are length = 1

'VERSION = 3. All vector’s components are scaled to the max of
'the absolute values

'VERSION = 4. All vector’s components are scaled to the mean of
'the absolute values

'VERSION = 5. All vector’s components are normalized respect to
'the mean and the standard deviation


Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Double
Dim TEMP_SUM As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

'VERSION = 1 (scaled to absolute max)
'2 (module=1)
'3 (scaled to absolute min)
'4 (scaled to absolute mean)
'5 (normalized mean = 0 and stdev = 1)

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

For j = SCOLUMN To NCOLUMNS
    For i = SROW To NROWS
        If Abs(DATA_MATRIX(i, j)) < epsilon Then
            DATA_MATRIX(i, j) = 0
        End If
    Next i
Next j

For j = SCOLUMN To NCOLUMNS
    Select Case VERSION
        Case 2  'module =1
            TEMP_SUM = 0 '
            For i = SROW To NROWS
                 TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j) ^ 2
            Next i
            TEMP_SUM = Sqr(TEMP_SUM)
        Case 3  'max of |vi|  =1
            TEMP_SUM = 0 '
            For i = SROW To NROWS
                If Abs(DATA_MATRIX(i, j)) > Abs(TEMP_SUM) Then _
                    TEMP_SUM = DATA_MATRIX(i, j)
            Next i
        Case 1  'min of |vi|  =1
            TEMP_SUM = 10 ^ 300
            For i = SROW To NROWS
                If Abs(DATA_MATRIX(i, j)) < Abs(TEMP_SUM) And _
                    Abs(DATA_MATRIX(i, j)) > 0 Then TEMP_SUM = DATA_MATRIX(i, j)
            Next i
        Case 4  'mean of |vi| element =1
            TEMP_SUM = 0
            For i = SROW To NROWS
                TEMP_SUM = TEMP_SUM + Abs(DATA_MATRIX(i, j))
            Next i
            TEMP_SUM = TEMP_SUM / (NROWS - SROW + 1)
        Case 5  'z=(x-NCOLUMNS)/s
            
            l = 0
            For k = SROW To NROWS
                TEMP_SUM = TEMP_SUM + DATA_MATRIX(k, j)
                l = l + 1
            Next k
            TEMP_VAL = TEMP_SUM / l 'mean
            
            l = 0
            TEMP_SUM = 0
            For k = SROW To NROWS
                TEMP_SUM = TEMP_SUM + (DATA_MATRIX(k, j) - TEMP_VAL) ^ 2
                l = l + 1
            Next
            TEMP_SUM = Sqr(TEMP_SUM / (l - 1)) 'sigma
            
            For i = SROW To NROWS
                DATA_MATRIX(i, j) = DATA_MATRIX(i, j) - TEMP_VAL
            Next i
        End Select
        If Abs(TEMP_SUM) > epsilon Then
            For i = SROW To NROWS
                DATA_MATRIX(i, j) = DATA_MATRIX(i, j) / TEMP_SUM
            Next i
        End If
Next j

MATRIX_NORMALIZED_VECTOR_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_NORMALIZED_VECTOR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_NORM_FUNC
'DESCRIPTION   : 'Returns the norm of a matrix or vector. The input can be a
'vector or a matrix; The optional parameter VERSION sets the specific norm to
'compute (default 2 for vectors, and 0 for matrices).

'For vectors
    'VERSION = 1    Absolute sum
    'VERSION = 2    Euclidean norm
    'VERSION = 3  ( also infinite)  Maximum absolute
'For matrices
    'VERSION = 0    Frobenius norm
    'VERSION = 1    Maximum absolute column sum
    'VERSION = 2    Euclidean norm
    'VERSION = 3 ( also infinite)   Maximum absolute row sum
    
'LIBRARY       : MATRIX
'GROUP         : EUCLIDEAN
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_NORM_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim EIGEN_MAX_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
BTEMP_VAL = 0
    
    Select Case VERSION
        Case 0
            GoSub 1983
        Case 1
            For j = 1 To UBound(DATA_MATRIX, 2)
                ATEMP_VAL = 0
                For i = 1 To UBound(DATA_MATRIX, 1)
                    ATEMP_VAL = ATEMP_VAL + Abs(DATA_MATRIX(i, j))
                Next i
                If BTEMP_VAL < ATEMP_VAL Then BTEMP_VAL = ATEMP_VAL
            Next j
        Case 2
            If UBound(DATA_MATRIX, 1) > 1 And UBound(DATA_MATRIX, 2) > 1 Then
                TEMP_MATRIX = MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(DATA_MATRIX), DATA_MATRIX, 70)
                EIGEN_MAX_VAL = MATRIX_DOMINANT_EIGEN_FUNC(TEMP_MATRIX, False, 1000, 0, 10 ^ -14)
                BTEMP_VAL = (EIGEN_MAX_VAL) ^ 0.5
            Else
                GoSub 1983
            End If
        Case 3
            For i = 1 To UBound(DATA_MATRIX, 1)
                ATEMP_VAL = 0
                For j = 1 To UBound(DATA_MATRIX, 2)
                    ATEMP_VAL = ATEMP_VAL + Abs(DATA_MATRIX(i, j))
                Next j
                If BTEMP_VAL < ATEMP_VAL Then BTEMP_VAL = ATEMP_VAL
            Next i
    End Select

    MATRIX_NORM_FUNC = BTEMP_VAL

Exit Function
'------------------------------------------------------------------------------
1983:
'------------------------------------------------------------------------------
    For i = 1 To UBound(DATA_MATRIX, 1)
        For j = 1 To UBound(DATA_MATRIX, 2)
            BTEMP_VAL = BTEMP_VAL + DATA_MATRIX(i, j) ^ 2
        Next j
    Next i
    BTEMP_VAL = (BTEMP_VAL) ^ 0.5
'------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------
ERROR_LABEL:
MATRIX_NORM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_EUCLIDEAN_NORM_FUNC
'DESCRIPTION   : Absolute of matrix (Euclidean norm)
'LIBRARY       : MATRIX
'GROUP         : EUCLIDEAN
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_EUCLIDEAN_NORM_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

TEMP_SUM = 0
For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j) ^ 2
    Next j
Next i

MATRIX_EUCLIDEAN_NORM_FUNC = Sqr(TEMP_SUM)

Exit Function
ERROR_LABEL:
MATRIX_EUCLIDEAN_NORM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ABSOLUTE_FUNC
'DESCRIPTION   : Returns the M = abs(B)
'LIBRARY       : MATRIX
'GROUP         : EUCLIDEAN
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ABSOLUTE_FUNC(ByRef DATA_RNG As Variant)
  
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
          TEMP_MATRIX(i, j) = Abs(DATA_MATRIX(i, j))
    Next j
  Next i
  
  MATRIX_ABSOLUTE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_ABSOLUTE_FUNC = Err.number
End Function

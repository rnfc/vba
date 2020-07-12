Attribute VB_Name = "MATRIX_ORTHOGONAL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GRAM_SCHMIDT_ORTHOGONAL_ROTATION_FUNC

'DESCRIPTION   : 'Returns orthogonal matrix from a set of independent vectors
'uses Modified Gram-Schmidt (Mayer) algorithm

'This popular method is used to build an orthogonal-normalized
'base from a set of n independent vectors.
'Developing this algorithm, we see that the vector k is built
'from all k-1 previous vectors

'At the end, all vectors of the bases U will be orthogonal and normalized.
'This method is very straightforward, but it is also very sensitive to
'the round-off error. This happens because the error propagate itself
'along the vectors from 1 to n

'LIBRARY       : MATRIX
'GROUP         : ROTATION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_GRAM_SCHMIDT_ORTHOGONAL_ROTATION_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim TEMP_ARR As Variant
Dim DATA_MATRIX As Variant

Dim LAMBDA As Double
Dim tolerance As Double
Dim epsilon As Double

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)


tolerance = 10 ^ -14
epsilon = 10 ^ -7

ReDim TEMP_ARR(1 To NCOLUMNS) 'compute the vectors-column norm
For k = 1 To NCOLUMNS
    For i = 1 To NROWS
        TEMP_ARR(k) = TEMP_ARR(k) + DATA_MATRIX(i, k) ^ 2
    Next i
    TEMP_ARR(k) = Sqr(TEMP_ARR(k))
Next k
    
For k = 1 To NCOLUMNS
    ATEMP_VAL = 0
    For i = 1 To NROWS
        ATEMP_VAL = ATEMP_VAL + DATA_MATRIX(i, k) ^ 2
    Next i
    ATEMP_VAL = Sqr(ATEMP_VAL) 'tolerance setting
    LAMBDA = tolerance * TEMP_ARR(k)
    If LAMBDA > epsilon Then LAMBDA = epsilon
    If ATEMP_VAL > LAMBDA Then 'normalize only if |aik| > lambda
        For i = 1 To NROWS
            DATA_MATRIX(i, k) = DATA_MATRIX(i, k) / ATEMP_VAL
        Next i
    End If
    For j = k + 1 To NCOLUMNS
        BTEMP_VAL = MATRIX_ELEMENTS_PRODUCT_SUM_FUNC(DATA_MATRIX, k, j)
        For i = 1 To NROWS
            DATA_MATRIX(i, j) = DATA_MATRIX(i, j) - _
                BTEMP_VAL * DATA_MATRIX(i, k)
            If Abs(DATA_MATRIX(i, j)) < 0.2 * tolerance Then DATA_MATRIX(i, j) = 0
        Next i
    Next j
Next k

MATRIX_GRAM_SCHMIDT_ORTHOGONAL_ROTATION_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_GRAM_SCHMIDT_ORTHOGONAL_ROTATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_KAISER_VARIMAX_ORTHOGONAL_ROTATION_FUNC

'DESCRIPTION   : Returns the rotate factor loading matrix with the Kaiser's
'Varimax criterion

'This function computes the orthogonal rotation for a Factor Loading
'matrix using the Kaiser Varimax method for 2D and 3D factors
'Parameter FL is the Factor Loading matrix to rotate (NROWS x
'NCOLUMNS). The number of factors NCOLUMNS, at this release,
'can only be 2 or 3. Optional parameter NORM_FLAG = True/False
'chooses the Varimax normalized criterion. That is, it indicates
'if the matrix of loading is to be row-normalized before rotation.
'Optional parameter MaxErr set sets the accuracy required.
'The algorithm stops when the absolute difference of two consecutive
'Varimax values is less than epsilon.

'Optional parameter nLOOPS sets the maximum number of iterations allowed.

'LIBRARY       : MATRIX
'GROUP         : ROTATION
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_KAISER_VARIMAX_ORTHOGONAL_ROTATION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NORM_FLAG As Boolean = False, _
Optional ByVal epsilon As Double = 10 ^ -4, _
Optional ByVal nLOOPS As Long = 500)

'DATA_RNG: factor loading matrix (NROWS x NCOLUMNS) - with NCOLUMNS<4

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TETA_VAL As Double
Dim MAX_VAL As Double
Dim DELTA_VAL As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim DATA_MATRIX As Variant

Dim LAMBDA As Double

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
If NCOLUMNS > 3 Then GoTo ERROR_LABEL 'only 2 and 3 factor for now

'Given a No.Points × No.Dimensions configuration,
'the procedure tries to find an orthonormal rotation matrix T such
'that the sum of variances of the columns of B*B is a maximum,
'where B = AT and * is the element wise product of
'matrices. A direct solution for the optimal T is not available,
'except for the case when numberOfDimensions equals two. Kaiser
'suggested an iterative algorithm based on planar rotations, i.e.,
'alternate rotations of all pairs of columns of A.

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
If NCOLUMNS < 2 Then GoTo ERROR_LABEL  'few factors

DELTA_VAL = nLOOPS
BTEMP_VAL = MATRIX_KAISER_VARIMAX_INDEX_FUNC(DATA_MATRIX, NORM_FLAG)

k = 0
LAMBDA = 0.5

Do
    For i = 1 To NCOLUMNS - 1
        For j = i + 1 To NCOLUMNS
          MAX_VAL = BTEMP_VAL
          TETA_VAL = LAMBDA
          ATEMP_VAL = BTEMP_VAL
            Do
                DATA_MATRIX = MATRIX_FAST_ROTATION_FUNC(DATA_MATRIX, TETA_VAL, i, j)
                BTEMP_VAL = MATRIX_KAISER_VARIMAX_INDEX_FUNC(DATA_MATRIX, NORM_FLAG)
                If BTEMP_VAL < ATEMP_VAL Then TETA_VAL = -TETA_VAL / 2
                If Abs(BTEMP_VAL - ATEMP_VAL) < epsilon Then GoTo 1983
                ATEMP_VAL = BTEMP_VAL
                k = k + 1
            Loop Until k > nLOOPS
1983:
        Next j
    Next i
    DELTA_VAL = Abs(BTEMP_VAL - MAX_VAL)
    If DELTA_VAL < epsilon Then Exit Do
    LAMBDA = DELTA_VAL
Loop Until k > nLOOPS

For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        DATA_MATRIX(i, j) = -DATA_MATRIX(i, j)
    Next j
Next i

If k < nLOOPS Then
      MATRIX_KAISER_VARIMAX_ORTHOGONAL_ROTATION_FUNC = DATA_MATRIX
Else: GoTo ERROR_LABEL
End If
      
Exit Function
ERROR_LABEL:
MATRIX_KAISER_VARIMAX_ORTHOGONAL_ROTATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_KAISER_VARIMAX_INDEX_FUNC

'DESCRIPTION   : 'Returns the Varimax value of a given Factor matrix
'Varimax is a popular criterion to perform orthogonal rotation
'of Factors Loading matrices. Usually, the rotation stops when
'Varimax is maximized.

'Optional parameter NORM_FLAG = True/False indicates if the matrix
'is to be row normalized before computing

'LIBRARY       : MATRIX
'GROUP         : ROTATION
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_KAISER_VARIMAX_INDEX_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NORM_FLAG As Boolean = True)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double
Dim CTEMP_SUM As Double
Dim DTEMP_SUM As Double

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If NORM_FLAG Then
    For i = 1 To NROWS
        ATEMP_SUM = 0
        For j = 1 To NCOLUMNS
            ATEMP_SUM = ATEMP_SUM + DATA_MATRIX(i, j) ^ 2
        Next j
        ATEMP_SUM = Sqr(ATEMP_SUM)
        For j = 1 To NCOLUMNS
            DATA_MATRIX(i, j) = DATA_MATRIX(i, j) / ATEMP_SUM
        Next j
    Next i
End If

BTEMP_SUM = 0
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        BTEMP_SUM = BTEMP_SUM + DATA_MATRIX(i, j) ^ 4
    Next i
Next j

DTEMP_SUM = 0
For j = 1 To NCOLUMNS
    CTEMP_SUM = 0
    For i = 1 To NROWS
        CTEMP_SUM = CTEMP_SUM + DATA_MATRIX(i, j) ^ 2
    Next i
    DTEMP_SUM = DTEMP_SUM + CTEMP_SUM ^ 2
Next j

MATRIX_KAISER_VARIMAX_INDEX_FUNC = BTEMP_SUM - DTEMP_SUM / NROWS

Exit Function
ERROR_LABEL:
MATRIX_KAISER_VARIMAX_INDEX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ORTHOGONAL_FUNC

'DESCRIPTION   : Returns the orthogonal matrix (NSIZE x NSIZE) that
'performs the planar rotation over the plane defined by axes ii
'and jj. Parameter theta sets the angle of rotation in radians
'Parameters ii and jj are the columns of the rotation
'Note that all rotation matrices have determinant = 1

'LIBRARY       : MATRIX
'GROUP         : ROTATION
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ORTHOGONAL_FUNC(ByVal NSIZE As Long, _
ByVal TETA_VAL As Double, _
ByVal ii As Long, _
ByVal jj As Long)

'ii = Lower Bound
'jj = Upper Bound

Dim i As Long
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
For i = 1 To NSIZE
    TEMP_MATRIX(i, i) = 1 'diagonal
Next i

If ii < 0 Or ii > NSIZE Or _
    jj < 0 Or jj > NSIZE Or _
    ii = jj Then GoTo ERROR_LABEL

TEMP_MATRIX(ii, ii) = Cos(TETA_VAL)
TEMP_MATRIX(ii, jj) = -Sin(TETA_VAL)
TEMP_MATRIX(jj, ii) = Sin(TETA_VAL)
TEMP_MATRIX(jj, jj) = Cos(TETA_VAL)

MATRIX_ORTHOGONAL_FUNC = TEMP_MATRIX
      
Exit Function
ERROR_LABEL:
MATRIX_ORTHOGONAL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_FAST_ROTATION_FUNC
'DESCRIPTION   : Fast matrix rotation
'LIBRARY       : MATRIX
'GROUP         : ROTATION
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_FAST_ROTATION_FUNC(ByRef DATA_RNG As Variant, _
ByVal TETA_VAL As Double, _
ByVal ii As Long, _
ByVal jj As Long)

Dim i As Long
Dim NSIZE As Long

Dim COS_VAL As Double
Dim SIN_VAL As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NSIZE = UBound(DATA_MATRIX)
COS_VAL = Cos(TETA_VAL)
SIN_VAL = Sin(TETA_VAL)

For i = 1 To NSIZE
    ATEMP_VAL = DATA_MATRIX(i, ii)
    BTEMP_VAL = DATA_MATRIX(i, jj)
    DATA_MATRIX(i, ii) = COS_VAL * ATEMP_VAL + SIN_VAL * BTEMP_VAL
    DATA_MATRIX(i, jj) = -SIN_VAL * ATEMP_VAL + COS_VAL * BTEMP_VAL
Next i

MATRIX_FAST_ROTATION_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_FAST_ROTATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_JACOBI_ORTHOGONAL_ROTATION_FUNC

'DESCRIPTION   : This function returns Jacobi orthogonal rotation matrix of a
'given symmetric matrix. This function search for the max absolute
'values out of the first diagonal and generates an orthogonal matrix
'in order to reduce it to zero by similarity transformation

'LIBRARY       : MATRIX
'GROUP         : ROTATION
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_JACOBI_ORTHOGONAL_ROTATION_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double

Dim DATA_MATRIX As Variant
Dim UTEMP_MATRIX As Variant
Dim WTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
'search for max value out of the first diagonal
    
ATEMP_VAL = 0
ii = 0
jj = 0
For i = 1 To NROWS
    For j = 1 To NROWS
        If i <> j And Abs(DATA_MATRIX(i, j)) > ATEMP_VAL Then
            ATEMP_VAL = Abs(DATA_MATRIX(i, j))
            ii = i
            jj = j
        End If
    Next j
Next i

BTEMP_VAL = DATA_MATRIX(jj, jj) - DATA_MATRIX(ii, ii)
If BTEMP_VAL = 0 Then
    CTEMP_VAL = 1
Else
    DTEMP_VAL = BTEMP_VAL / DATA_MATRIX(ii, jj) / 2
    CTEMP_VAL = Sgn(DTEMP_VAL) / (Abs(DTEMP_VAL) + Sqr(DTEMP_VAL ^ 2 + 1))
End If
ETEMP_VAL = 1 / Sqr(CTEMP_VAL ^ 2 + 1)
FTEMP_VAL = CTEMP_VAL * ETEMP_VAL
ReDim UTEMP_MATRIX(1 To NROWS, 1 To NROWS)
ReDim WTEMP_MATRIX(1 To NROWS, 1 To NROWS)

For i = 1 To NROWS
    UTEMP_MATRIX(i, i) = 1
    WTEMP_MATRIX(i, i) = 1
Next i

UTEMP_MATRIX(ii, ii) = ETEMP_VAL
UTEMP_MATRIX(ii, jj) = FTEMP_VAL
UTEMP_MATRIX(jj, ii) = -FTEMP_VAL
UTEMP_MATRIX(jj, jj) = ETEMP_VAL
WTEMP_MATRIX(ii, ii) = ETEMP_VAL
WTEMP_MATRIX(ii, jj) = -FTEMP_VAL
WTEMP_MATRIX(jj, ii) = FTEMP_VAL
WTEMP_MATRIX(jj, jj) = ETEMP_VAL

DATA_MATRIX = MMULT_FUNC(DATA_MATRIX, UTEMP_MATRIX, 70)
DATA_MATRIX = MMULT_FUNC(WTEMP_MATRIX, DATA_MATRIX, 70)
MATRIX_JACOBI_ORTHOGONAL_ROTATION_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_JACOBI_ORTHOGONAL_ROTATION_FUNC = Err.number
End Function

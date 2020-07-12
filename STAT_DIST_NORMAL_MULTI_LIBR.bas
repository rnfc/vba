Attribute VB_Name = "STAT_DIST_NORMAL_MULTI_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_RANDOM_NORMAL_FUNC
'DESCRIPTION   : Correlated Random Number Generator
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_MULTI
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function BIVAR_RANDOM_NORMAL_FUNC(ByVal NRV1_VAL As Double, _
ByVal NRV2_VAL As Double, _
ByVal RHO_VAL As Double)

On Error GoTo ERROR_LABEL

BIVAR_RANDOM_NORMAL_FUNC = NRV1_VAL * RHO_VAL + (1 - RHO_VAL ^ 2) ^ 0.5 * NRV2_VAL

Exit Function
ERROR_LABEL:
BIVAR_RANDOM_NORMAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_NORMAL_RANDOM_DRAW_FUNC
'DESCRIPTION   : Returns one draw from binormal distributions
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_MULTI
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function BIVAR_NORMAL_RANDOM_DRAW_FUNC(ByVal MEAN1_VAL As Double, _
ByVal SIGMA1_VAL As Double, _
ByVal MEAN2_VAL As Double, _
ByVal SIGMA2_VAL As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal RANDOM_FLAG As Boolean = True)

Dim j As Long
Dim k As Long

Dim TEMP_SUM As Double
Dim TEMP_STR As String

Dim MEAN_VECTOR As Variant
Dim SIGMA_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim RANDOM_VECTOR As Variant
Dim CHOLESKY_MATRIX As Variant

On Error GoTo ERROR_LABEL

If RANDOM_FLAG = True Then: Randomize

ReDim TEMP_VECTOR(1 To 2, 1 To 1)
ReDim SIGMA_VECTOR(1 To 2, 1 To 1)
ReDim MEAN_VECTOR(1 To 2, 1 To 1)
ReDim CHOLESKY_MATRIX(1 To 2, 1 To 2)

CHOLESKY_MATRIX(1, 1) = 1
CHOLESKY_MATRIX(2, 1) = RHO_VAL

CHOLESKY_MATRIX(1, 2) = RHO_VAL
CHOLESKY_MATRIX(2, 2) = 1

SIGMA_VECTOR(1, 1) = SIGMA1_VAL
MEAN_VECTOR(1, 1) = MEAN1_VAL

SIGMA_VECTOR(2, 1) = SIGMA2_VAL
MEAN_VECTOR(2, 1) = MEAN2_VAL

If (SIGMA1_VAL < 0) Or (SIGMA2_VAL < 0) Then
    TEMP_STR = MULTI_NORMAL_CORREL_ERROR_FUNC(1)
    BIVAR_NORMAL_RANDOM_DRAW_FUNC = TEMP_STR
    Exit Function
ElseIf Abs(RHO_VAL) > 1 Then
    TEMP_STR = MULTI_NORMAL_CORREL_ERROR_FUNC(0)
    BIVAR_NORMAL_RANDOM_DRAW_FUNC = TEMP_STR
    Exit Function
End If

CHOLESKY_MATRIX = MATRIX_CHOLESKY_FUNC(CHOLESKY_MATRIX)
RANDOM_VECTOR = MATRIX_RANDOM_NORMAL_FUNC(1, 2, 0, 0, 1, 0)

For j = 1 To 2
    TEMP_SUM = 0
    For k = 1 To j
        TEMP_SUM = TEMP_SUM + CHOLESKY_MATRIX(j, k) * RANDOM_VECTOR(1, k)
    Next k
    TEMP_VECTOR(j, 1) = TEMP_SUM * SIGMA_VECTOR(j, 1) + MEAN_VECTOR(j, 1)
Next j

BIVAR_NORMAL_RANDOM_DRAW_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
BIVAR_NORMAL_RANDOM_DRAW_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTI_NORMAL_RANDOM_DRAW_FUNC
'DESCRIPTION   : Returns one draw from Multi-normal distributions
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_MULTI
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MULTI_NORMAL_RANDOM_DRAW_FUNC(ByRef CORREL_RNG As Variant, _
ByRef MEAN_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
Optional ByVal RANDOM_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long

Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NCOLUMNS As Long

Dim TEMP_STR As String
Dim TEMP_SUM As Double

Dim MEAN_VECTOR As Variant
Dim SIGMA_VECTOR As Variant
Dim CORRELATION_MATRIX As Variant

Dim TEMP_VECTOR As Variant
Dim RANDOM_VECTOR As Variant
Dim CHOLESKY_MATRIX As Variant

On Error GoTo ERROR_LABEL

If RANDOM_FLAG = True Then: Randomize

CORRELATION_MATRIX = CORREL_RNG
NCOLUMNS = UBound(CORRELATION_MATRIX, 2)

MEAN_VECTOR = MEAN_RNG
If UBound(MEAN_VECTOR, 1) = 1 Then
    MEAN_VECTOR = MATRIX_TRANSPOSE_FUNC(MEAN_VECTOR)
End If
If UBound(MEAN_VECTOR, 1) <> NCOLUMNS Then: GoTo ERROR_LABEL

SIGMA_VECTOR = SIGMA_RNG
If UBound(SIGMA_VECTOR, 1) = 1 Then
    SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)
End If
If UBound(SIGMA_VECTOR, 1) <> NCOLUMNS Then: GoTo ERROR_LABEL

hh = UBound(MEAN_VECTOR, 1)
kk = UBound(SIGMA_VECTOR, 1)
ii = UBound(CORRELATION_MATRIX, 1)
jj = UBound(CORRELATION_MATRIX, 2)

If hh <> NCOLUMNS Then
    TEMP_STR = MULTI_NORMAL_CORREL_ERROR_FUNC(2)
    MULTI_NORMAL_RANDOM_DRAW_FUNC = TEMP_STR
    Exit Function
ElseIf kk <> NCOLUMNS Then
    TEMP_STR = MULTI_NORMAL_CORREL_ERROR_FUNC(3)
    MULTI_NORMAL_RANDOM_DRAW_FUNC = TEMP_STR
    Exit Function
ElseIf ii <> NCOLUMNS Then
    TEMP_STR = MULTI_NORMAL_CORREL_ERROR_FUNC(4)
    MULTI_NORMAL_RANDOM_DRAW_FUNC = TEMP_STR
    Exit Function
ElseIf jj <> NCOLUMNS Then
    TEMP_STR = MULTI_NORMAL_CORREL_ERROR_FUNC(5)
    MULTI_NORMAL_RANDOM_DRAW_FUNC = TEMP_STR
    Exit Function
End If

For i = 1 To NCOLUMNS
    If SIGMA_VECTOR(i, 1) < 0 Then
        MULTI_NORMAL_RANDOM_DRAW_FUNC = MULTI_NORMAL_CORREL_ERROR_FUNC(1)
        Exit Function
    End If
    
    For j = 1 To NCOLUMNS
            If Abs(CORRELATION_MATRIX(i, j)) > 1 Then
                TEMP_STR = MULTI_NORMAL_CORREL_ERROR_FUNC(7)
                MULTI_NORMAL_RANDOM_DRAW_FUNC = TEMP_STR
                Exit Function
            End If
        If i = j Then
            If CORRELATION_MATRIX(i, j) <> 1 Then
                TEMP_STR = MULTI_NORMAL_CORREL_ERROR_FUNC(8)
                MULTI_NORMAL_RANDOM_DRAW_FUNC = TEMP_STR
                Exit Function
            End If
        End If
    Next j
Next i

ReDim CHOLESKY_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)

For i = 1 To NCOLUMNS
    For j = 1 To i
        CHOLESKY_MATRIX(i, j) = CORRELATION_MATRIX(i, j)
        CHOLESKY_MATRIX(j, i) = CHOLESKY_MATRIX(i, j)
    Next j
Next i

CHOLESKY_MATRIX = MATRIX_CHOLESKY_FUNC(CHOLESKY_MATRIX)
RANDOM_VECTOR = MATRIX_RANDOM_NORMAL_FUNC(1, NCOLUMNS, 0, 0, 1, 0)

For j = 1 To NCOLUMNS
    TEMP_SUM = 0
    For k = 1 To j
        TEMP_SUM = TEMP_SUM + CHOLESKY_MATRIX(j, k) * RANDOM_VECTOR(1, k)
    Next k
    TEMP_VECTOR(1, j) = TEMP_SUM * SIGMA_VECTOR(j, 1) + MEAN_VECTOR(j, 1)
Next j

MULTI_NORMAL_RANDOM_DRAW_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MULTI_NORMAL_RANDOM_DRAW_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTI_NORMAL_EIGEN_MC_FUNC
'DESCRIPTION   : Computation of Multivariate Standard Normal Distribution
'probability; Using EIGEN VECTORS AND EIGEN VALUES: Triadiagonal
'reductional form

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_MULTI
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MULTI_NORMAL_EIGEN_MC_FUNC(ByVal nLOOPS As Long, _
ByRef CORREL_RNG As Variant, _
ByRef ZSCORE_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal RANDOM_FLAG As Boolean = True)

'With ZSCORE_RNG, assume that we have 3 variables that are
'distributed normally with mean of 0 and standard deviation
'of 1 with its respective correlations

'We want to find the probability that variable1 <= 0, variable
'2 <= 1.5, and variable 3 <= 2.5: The JOINT PROBABILITY

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim NSIZE As Long

Dim TEMP_SUM As Double
Dim TEMP_VAL As Double

Dim ZSCORE_VECTOR As Variant

Dim TEMP_MATRIX As Variant

Dim EIGEN_VALUES As Variant
Dim EIGEN_VECTOR As Variant

Dim RANDOM_MATRIX As Variant
Dim CORRELATION_MATRIX As Variant

On Error GoTo ERROR_LABEL

CORRELATION_MATRIX = CORREL_RNG
If UBound(CORRELATION_MATRIX, 1) <> UBound(CORRELATION_MATRIX, 2) Then: GoTo ERROR_LABEL
NSIZE = UBound(CORRELATION_MATRIX, 1)

ZSCORE_VECTOR = ZSCORE_RNG
If UBound(ZSCORE_VECTOR, 2) = 1 Then
    ZSCORE_VECTOR = MATRIX_TRANSPOSE_FUNC(ZSCORE_VECTOR)
End If
If UBound(ZSCORE_VECTOR, 2) <> NSIZE Then: GoTo ERROR_LABEL

For i = 1 To NSIZE
    For j = 1 To i
        CORRELATION_MATRIX(j, i) = CORRELATION_MATRIX(i, j)
    Next j
Next i
    
EIGEN_VECTOR = MATRIX_PCA_FUNC(CORRELATION_MATRIX, False, 0)
EIGEN_VALUES = MATRIX_PCA_FUNC(CORRELATION_MATRIX, False, 1)
    
'** Form diagonal matrix using the Eigenvalue
ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
For i = 1 To NSIZE
    TEMP_MATRIX(i, i) = Sqr(Abs(EIGEN_VALUES(i, 1)))
Next i

'** Transpose the Eigenvector matrix
CORRELATION_MATRIX = MATRIX_TRANSPOSE_FUNC(EIGEN_VECTOR)
'** Multiply the transpose the Eigenvector matrix by the diagonal matrix using the Eigenvalue
CORRELATION_MATRIX = MMULT_FUNC(TEMP_MATRIX, CORRELATION_MATRIX, 70)
'** Generate random numbers from multivariant standard normal distribution and
'   compute the probability
        
l = 0
ReDim TEMP_MATRIX(1 To nLOOPS, 1 To NSIZE)
        
If RANDOM_FLAG = True Then: Randomize
RANDOM_MATRIX = MATRIX_RANDOM_NORMAL_FUNC(NSIZE, nLOOPS, 0, 0, 1, 0)

For k = 1 To nLOOPS
    TEMP_SUM = 0
    For j = 1 To NSIZE
        TEMP_MATRIX(k, j) = 0
        For i = 1 To NSIZE
            TEMP_MATRIX(k, j) = TEMP_MATRIX(k, j) + CORRELATION_MATRIX(i, j) * RANDOM_MATRIX(i, k)
        Next i
        If TEMP_MATRIX(k, j) < ZSCORE_VECTOR(1, j) Then TEMP_SUM = TEMP_SUM + 1
    Next j
    If TEMP_SUM = NSIZE Then l = l + 1
Next k
        
TEMP_VAL = l / nLOOPS
Select Case OUTPUT
Case 0
    MULTI_NORMAL_EIGEN_MC_FUNC = MULTI_NORMAL_CORREL_VOLATILITY_FUNC(TEMP_MATRIX)
Case Else
    MULTI_NORMAL_EIGEN_MC_FUNC = TEMP_VAL 'JOINT PROB DIST
End Select

Exit Function
ERROR_LABEL:
MULTI_NORMAL_EIGEN_MC_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTI_NORMAL_CORREL_SIMUL_FUNC
'DESCRIPTION   : Runs a Multivariate Normal Simulation
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_MULTI
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MULTI_NORMAL_CORREL_SIMUL_FUNC(ByRef CORREL_RNG As Variant, _
ByVal nLOOPS As Long, _
Optional ByVal NORM_TYPE As Integer = 0, _
Optional ByVal RANDOM_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim TEMP_MATRIX As Variant
Dim RANDOM_MATRIX As Variant
Dim CHOLESKY_MATRIX As Variant

On Error GoTo ERROR_LABEL

If RANDOM_FLAG = True Then: Randomize

CHOLESKY_MATRIX = MATRIX_CHOLESKY_FUNC(CORREL_RNG)
If UBound(CHOLESKY_MATRIX, 1) <> UBound(CHOLESKY_MATRIX, 2) Then: GoTo ERROR_LABEL
NSIZE = UBound(CHOLESKY_MATRIX, 2)

RANDOM_MATRIX = MATRIX_RANDOM_NORMAL_FUNC(nLOOPS, NSIZE, NORM_TYPE, 0, 1, 0)

ReDim TEMP_MATRIX(1 To nLOOPS, 1 To NSIZE)
For i = 1 To nLOOPS
    For j = 1 To NSIZE
          For k = 1 To j
                TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + RANDOM_MATRIX(i, k) * CHOLESKY_MATRIX(j, k)
          Next k
    Next j
Next i

MULTI_NORMAL_CORREL_SIMUL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MULTI_NORMAL_CORREL_SIMUL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTI_NORMAL_SIMULATION_FUNC
'DESCRIPTION   : Runs a Multivariate Normal Simulation
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_MULTI
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MULTI_NORMAL_SIMULATION_FUNC(ByRef CORREL_RNG As Variant, _
ByRef MEAN_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
Optional ByVal nLOOPS As Long = 10, _
Optional ByVal NORM_TYPE As Integer = 0, _
Optional ByVal RANDOM_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant

Dim MEAN_VECTOR As Variant
Dim SIGMA_VECTOR As Variant

Dim RANDOM_MATRIX As Variant
Dim CHOLESKY_MATRIX As Variant

On Error GoTo ERROR_LABEL

If RANDOM_FLAG = True Then: Randomize

CHOLESKY_MATRIX = CORREL_RNG
CHOLESKY_MATRIX = MATRIX_CHOLESKY_FUNC(CHOLESKY_MATRIX)
NSIZE = UBound(CHOLESKY_MATRIX, 2)
If UBound(CHOLESKY_MATRIX, 1) <> UBound(CHOLESKY_MATRIX, 2) Then: GoTo ERROR_LABEL

MEAN_VECTOR = MEAN_RNG
If UBound(MEAN_VECTOR, 2) = 1 Then
    MEAN_VECTOR = MATRIX_TRANSPOSE_FUNC(MEAN_VECTOR)
End If
If UBound(CHOLESKY_MATRIX, 2) <> UBound(MEAN_VECTOR, 2) Then: GoTo ERROR_LABEL

SIGMA_VECTOR = SIGMA_RNG
If UBound(SIGMA_VECTOR, 2) = 1 Then
    SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)
End If
If UBound(CHOLESKY_MATRIX, 2) <> UBound(SIGMA_VECTOR, 2) Then: GoTo ERROR_LABEL

RANDOM_MATRIX = MATRIX_RANDOM_NORMAL_FUNC(nLOOPS, NSIZE, NORM_TYPE, 0, 1, 0)

ReDim TEMP_MATRIX(1 To nLOOPS, 1 To NSIZE)
For i = 1 To nLOOPS
    For j = 1 To NSIZE
    TEMP_SUM = 0
          For k = 1 To j
                TEMP_SUM = TEMP_SUM + RANDOM_MATRIX(i, k) * CHOLESKY_MATRIX(j, k)
          Next k
    TEMP_MATRIX(i, j) = TEMP_SUM * SIGMA_VECTOR(1, j) + MEAN_VECTOR(1, j)
    Next j
Next i

MULTI_NORMAL_SIMULATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MULTI_NORMAL_SIMULATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTI_NORMAL_RANDOM_MATRIX_FUNC

'DESCRIPTION   : Multi Normal Random Matrix with antithetic variable technique
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_MULTI
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MULTI_NORMAL_RANDOM_MATRIX_FUNC(ByVal VERSION As Integer, _
ByVal NROWS As Long, _
ByVal NCOLUMNS As Long, _
Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal SIGMA_VAL As Double = 1, _
Optional ByVal RANDOM_FLAG As Boolean = True, _
Optional ByVal MOMENTS_FLAG As Boolean = True, _
Optional ByRef CORREL_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long

Dim NRV1_VAL As Double
Dim NRV2_VAL As Double

Dim MEAN_VECTOR As Variant
Dim SIGMA_VECTOR As Variant

Dim TEMP_MATRIX As Variant

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

NROWS = Fix(NROWS / 2#) * 2# 'AVOID ODD ENTRIES
PI_VAL = 3.14159265358979

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
ReDim MEAN_VECTOR(1 To 1, 1 To NCOLUMNS)
ReDim SIGMA_VECTOR(1 To 1, 1 To NCOLUMNS)

'-------------------------------------------------------------------------
'To repeat sequences of random numbers, call Rnd with a negative
'argument immediately before using Randomize with a numeric argument.
'Using Randomize with the same value for number does not repeat the
'previous sequence.
'-------------------------------------------------------------------------

If RANDOM_FLAG = False Then: Rnd (-1)

'-------------------------------------------------------------------------------------
Select Case VERSION
'-------------------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------------------
    For j = 1 To NCOLUMNS
        MEAN_VECTOR(1, j) = 0
        For i = 1 To NROWS Step 2
            NRV1_VAL = PSEUDO_RANDOM_FUNC(1)
            NRV2_VAL = PSEUDO_RANDOM_FUNC(1)
            TEMP_MATRIX(i, j) = NORMSINV_FUNC(NRV1_VAL, MEAN_VAL, SIGMA_VAL, 0)
            TEMP_MATRIX(i + 1, j) = NORMSINV_FUNC(NRV2_VAL, MEAN_VAL, SIGMA_VAL, 0)
            MEAN_VECTOR(1, j) = MEAN_VECTOR(1, j) + TEMP_MATRIX(i, j) / NROWS + TEMP_MATRIX(i + 1, j) / NROWS
        Next i
    Next j
'-------------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------------
'-------The following routine generates a Gaussian variable N(0,1): Transformation
'-------to get 2*n Normal random numbers from 2*n Uniform random numbers.
    For j = 1 To NCOLUMNS
        MEAN_VECTOR(1, j) = 0
        For i = 1 To NROWS Step 2
            NRV1_VAL = PSEUDO_RANDOM_FUNC(1)
            NRV2_VAL = PSEUDO_RANDOM_FUNC(1)
            TEMP_MATRIX(i, j) = (((-2 * Log(NRV1_VAL)) ^ 0.5) * Cos(2 * PI_VAL * NRV2_VAL)) * SIGMA_VAL + MEAN_VAL
            TEMP_MATRIX(i + 1, j) = (((-2 * Log(NRV1_VAL)) ^ 0.5) * Sin(2 * PI_VAL * NRV2_VAL)) * SIGMA_VAL + MEAN_VAL
            MEAN_VECTOR(1, j) = MEAN_VECTOR(1, j) + TEMP_MATRIX(i, j) / NROWS + TEMP_MATRIX(i + 1, j) / NROWS
        Next i
    Next j
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
If MOMENTS_FLAG = True Then 'Performing Moment Matching Calculation...
'-------------------------------------------------------------------------------------
    
'MOMENTS_FLAG MATCHING: Sometimes termed quadratic resampling, this method
'involves adjusting the samples taken from a standardized normal
'distribution so that the first (mean) and second (std dev)
'moments are matched.
    
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            SIGMA_VECTOR(1, j) = SIGMA_VECTOR(1, j) + (TEMP_MATRIX(i, j) - MEAN_VECTOR(1, j)) ^ 2
        Next i
        SIGMA_VECTOR(1, j) = SIGMA_VECTOR(1, j) / NROWS
        For i = 1 To NROWS 'Store Adjusted Random Normal in Array
            TEMP_MATRIX(i, j) = (TEMP_MATRIX(i, j) - MEAN_VECTOR(1, j)) / SIGMA_VECTOR(1, j) ^ 0.5
        Next i
    Next j
'-------------------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------------------

'------------All random Normals are adjusted for correlation
'------------when the random numbers are generated.

'-------------------------------------------------------------------------------------
If IsArray(CORREL_RNG) = True Then
    TEMP_MATRIX = MULTI_NORMAL_ADJUST_CORREL_FUNC(CORREL_RNG, TEMP_MATRIX)
End If
'-------------------------------------------------------------------------------------

Select Case OUTPUT
Case 0
    MULTI_NORMAL_RANDOM_MATRIX_FUNC = TEMP_MATRIX
Case 1
    MULTI_NORMAL_RANDOM_MATRIX_FUNC = MEAN_VECTOR
Case 2
    MULTI_NORMAL_RANDOM_MATRIX_FUNC = SIGMA_VECTOR
Case Else
    MULTI_NORMAL_RANDOM_MATRIX_FUNC = Array(TEMP_MATRIX, MEAN_VECTOR, SIGMA_VECTOR)
End Select

Exit Function
ERROR_LABEL:
MULTI_NORMAL_RANDOM_MATRIX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTI_NORMAL_MOMENTS_FUNC
'DESCRIPTION   : Multi Normal Simulation Moments
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_MULTI
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MULTI_NORMAL_MOMENTS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim nLOOPS As Long

Dim TEMP_SUM As Variant

Dim MEAN_VECTOR As Variant
Dim SIGMA_VECTOR As Variant

Dim CORRELATION_MATRIX As Variant
Dim COVARIANCE_MATRIX As Variant
Dim TEMP_MATRIX As Variant

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

nLOOPS = UBound(DATA_MATRIX, 1)
NSIZE = UBound(DATA_MATRIX, 2)

ReDim MEAN_VECTOR(1 To NSIZE, 1 To 1)
ReDim SIGMA_VECTOR(1 To NSIZE, 1 To 1)
ReDim COVARIANCE_MATRIX(1 To NSIZE, 1 To NSIZE)
ReDim CORRELATION_MATRIX(1 To NSIZE, 1 To NSIZE)

ReDim TEMP_MATRIX(1 To nLOOPS, 1 To NSIZE)

For j = 1 To NSIZE
    TEMP_SUM = 0
    For i = 1 To nLOOPS
        TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j)
    Next i
    MEAN_VECTOR(j, 1) = TEMP_SUM / nLOOPS  'mean
Next j

For j = 1 To NSIZE
    For i = 1 To nLOOPS
        TEMP_MATRIX(i, j) = DATA_MATRIX(i, j) - MEAN_VECTOR(j, 1)
    Next i
Next j

For i = 1 To NSIZE
    For j = 1 To NSIZE
        TEMP_SUM = 0
        For k = 1 To nLOOPS
            TEMP_SUM = TEMP_SUM + TEMP_MATRIX(k, i) * TEMP_MATRIX(k, j)
        Next k
        COVARIANCE_MATRIX(i, j) = TEMP_SUM / (nLOOPS - 1)
        'covariance
    Next j
    SIGMA_VECTOR(i, 1) = Sqr(COVARIANCE_MATRIX(i, i))     'standard deviation
Next i

For i = 1 To NSIZE
    For j = 1 To NSIZE
        CORRELATION_MATRIX(i, j) = COVARIANCE_MATRIX(i, j) / (SIGMA_VECTOR(i, 1) * SIGMA_VECTOR(j, 1))  'correlation
    Next j
Next i
    
Select Case VERSION
Case 0
    MULTI_NORMAL_MOMENTS_FUNC = MEAN_VECTOR
Case 1
    MULTI_NORMAL_MOMENTS_FUNC = SIGMA_VECTOR
Case 2
    MULTI_NORMAL_MOMENTS_FUNC = COVARIANCE_MATRIX
Case 3
    MULTI_NORMAL_MOMENTS_FUNC = CORRELATION_MATRIX
Case 4
    MULTI_NORMAL_MOMENTS_FUNC = TEMP_MATRIX
Case Else
    MULTI_NORMAL_MOMENTS_FUNC = Array(MEAN_VECTOR, SIGMA_VECTOR, COVARIANCE_MATRIX, CORRELATION_MATRIX, TEMP_MATRIX)
End Select

Exit Function
ERROR_LABEL:
MULTI_NORMAL_MOMENTS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTI_NORMAL_ADJUST_CORREL_FUNC
'DESCRIPTION   : Adjusting Random Normals for Correlations...
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_MULTI
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function MULTI_NORMAL_ADJUST_CORREL_FUNC(ByRef CORREL_RNG As Variant, _
ByRef RANDOM_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim nLOOPS As Long

Dim TEMP_SUM As Double

Dim CORREL_MATRIX As Variant
Dim RANDOM_MATRIX As Variant
Dim CHOLESKY_MATRIX As Variant

On Error GoTo ERROR_LABEL

CORREL_MATRIX = CORREL_RNG
If UBound(CORREL_MATRIX, 1) <> UBound(CORREL_MATRIX, 2) Then: GoTo ERROR_LABEL

CHOLESKY_MATRIX = MATRIX_CHOLESKY_FUNC(CORREL_MATRIX)
RANDOM_MATRIX = RANDOM_RNG

NSIZE = UBound(CHOLESKY_MATRIX, 2)
nLOOPS = UBound(RANDOM_MATRIX, 1)

'-------------Note it goes from bottom to top since the last
'random normal is only used once and the second from the bottom
'only twice, etc.

For i = 1 To nLOOPS
    For j = NSIZE To 1 Step -1
        TEMP_SUM = 0
        For k = 1 To j
            TEMP_SUM = TEMP_SUM + (CHOLESKY_MATRIX(k, j) * RANDOM_MATRIX(i, k))
        Next k
        RANDOM_MATRIX(i, j) = TEMP_SUM
    Next j
Next i

MULTI_NORMAL_ADJUST_CORREL_FUNC = RANDOM_MATRIX

Exit Function
ERROR_LABEL:
MULTI_NORMAL_ADJUST_CORREL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTI_NORMAL_CORREL_VOLATILITY_FUNC
'DESCRIPTION   : Returns a diagonal matrix with the correlation outputs
'and STDEVs of the simulation
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_MULTI
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function MULTI_NORMAL_CORREL_VOLATILITY_FUNC(ByRef DATA_RNG As Variant)

Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NSIZE + 2, 1 To NSIZE)

For j = 1 To NSIZE
    For k = j + 1 To NSIZE
        TEMP_MATRIX(j, k) = CORRELATION_FUNC(MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, j, 1), MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, k, 1), 0, 0)
        TEMP_MATRIX(k, j) = TEMP_MATRIX(j, k)
    Next k
    TEMP_MATRIX(j, j) = 1
    TEMP_MATRIX(NSIZE + 1, j) = ""
    TEMP_MATRIX(NSIZE + 2, j) = MATRIX_STDEV_FUNC(MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, j, 1))(1, 1)
Next j

MULTI_NORMAL_CORREL_VOLATILITY_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MULTI_NORMAL_CORREL_VOLATILITY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTI_NORMAL_CORREL_ERROR_FUNC
'DESCRIPTION   : Correlation Matrix Error Strings
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_MULTI
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function MULTI_NORMAL_CORREL_ERROR_FUNC(ByVal OUTPUT As Integer)

On Error GoTo ERROR_LABEL
    
Select Case OUTPUT
Case 0
    MULTI_NORMAL_CORREL_ERROR_FUNC = "Absolute value of rho must be less than or equal to one."
Case 1
    MULTI_NORMAL_CORREL_ERROR_FUNC = "Sigma must be positive"
Case 2
    MULTI_NORMAL_CORREL_ERROR_FUNC = "Mean array must have the same number of rows as the specificed number of variables."
Case 3
    MULTI_NORMAL_CORREL_ERROR_FUNC = "Sigma array must have the same number of rows as the specificed number of variables."
Case 4
    MULTI_NORMAL_CORREL_ERROR_FUNC = "Correlation matrix must have the same number of rows as the specificed number of variables."
Case 5
    MULTI_NORMAL_CORREL_ERROR_FUNC = "Correlation matrix must have the same number of columns as the specificed number of variables."
Case 6
    MULTI_NORMAL_CORREL_ERROR_FUNC = "Sigma must be positive"
Case 7
    MULTI_NORMAL_CORREL_ERROR_FUNC = "Correlation matrix must have entries less than or equal to one in absolute value."
Case 8
    MULTI_NORMAL_CORREL_ERROR_FUNC = "Correlation matrix must have ones on the diagonal."
End Select

Exit Function
ERROR_LABEL:
MULTI_NORMAL_CORREL_ERROR_FUNC = Err.number
End Function

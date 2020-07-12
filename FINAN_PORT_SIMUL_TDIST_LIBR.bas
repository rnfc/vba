Attribute VB_Name = "FINAN_PORT_SIMUL_TDIST_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_GAUSS_TDIST_SIMULATION_FUNC

'DESCRIPTION   : Multivariate Gaussian & Student T Simulation: Simulation of
'correlated returns with gaussian or student t (with same df) copulas and gaussian
'or (scaled) student t (with differing df) marginals. This simulation
'allows capturing the observed 'fat/long tails' as well as 'tail dependency'

'LIBRARY       : PORTFOLIO
'GROUP         : SIMULATION_NORMAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_GAUSS_TDIST_SIMULATION_FUNC(ByRef EXPECTED_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
ByRef CORREL_RNG As Variant, _
ByRef MARGINAL_RNG As Variant, _
Optional ByRef WEIGHTS_RNG As Variant, _
Optional ByVal nDEGREES As Integer = 4, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal NBINS As Long = 50, _
Optional ByVal MARG_VERSION As Integer = 0, _
Optional ByVal COPULA_VERSION As Integer = 0, _
Optional ByVal RANDOM_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'MARG_VERSION = 0; GAUSSIAN
'MARG_VERSION = 1; T-DIST

'COPULA_VERSION = 0; GAUSSIAN
'COPULA_VERSION = 1; T-DIST

'DEGREES --> DF COPULA

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim TEMP_VAL As Variant
Dim TEMP_SUM As Variant

Dim TEMP_VECTOR As Variant

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

Dim CORREL_MATRIX As Variant
Dim EXPECTED_VECTOR As Variant
Dim SIGMA_VECTOR As Variant
Dim MARGINAL_VECTOR As Variant

On Error GoTo ERROR_LABEL

EXPECTED_VECTOR = EXPECTED_RNG
If UBound(EXPECTED_VECTOR, 1) = 1 Then: EXPECTED_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPECTED_VECTOR)

CORREL_MATRIX = CORREL_RNG
If UBound(CORREL_MATRIX, 1) <> UBound(CORREL_MATRIX, 2) Then: GoTo ERROR_LABEL
If UBound(EXPECTED_VECTOR, 1) <> UBound(CORREL_MATRIX, 2) Then: GoTo ERROR_LABEL

SIGMA_VECTOR = SIGMA_RNG
If UBound(SIGMA_VECTOR, 1) = 1 Then: SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)
If UBound(SIGMA_VECTOR, 1) <> UBound(CORREL_MATRIX, 2) Then: GoTo ERROR_LABEL

MARGINAL_VECTOR = MARGINAL_RNG
If UBound(MARGINAL_VECTOR, 1) = 1 Then: MARGINAL_VECTOR = MATRIX_TRANSPOSE_FUNC(MARGINAL_VECTOR)
If UBound(MARGINAL_VECTOR, 1) <> UBound(CORREL_MATRIX, 2) Then: GoTo ERROR_LABEL

NSIZE = UBound(EXPECTED_VECTOR, 1)


ReDim TEMP_VECTOR(1 To 1, 1 To nLOOPS)
ReDim TEMP1_MATRIX(1 To NSIZE, 1 To nLOOPS)


For i = 1 To NSIZE
    For j = 1 To nLOOPS
        TEMP1_MATRIX(i, j) = NORMSINV_FUNC(PSEUDO_RANDOM_FUNC(RANDOM_TYPE), 0, 1, 0)
        TEMP_VECTOR(1, j) = INVERSE_CHI_SQUARED_DIST_FUNC(PSEUDO_RANDOM_FUNC(RANDOM_TYPE), nDEGREES, False)
    Next j
Next i

'-------------------------------------------------------------------------------------------
'--------------------------------CHOLESKY Decomposition-------------------------------------
'-------------------------------------------------------------------------------------------

ReDim TEMP2_MATRIX(1 To NSIZE, 1 To NSIZE)
For j = 1 To NSIZE
    TEMP_SUM = 0
    For k = 1 To j - 1
        TEMP_SUM = TEMP_SUM + TEMP2_MATRIX(j, k) ^ 2
    Next k
    TEMP2_MATRIX(j, j) = CORREL_MATRIX(j, j) - TEMP_SUM
    ' Matrix is not semi-positive definite, no solution exists
    If TEMP2_MATRIX(j, j) < 0 Then: GoTo ERROR_LABEL
    TEMP2_MATRIX(j, j) = Sqr(MAXIMUM_FUNC(0, TEMP2_MATRIX(j, j)))
    
    For i = j + 1 To NSIZE
        TEMP_SUM = 0
        For k = 1 To j - 1
            TEMP_SUM = TEMP_SUM + TEMP2_MATRIX(i, k) * TEMP2_MATRIX(j, k)
        Next k
        If TEMP2_MATRIX(j, j) = 0 Then
           TEMP2_MATRIX(i, j) = 0
        Else
           TEMP2_MATRIX(i, j) = (CORREL_MATRIX(i, j) - TEMP_SUM) / TEMP2_MATRIX(j, j)
        End If
    Next i
Next j
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------

TEMP1_MATRIX = MMULT_FUNC(TEMP2_MATRIX, TEMP1_MATRIX, 70)

ReDim TEMP2_MATRIX(1 To NSIZE, 1 To nLOOPS)
For i = 1 To NSIZE
    For j = 1 To nLOOPS
        If MARG_VERSION = 0 And COPULA_VERSION = 0 Then
            TEMP2_MATRIX(i, j) = EXPECTED_VECTOR(i, 1) + _
                                 SIGMA_VECTOR(i, 1) * TEMP1_MATRIX(i, j)
        End If
        If MARG_VERSION = 0 And COPULA_VERSION <> 0 Then
            TEMP_VAL = TEMP1_MATRIX(i, j) * Sqr(nDEGREES / TEMP_VECTOR(1, j))
            If TEMP_VAL > 0 Then
                TEMP_VAL = TDIST_FUNC(TEMP_VAL, nDEGREES, True)
            Else
                TEMP_VAL = 1 - TDIST_FUNC(-TEMP_VAL, nDEGREES, True)
            End If
            
            TEMP2_MATRIX(i, j) = EXPECTED_VECTOR(i, 1) + SIGMA_VECTOR(i, 1) * NORMSINV_FUNC(TEMP_VAL, 0, 1, 0)
        End If
        If MARG_VERSION <> 0 And COPULA_VERSION = 0 Then
            TEMP_VAL = NORMSDIST_FUNC(TEMP1_MATRIX(i, j), 0, 1, 0)
            If TEMP_VAL < 0.5 Then
                TEMP_VAL = INVERSE_TDIST_FUNC(TEMP_VAL, MARGINAL_VECTOR(i, 1))
            Else
                TEMP_VAL = -1 * INVERSE_TDIST_FUNC((1 - TEMP_VAL), MARGINAL_VECTOR(i, 1))
            End If
            ' Scaled T Marginals with scale factor volatility
            TEMP2_MATRIX(i, j) = EXPECTED_VECTOR(i, 1) + SIGMA_VECTOR(i, 1) * TEMP_VAL
        End If
        If MARG_VERSION <> 0 And COPULA_VERSION <> 0 Then
            TEMP_VAL = TEMP1_MATRIX(i, j) * Sqr(nDEGREES / TEMP_VECTOR(1, j))
            If TEMP_VAL > 0 Then
                TEMP_VAL = TDIST_FUNC(TEMP_VAL, nDEGREES, True)
            Else
                TEMP_VAL = 1 - TDIST_FUNC(-TEMP_VAL, nDEGREES, True)
            End If
            
            If TEMP_VAL < 0.5 Then
                TEMP_VAL = INVERSE_TDIST_FUNC(TEMP_VAL, MARGINAL_VECTOR(i, 1))
            Else
                TEMP_VAL = -1 * INVERSE_TDIST_FUNC((1 - TEMP_VAL), MARGINAL_VECTOR(i, 1))
            End If
            
            ' Scaled T Marginals with scale factor volatility
            TEMP2_MATRIX(i, j) = EXPECTED_VECTOR(i, 1) + SIGMA_VECTOR(i, 1) * TEMP_VAL
        End If
    Next j
Next i


Select Case OUTPUT
Case 0
    PORT_GAUSS_TDIST_SIMULATION_FUNC = MATRIX_TRANSPOSE_FUNC(TEMP2_MATRIX)
Case Else
    If IsArray(WEIGHTS_RNG) = True Then
        TEMP2_MATRIX = MATRIX_TRANSPOSE_FUNC(TEMP2_MATRIX)
        NSIZE = UBound(TEMP2_MATRIX, 1)
        ReDim TEMP1_MATRIX(1 To NSIZE, 1 To 1)
        For i = 1 To NSIZE
            TEMP1_MATRIX(i, 1) = PORT_WEIGHTED_RETURN2_FUNC(MATRIX_GET_ROW_FUNC(TEMP2_MATRIX, i, 1), WEIGHTS_RNG)
        Next i
        PORT_GAUSS_TDIST_SIMULATION_FUNC = HISTOGRAM_DYNAMIC_FREQUENCY_FUNC(TEMP1_MATRIX, NBINS, 0, 0)
    Else
        GoTo ERROR_LABEL
    End If
End Select

Exit Function
ERROR_LABEL:
PORT_GAUSS_TDIST_SIMULATION_FUNC = Err.number
End Function


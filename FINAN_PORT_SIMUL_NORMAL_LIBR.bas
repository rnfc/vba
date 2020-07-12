Attribute VB_Name = "FINAN_PORT_SIMUL_NORMAL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RETURNS_NORMAL_SIMULATION_FUNC

'DESCRIPTION   : Simulating Multivariate Normal Distributed Returns:
'Function to generate normally distributed and correlated returns.
'Correlation Matrix (needs to be semi-positive definite!)

'LIBRARY       : PORTFOLIO
'GROUP         : SIMULATION_NORMAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_RETURNS_NORMAL_SIMULATION_FUNC(ByRef EXPECTED_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
ByRef CORREL_RNG As Variant, _
Optional ByVal nLOOPS As Variant = 500, _
Optional ByVal RANDOM_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim TEMP_SUM As Double

Dim CORREL_MATRIX As Variant
Dim SIGMA_VECTOR As Variant
Dim EXPECTED_VECTOR As Variant

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

EXPECTED_VECTOR = EXPECTED_RNG
If UBound(EXPECTED_VECTOR, 1) = 1 Then: _
    EXPECTED_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPECTED_VECTOR)

SIGMA_VECTOR = SIGMA_RNG
If UBound(SIGMA_VECTOR, 1) = 1 Then: SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)

CORREL_MATRIX = CORREL_RNG
If UBound(CORREL_MATRIX, 1) <> UBound(CORREL_MATRIX, 2) Then: GoTo ERROR_LABEL

If UBound(CORREL_MATRIX, 1) <> UBound(SIGMA_VECTOR, 1) Then: GoTo ERROR_LABEL
If UBound(CORREL_MATRIX, 1) <> UBound(EXPECTED_VECTOR, 1) Then: GoTo ERROR_LABEL

NSIZE = UBound(EXPECTED_VECTOR, 1)

ReDim ATEMP_MATRIX(1 To NSIZE, 1 To nLOOPS)
For i = 1 To NSIZE
    For j = 1 To nLOOPS
        ATEMP_MATRIX(i, j) = NORMSINV_FUNC(PSEUDO_RANDOM_FUNC(RANDOM_TYPE), 0, 1, 0)
    Next j
Next i

'-------------------------------------------------------------------------------
'------------------------------CHOLESKY Decomposition---------------------------
'-------------------------------------------------------------------------------

ReDim BTEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

For j = 1 To NSIZE
    TEMP_SUM = 0
    For k = 1 To j - 1
        TEMP_SUM = TEMP_SUM + BTEMP_MATRIX(j, k) ^ 2
    Next k
    BTEMP_MATRIX(j, j) = CORREL_MATRIX(j, j) - TEMP_SUM
    ' Matrix is not semi-positive definite, no solution exists
    If BTEMP_MATRIX(j, j) < 0 Then: GoTo ERROR_LABEL
    BTEMP_MATRIX(j, j) = Sqr(MAXIMUM_FUNC(0, BTEMP_MATRIX(j, j)))
    For i = j + 1 To NSIZE
        TEMP_SUM = 0
        For k = 1 To j - 1
            TEMP_SUM = TEMP_SUM + BTEMP_MATRIX(i, k) * BTEMP_MATRIX(j, k)
        Next k
        If BTEMP_MATRIX(j, j) = 0 Then
           BTEMP_MATRIX(i, j) = 0
        Else
           BTEMP_MATRIX(i, j) = (CORREL_MATRIX(i, j) - TEMP_SUM) / BTEMP_MATRIX(j, j)
        End If
    Next i
Next j

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

ATEMP_MATRIX = MMULT_FUNC(BTEMP_MATRIX, ATEMP_MATRIX, 70)

For i = 1 To NSIZE
    For j = 1 To nLOOPS
        ATEMP_MATRIX(i, j) = ATEMP_MATRIX(i, j) * _
                        SIGMA_VECTOR(i, 1) + _
                        EXPECTED_VECTOR(i, 1)
    Next j
Next i

'-------------------------------------------------------------------------------
ATEMP_MATRIX = MATRIX_TRANSPOSE_FUNC(ATEMP_MATRIX)
'-------------------------------------------------------------------------------

Select Case OUTPUT
Case 0
    PORT_RETURNS_NORMAL_SIMULATION_FUNC = MATRIX_TRANSPOSE_FUNC(MATRIX_MEAN_FUNC(ATEMP_MATRIX))
Case 1
    PORT_RETURNS_NORMAL_SIMULATION_FUNC = _
        PORT_SIGMA_COVAR_FUNC(MATRIX_COVARIANCE_FRAME1_FUNC(ATEMP_MATRIX, 0, 0))
Case 2 'Covariance matrix from sigma and correlation
    PORT_RETURNS_NORMAL_SIMULATION_FUNC = _
        PORT_COVAR_SIGMA_FUNC(PORT_SIGMA_COVAR_FUNC(MATRIX_COVARIANCE_FRAME1_FUNC(ATEMP_MATRIX, 0, 0)), _
        PORT_CORREL_COVAR_FUNC(MATRIX_COVARIANCE_FRAME1_FUNC(ATEMP_MATRIX, 0, 0)))
Case 3 'Correlation coefficient matrix from covariance matrix
    PORT_RETURNS_NORMAL_SIMULATION_FUNC = _
        PORT_CORREL_COVAR_FUNC(MATRIX_COVARIANCE_FRAME1_FUNC(ATEMP_MATRIX, 0, 0))
Case Else 'Resampled Returns
    PORT_RETURNS_NORMAL_SIMULATION_FUNC = ATEMP_MATRIX
End Select
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_RETURNS_NORMAL_SIMULATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_WEIGHTS_NORMAL_SIMULATION_FUNC
'DESCRIPTION   : Runs a Multivariate Portfolio Normal Simulation with GIVEN Weights
'LIBRARY       : PORTFOLIO
'GROUP         : SIMULATION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************
'(PORT_NORMAL_SIMULATION_FUNC)

Function PORT_WEIGHTS_NORMAL_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByRef CORREL_RNG As Variant, _
ByRef WEIGHTS_RNG As Variant, _
ByRef MEAN_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
ByVal INITIAL_INVEST As Double, _
ByVal CASH_RATE As Double, _
ByVal TENOR As Double, _
Optional ByVal COUNT_BASIS As Double = 252, _
Optional ByVal NORM_TYPE As Integer = 0, _
Optional ByVal RANDOM_FLAG As Boolean = True)

'CASH_RATE = Cash Rate; Borrowing Rate --> Annualized

Dim i As Long
Dim j As Long
Dim k As Long
Dim NSIZE As Long

Dim PERIODS As Double
Dim DELTA_TIME As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim PORT_VAL As Double
Dim CASH_WEIGHT_VAL As Double

Dim MEAN_VECTOR As Variant
Dim CORREL_MATRIX As Variant
Dim SIGMA_VECTOR As Variant
Dim WEIGHTS_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL

CORREL_MATRIX = CORREL_RNG
If UBound(CORREL_MATRIX, 1) <> UBound(CORREL_MATRIX, 2) Then: GoTo ERROR_LABEL

NSIZE = UBound(CORREL_MATRIX, 2)

SIGMA_VECTOR = SIGMA_RNG
If UBound(SIGMA_VECTOR, 2) = 1 Then
    SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)
End If
MEAN_VECTOR = MEAN_RNG
If UBound(MEAN_VECTOR, 2) = 1 Then
    MEAN_VECTOR = MATRIX_TRANSPOSE_FUNC(MEAN_VECTOR)
End If
WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 2) = 1 Then
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
End If
    
CASH_WEIGHT_VAL = 1 - MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC(WEIGHTS_VECTOR)
PERIODS = (TENOR * COUNT_BASIS)
DELTA_TIME = TENOR / PERIODS '-----> SAME AS: 1 / COUNT_BASIS

ReDim TEMP_MATRIX(1 To nLOOPS, 1 To 3)

For i = 1 To nLOOPS
    PORT_VAL = INITIAL_INVEST
    
    RANDOM_MATRIX = MULTI_NORMAL_SIMULATION_FUNC(CORREL_MATRIX, MEAN_VECTOR, SIGMA_VECTOR, _
    PERIODS, NORM_TYPE, RANDOM_FLAG)
    
    BTEMP_SUM = 0
    For k = 1 To PERIODS
        ATEMP_SUM = 0
        
        For j = 1 To NSIZE
            ATEMP_SUM = ATEMP_SUM + WEIGHTS_VECTOR(1, j) * (RANDOM_MATRIX(k, j))
        Next j
        
        ATEMP_SUM = ATEMP_SUM + (CASH_WEIGHT_VAL * CASH_RATE * DELTA_TIME)
        PORT_VAL = PORT_VAL * Exp(ATEMP_SUM) 'PORTFOLIO VALUE PER _
        PERIOD [using log-normal returns]
        BTEMP_SUM = BTEMP_SUM + ATEMP_SUM
    Next k
    
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = PORT_VAL
    TEMP_MATRIX(i, 3) = (BTEMP_SUM / PERIODS) * (1 / DELTA_TIME) _
    '--> Annualized Returns
Next i
    
PORT_WEIGHTS_NORMAL_SIMULATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_WEIGHTS_NORMAL_SIMULATION_FUNC = Err.number
End Function

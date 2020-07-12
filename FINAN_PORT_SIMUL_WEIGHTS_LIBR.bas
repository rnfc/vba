Attribute VB_Name = "FINAN_PORT_SIMUL_WEIGHTS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_WEIGHTS_UNIFORM_SIMULATION_FUNC
'DESCRIPTION   : Run an Optimal Portfolio Simulation without Short Sales
'LIBRARY       : PORTFOLIO
'GROUP         : SIMULATION_WEIGHTS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************
'PORT_WITHOUT_SHORT_OPTIMIZER_FUNC

Function PORT_WEIGHTS_UNIFORM_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByRef COVAR_RNG As Variant, _
ByRef MEAN_RNG As Variant, _
ByVal CASH_RATE As Double, _
Optional ByVal COUNT_BASIS As Double = 52, _
Optional ByVal RANDOM_FLAG As Boolean = True)

'CASH_RATE = Cash Rate; Borrowing Rate Annualized
'MEAN_RNG = Mean Vector Annualized

Dim i As Long
Dim j As Long
Dim k As Long
Dim NSIZE As Long

Dim VAR_SUM As Double
Dim MEAN_SUM As Double
Dim TEMP_SUM As Double

Dim MEAN_VECTOR As Variant
Dim WEIGHTS_VECTOR As Variant

Dim COVAR_MATRIX As Variant
Dim RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL

COVAR_MATRIX = COVAR_RNG
If UBound(COVAR_MATRIX, 2) <> UBound(COVAR_MATRIX, 1) Then: GoTo ERROR_LABEL
NSIZE = UBound(COVAR_MATRIX, 2)

MEAN_VECTOR = MEAN_RNG
If UBound(MEAN_VECTOR, 1) = 1 Then: MEAN_VECTOR = MATRIX_TRANSPOSE_FUNC(MEAN_VECTOR)

ReDim WEIGHTS_VECTOR(1 To 1, 1 To NSIZE)
ReDim TEMP_MATRIX(1 To nLOOPS, 1 To 4 + NSIZE)

If RANDOM_FLAG = True Then: Randomize

For k = 1 To nLOOPS 'start iteration
    
    VAR_SUM = 0
    TEMP_SUM = 0
    MEAN_SUM = 0
    
    RANDOM_MATRIX = MATRIX_RANDOM_UNIFORM_FUNC(NSIZE, 1, 0, 0)
    TEMP_SUM = MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC(RANDOM_MATRIX)
    
    For i = 1 To NSIZE
        WEIGHTS_VECTOR(1, i) = RANDOM_MATRIX(i, 1) / TEMP_SUM
        MEAN_SUM = MEAN_SUM + WEIGHTS_VECTOR(1, i) * MEAN_VECTOR(i, 1)
    Next i
    
    TEMP_MATRIX(k, 1) = k 'COUNTER
    TEMP_MATRIX(k, 2) = MEAN_SUM 'PORT MEAN
        
    For i = 1 To NSIZE  'compute the diagonal sum = Wi x Wi x Vari
        VAR_SUM = VAR_SUM + WEIGHTS_VECTOR(1, i) ^ 2 * COVAR_MATRIX(i, i)
    Next i

    For i = 1 To NSIZE  'compute the other sum = 2 x Wi x Wj x Varij
        For j = i + 1 To NSIZE
            VAR_SUM = VAR_SUM + 2 * WEIGHTS_VECTOR(1, i) * WEIGHTS_VECTOR(1, j) * COVAR_MATRIX(i, j)
        Next j
    Next i

    TEMP_MATRIX(k, 3) = (VAR_SUM) ^ 0.5 * Sqr(COUNT_BASIS) _
        'STANDARD DEVIATION ANNUALIZED
    TEMP_MATRIX(k, 4) = (TEMP_MATRIX(k, 2) - CASH_RATE) / TEMP_MATRIX(k, 3) _
        'SHARPE RATIO
    For i = 1 To NSIZE
        TEMP_MATRIX(k, 4 + i) = WEIGHTS_VECTOR(1, i)
    Next i
    
Next k

TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 4, 0)
'SORT ARRAY BASED ON SHARPE RATIO

'AFTER COLUMN 4 THE RESULTS ARE THE WEIGHTS FOR EACH ASSET

PORT_WEIGHTS_UNIFORM_SIMULATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_WEIGHTS_UNIFORM_SIMULATION_FUNC = Err.number
End Function

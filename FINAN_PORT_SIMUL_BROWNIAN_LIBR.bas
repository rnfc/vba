Attribute VB_Name = "FINAN_PORT_SIMUL_BROWNIAN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_GEOMETRIC_BROWNIAN_SIMULATION_FUNC
'DESCRIPTION   : Multiple Asset Simulation: Geometric Brownian Motion
'LIBRARY       : PORTFOLIO
'GROUP         : SIMULATION_BROWNIAN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_GEOMETRIC_BROWNIAN_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByRef DATA_RNG As Variant, _
ByVal TENOR As Double, _
Optional ByVal COUNT_BASIS As Double = 52, _
Optional ByVal DATA_TYPE As Integer = 1, _
Optional ByVal LOG_SCALE As Integer = 1, _
Optional ByVal NORM_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal RANDOM_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim PERIODS As Double
Dim DELTA_TIME As Double

Dim MEAN_VECTOR As Variant
Dim SIGMA_VECTOR As Variant
Dim SPOT_VECTOR As Variant

Dim DATA_MATRIX As Variant
Dim CORREL_MATRIX As Variant

Dim TEMP_MATRIX As Variant
Dim SUMMARY_MATRIX As Variant
Dim NORMAL_RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
SPOT_VECTOR = MATRIX_GET_ROW_FUNC(DATA_MATRIX, UBound(DATA_MATRIX, 1), 1)

If DATA_TYPE <> 0 Then: DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

'-----------------ANNUALIZED MEAN & SIGMA VECTOR--------------------

MEAN_VECTOR = MATRIX_MEAN_FUNC(DATA_MATRIX)
MEAN_VECTOR = VECTOR_ELEMENTS_MULT_SCALAR_FUNC(MEAN_VECTOR, COUNT_BASIS, 0)
SIGMA_VECTOR = MATRIX_STDEVP_FUNC(DATA_MATRIX)
SIGMA_VECTOR = VECTOR_ELEMENTS_MULT_SCALAR_FUNC(SIGMA_VECTOR, (COUNT_BASIS) ^ 0.5, 0)
CORREL_MATRIX = MATRIX_CORRELATION_PEARSON_FUNC(DATA_MATRIX, 0, 0)
PERIODS = COUNT_BASIS * TENOR
DELTA_TIME = TENOR / PERIODS '--> SAME AS 1 / COUNT_BASIS

'----------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To PERIODS, 1 To NCOLUMNS)
ReDim SUMMARY_MATRIX(1 To nLOOPS, 1 To NCOLUMNS)
'----------------------------------------------------------------------------------
For i = 1 To nLOOPS
'----------------------------------------------------------------------------------
    NORMAL_RANDOM_MATRIX = MULTI_NORMAL_CORREL_SIMUL_FUNC(CORREL_MATRIX, PERIODS, NORM_TYPE, RANDOM_FLAG)
'----------------------------------------------------------------------------------
    For j = 1 To NCOLUMNS
'----------------------------------------------------------------------------------
        TEMP_MATRIX(0, j) = SPOT_VECTOR(1, j)
'----------------------------------------------------------------------------------
        For k = 1 To PERIODS 'Geometric Brownian Motion Function
            TEMP_MATRIX(k, j) = TEMP_MATRIX(k - 1, j) * Exp((MEAN_VECTOR(1, j) - 0.5 * SIGMA_VECTOR(1, j) ^ 2) * DELTA_TIME + SIGMA_VECTOR(1, j) * Sqr(DELTA_TIME) * NORMAL_RANDOM_MATRIX(k, j))
            'Source: Hull, John C., Options, Futures & Other Derivatives.
            'Fourth edition (2000). Prentice-Hall. P. 220
        Next k
'----------------------------------------------------------------------------------
        SUMMARY_MATRIX(i, j) = TEMP_MATRIX(PERIODS, j)
'----------------------------------------------------------------------------------
    Next j
'----------------------------------------------------------------------------------
Next i
'----------------------------------------------------------------------------------
    
Select Case OUTPUT
Case 0
    PORT_GEOMETRIC_BROWNIAN_SIMULATION_FUNC = SUMMARY_MATRIX
Case Else
    PORT_GEOMETRIC_BROWNIAN_SIMULATION_FUNC = TEMP_MATRIX
End Select
    
Exit Function
ERROR_LABEL:
PORT_GEOMETRIC_BROWNIAN_SIMULATION_FUNC = Err.number
End Function


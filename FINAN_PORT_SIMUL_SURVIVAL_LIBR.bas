Attribute VB_Name = "FINAN_PORT_SIMUL_SURVIVAL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_SURVIVAL_RATE_SIMULATION_FUNC
'DESCRIPTION   : PORTFOLIO SURVIVAL RATE SIMULATION
'LIBRARY       : PORTFOLIO
'GROUP         : SIMULATION_SURVIVAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_SURVIVAL_RATE_SIMULATION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef WEIGHTS_RNG As Variant, _
Optional ByRef INFLAT_RATE As Double = 0.03, _
Optional ByRef INIT_WITHDR_RATE As Double = 0.08, _
Optional ByVal REBALANCE_PERIODS As Long = 4, _
Optional ByVal COUNT_BASIS As Long = 52, _
Optional ByVal TENOR As Long = 40, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal DATA_TYPE As Integer = 0)

'INFLAT_RATE: Annual Inflation Rate
'INIT_WITHDR_RATE: Annual Withdrawal Rate

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim COUNTER As Long

Dim NROWS As Long
Dim NCOLUMNS As Long 'NO of ASSETS IN THE PORTFOLIO

Dim PORT_VAL As Double
Dim WITHDR_VAL As Double

Dim DATA_ARR As Variant

Dim DATA_MATRIX As Variant
Dim WEIGHTS_VECTOR As Variant

On Error GoTo ERROR_LABEL

'REBALANCE_PERIODS: withdraw every x PERIODS

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, 0)
NROWS = UBound(DATA_MATRIX, 1) ' number of points
NCOLUMNS = UBound(DATA_MATRIX, 2) 'number of assets

If IsArray(WEIGHTS_RNG) = True Then
    WEIGHTS_VECTOR = WEIGHTS_RNG
    If UBound(WEIGHTS_VECTOR, 1) = 1 Then: WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
    If UBound(WEIGHTS_VECTOR, 1) <> NCOLUMNS Then: GoTo ERROR_LABEL
Else
    ReDim WEIGHTS_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        WEIGHTS_VECTOR(i, 1) = 1 / NCOLUMNS 'Equal Weight
    Next i
End If

INFLAT_RATE = INFLAT_RATE / (COUNT_BASIS / REBALANCE_PERIODS) 'Per Period
INFLAT_RATE = 1 + INFLAT_RATE
INIT_WITHDR_RATE = INIT_WITHDR_RATE / (COUNT_BASIS / REBALANCE_PERIODS) 'Per Period

COUNTER = 0               ' COUNTER failures

ReDim DATA_ARR(1 To NCOLUMNS)

For i = 1 To nLOOPS       ' number of iterations
    Randomize           ' set random seed
    PORT_VAL = 1               ' set initial Portfolio to $1 for each iteration
    WITHDR_VAL = INIT_WITHDR_RATE   ' set initial withdrawal rate for each iteration

    For j = 1 To TENOR    ' go thru TENOR years
        For h = 1 To NCOLUMNS 'set / rebalance asset values
            DATA_ARR(h) = WEIGHTS_VECTOR(h, 1) * PORT_VAL
        Next h
        For k = 1 To COUNT_BASIS                 ' number of PERIODS
            l = 1 + (NROWS - 1) * Rnd                ' pick a random row / gain
            PORT_VAL = 0
            For h = 1 To NCOLUMNS 'set / rebalance asset values
                DATA_ARR(h) = DATA_ARR(h) * (1 + DATA_MATRIX(l, h))
                ' change x- and y-values asset values
                PORT_VAL = PORT_VAL + DATA_ARR(h)
                ' change portfolio ... daily/weekly/monthly
            Next h
            If Int(k / REBALANCE_PERIODS) = k / REBALANCE_PERIODS Then
            ' withdraw every x PERIODS
                WITHDR_VAL = WITHDR_VAL * INFLAT_RATE
                ' increase withdrawal by x-period increase
                PORT_VAL = PORT_VAL - WITHDR_VAL ' subtract withdrawal
            End If
            If PORT_VAL < 0 Then
                PORT_VAL = 0                           ' stop if PORT_VAL < 0
                k = COUNT_BASIS
                j = TENOR
            End If
        Next k                    ' continue to end of the period/year
    Next j                        ' start another year

    If PORT_VAL = 0 Then COUNTER = COUNTER + 1   ' COUNTER dead portfolios
Next i  ' next iteration

PORT_SURVIVAL_RATE_SIMULATION_FUNC = 1 - COUNTER / nLOOPS

Exit Function
ERROR_LABEL:
PORT_SURVIVAL_RATE_SIMULATION_FUNC = Err.number
End Function


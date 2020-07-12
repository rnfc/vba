Attribute VB_Name = "FINAN_PORT_WEIGHTS_REBALAN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RETURNS_REBALANCE1_FUNC

'DESCRIPTION   : Algorithm to calculate portfolio return given initial weights, a
'rebalancing frequency, transcation costs and a drift tolerance.
'Output consists of a matrix of portfolio returns, eop weights,
'incurred transcation costs and a flag whether bop period rebalancing
'tool place or not.

'LIBRARY       : PORTFOLIO
'GROUP         : REBALANCE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_RETURNS_REBALANCE1_FUNC(ByRef DATA_RNG As Variant, _
ByRef WEIGHTS_RNG As Variant, _
Optional ByVal REBALANCE_FREQUENCY_VAL As Double = 12, _
Optional ByVal REBALANCE_TOLERANCE_VAL As Double = 0.05, _
Optional ByVal COSTS_BP_VAL As Double = 10, _
Optional ByVal FACTOR_VAL As Double = 10000, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

'IF REBALANCE_FREQUENCY_VAL = 0 And REBALANCE_TOLERANCE_VAL = 1E+200 Then:
'Portfolio Return --> Buy & Hold Portfolio Return

'REBALANCE_FREQUENCY_VAL: A rebalancing frequency of zero means "buy & hold"

'REBALANCE_TOLERANCE_VAL: If "Rebalancing Tolerance" is set, then the portfolio is
'rebalanced whenever the drift of any asset exceeds the
'tolerance, independent of the rebalancing frequency set.

'COSTS_BP_VAL: Transaction Cost Per $ Transacted [bp] ; also should includes
'initial setup costs.

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double

Dim DRIFTED_FLAG As Boolean
Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim WEIGHTS_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
NROWS = UBound(DATA_MATRIX, 1) ' number of observations
NCOLUMNS = UBound(DATA_MATRIX, 2) ' number of assets

DATA_VECTOR = WEIGHTS_RNG
If UBound(DATA_VECTOR, 2) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
'If UBound(DATA_VECTOR, 2) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
WEIGHTS_VECTOR = DATA_VECTOR
k = REBALANCE_FREQUENCY_VAL
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 3)
TEMP_MATRIX(0, 1) = ("PORT_RETURN")
For j = 1 To NCOLUMNS
    TEMP_MATRIX(0, j + 1) = ("ASSET: ") & j
Next j

'model buy & hold as rebalancing at > NObs
If k = 0 Then k = NROWS 'model rebalancing tolerance
TEMP_SUM = 0
DRIFTED_FLAG = True
For i = 1 To NROWS 'Calc beginning of period weights
    If DRIFTED_FLAG Or (i Mod k = 1) Or (k = 1) Then
        TEMP_MATRIX(i, NCOLUMNS + 2) = True
        WEIGHTS_VECTOR = DATA_VECTOR 'Calc transcation costs
        If i > 1 Then
            For j = 1 To NCOLUMNS
                TEMP_MATRIX(i, NCOLUMNS + 3) = TEMP_MATRIX(i, NCOLUMNS + 3) + COSTS_BP_VAL * Abs(WEIGHTS_VECTOR(1, j) - TEMP_MATRIX(i - 1, 1 + j)) / FACTOR_VAL
            Next j
        Else
            For j = 1 To NCOLUMNS
                TEMP_MATRIX(i, NCOLUMNS + 3) = TEMP_MATRIX(i, NCOLUMNS + 3) + COSTS_BP_VAL * Abs(WEIGHTS_VECTOR(1, j)) / FACTOR_VAL
            Next j
        End If
    Else
        TEMP_MATRIX(i, NCOLUMNS + 2) = False
        For j = 1 To NCOLUMNS
            WEIGHTS_VECTOR(1, j) = TEMP_MATRIX(i - 1, 1 + j)
        Next j
    End If
    For j = 1 To NCOLUMNS ' Calc portfolio return
        TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 1) + WEIGHTS_VECTOR(1, j) * DATA_MATRIX(i, j)
    Next j
    DRIFTED_FLAG = False
    For j = 1 To NCOLUMNS ' Calc end of period weights & check for drift
        TEMP_MATRIX(i, 1 + j) = WEIGHTS_VECTOR(1, j) * (1 + DATA_MATRIX(i, j)) / (1 + TEMP_MATRIX(i, 1))
        DRIFTED_FLAG = DRIFTED_FLAG Or (Abs(TEMP_MATRIX(i, 1 + j) - DATA_VECTOR(1, j)) > REBALANCE_TOLERANCE_VAL)
    Next j
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, NCOLUMNS + 3)
Next i

TEMP_MATRIX(0, NCOLUMNS + 2) = "WAS REBALANCE BOP"
TEMP_MATRIX(0, NCOLUMNS + 3) = "PORT TRANSACTION COSTS " & Format(TEMP_SUM, "0.0000%")

PORT_RETURNS_REBALANCE1_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_RETURNS_REBALANCE1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RETURNS_REBALANCE2_FUNC
'DESCRIPTION   :
'LIBRARY       : PORTFOLIO
'GROUP         : REBALANCE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_RETURNS_REBALANCE2_FUNC(ByRef TICKERS_RNG As Variant, _
ByRef DATES_RNG As Variant, _
ByRef DATA_RNG As Variant, _
Optional ByRef WEIGHTS_RNG As Variant, _
Optional ByVal LOWER_RNG As Variant = 0.15, _
Optional ByVal UPPER_RNG As Variant = 0.35, _
Optional ByVal INITIAL_CASH_VAL As Double = 10000, _
Optional ByVal FACTOR_VAL As Double = 100)

'DataSources:
'SPY, GLD, SHY, TLT
'Ibbotson SBBI 2008 Classic Yearbook: Market Results for Stocks, Bonds, Bills, and Inflation (Morningstar).
'Kitco Historical Gold Charts: http://www.kitco.com/charts/historicalgold.html
'Permanent Portfolio Fund (PRPFX) annual returns from the Pacific Heights Asset Management Company: http://www.permanentportfoliofunds.com/pdfs/2008%20PPF%20Annual%20Returns.pdf

'Total returns for S&P 500 Index from Ibbotson. Note: 2008 figure is the return for the Vanguard S&P 500 index fund (VFINX).
'Total returns for 30-day US treasury bills from Ibbotson. Note: 2008 figure is the return for Vanguard's treasury money market fund (VMPXX).
'Total returns for 20-year US treasury bonds from Ibbotson. Notes: 2008 figure is the return for iShares Barclay's 20+ year Treasury bond ETF (TLT). The Permanent Portfolio calls for 30-year bonds, but historic figures are only available for 20-year bonds.

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TICKERS_VECTOR As Variant
Dim DATES_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim WEIGHTS_VECTOR As Variant
Dim LOWER_VECTOR As Variant
Dim UPPER_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TICKERS_VECTOR = TICKERS_RNG
If UBound(TICKERS_VECTOR, 2) = 1 Then
    TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If
NCOLUMNS = UBound(TICKERS_VECTOR, 2)

DATES_VECTOR = DATES_RNG
If UBound(DATES_VECTOR, 1) = 1 Then
    DATES_VECTOR = MATRIX_TRANSPOSE_FUNC(DATES_VECTOR)
End If
NROWS = UBound(DATES_VECTOR, 1)
DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) <> NROWS Then: GoTo ERROR_LABEL
If UBound(DATA_MATRIX, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL

If IsArray(WEIGHTS_RNG) Then
    WEIGHTS_VECTOR = WEIGHTS_RNG
    If UBound(WEIGHTS_VECTOR, 2) = 1 Then
        WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
    End If
Else
    ReDim WEIGHTS_VECTOR(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS: WEIGHTS_VECTOR(1, j) = 1 / NCOLUMNS: Next j
End If
If NCOLUMNS <> UBound(WEIGHTS_VECTOR, 2) Then: GoTo ERROR_LABEL
If IsArray(LOWER_RNG) Then
    LOWER_VECTOR = LOWER_RNG
    If UBound(LOWER_VECTOR, 2) = 1 Then
        LOWER_VECTOR = MATRIX_TRANSPOSE_FUNC(LOWER_VECTOR)
    End If
Else
    ReDim LOWER_VECTOR(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS: LOWER_VECTOR(1, j) = LOWER_RNG: Next j
End If
If NCOLUMNS <> UBound(LOWER_VECTOR, 2) Then: GoTo ERROR_LABEL
If IsArray(UPPER_RNG) Then
    UPPER_VECTOR = UPPER_RNG
    If UBound(UPPER_VECTOR, 2) = 1 Then
        UPPER_VECTOR = MATRIX_TRANSPOSE_FUNC(UPPER_VECTOR)
    End If
Else
    ReDim UPPER_VECTOR(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS: UPPER_VECTOR(1, j) = UPPER_RNG: Next j
End If
If NCOLUMNS <> UBound(UPPER_VECTOR, 2) Then: GoTo ERROR_LABEL

'---------------------------------------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS + 1, 1 To 31)
'---------------------------------------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "RETURNS"
TEMP_MATRIX(1, 1) = "DATES/WEIGHTS"
TEMP_MATRIX(1, NCOLUMNS + 2) = 0
TEMP_MATRIX(1, NCOLUMNS * 2 + 3) = 0
TEMP_MATRIX(1, NCOLUMNS * 3 + 4) = 0
For j = 1 To NCOLUMNS
    TEMP_MATRIX(0, j + 1) = "HISTORIC DATA FOR % GAIN: " & TICKERS_VECTOR(1, j)
    TEMP_MATRIX(1, j + 1) = WEIGHTS_VECTOR(1, j)
    TEMP_MATRIX(1, NCOLUMNS + 2) = TEMP_MATRIX(1, NCOLUMNS + 2) + TEMP_MATRIX(1, j + 1)
    
    TEMP_MATRIX(0, NCOLUMNS + j + 2) = "HISTORIC REBALANCE: " & TICKERS_VECTOR(1, j)
    TEMP_MATRIX(1, NCOLUMNS + j + 2) = INITIAL_CASH_VAL * WEIGHTS_VECTOR(1, j)
    TEMP_MATRIX(1, NCOLUMNS * 2 + 3) = TEMP_MATRIX(1, NCOLUMNS * 2 + 3) + TEMP_MATRIX(1, NCOLUMNS + j + 2)
    
    TEMP_MATRIX(0, NCOLUMNS * 2 + j + 3) = "DYNAMIC REBALANCE: " & TICKERS_VECTOR(1, j)
    TEMP_MATRIX(1, NCOLUMNS * 2 + j + 3) = INITIAL_CASH_VAL * WEIGHTS_VECTOR(1, j) '/ NCOLUMNS
    TEMP_MATRIX(1, NCOLUMNS * 3 + 4) = TEMP_MATRIX(1, NCOLUMNS * 3 + 4) + TEMP_MATRIX(1, NCOLUMNS * 2 + j + 3)
Next j

TEMP_MATRIX(0, 1 + NCOLUMNS + 1) = "HISTORIC PORTFOLIO TOTAL RETURN"
TEMP_MATRIX(0, NCOLUMNS * 2 + 3) = "HISTORIC PORTFOLIO TOTAL GROWTH"
TEMP_MATRIX(0, NCOLUMNS * 3 + 4) = "DYNAMIC PORTFOLIO TOTAL GROWTH"

TEMP_MATRIX(1, NCOLUMNS * 4 + 5) = False
For j = 1 To NCOLUMNS
    TEMP_MATRIX(0, NCOLUMNS * 3 + j + 4) = "PROPORTIONS: " & TICKERS_VECTOR(1, j)
    TEMP_MATRIX(1, NCOLUMNS * 3 + j + 4) = TEMP_MATRIX(1, NCOLUMNS * 2 + j + 3) / TEMP_MATRIX(1, NCOLUMNS * 3 + 4)
    If TEMP_MATRIX(1, NCOLUMNS * 3 + j + 4) < LOWER_VECTOR(1, j) Or TEMP_MATRIX(1, NCOLUMNS * 3 + j + 4) > UPPER_VECTOR(1, j) Then: TEMP_MATRIX(1, NCOLUMNS * 4 + 5) = True
Next j
For j = 1 To NCOLUMNS
    TEMP_MATRIX(0, NCOLUMNS * 4 + j + 5) = "NEW ALLOCATION: " & TICKERS_VECTOR(1, j)
    TEMP_MATRIX(1, NCOLUMNS * 4 + j + 5) = IIf(TEMP_MATRIX(1, NCOLUMNS * 4 + 5) = True, TEMP_MATRIX(1, NCOLUMNS * 3 + 4) * WEIGHTS_VECTOR(1, j), TEMP_MATRIX(1, NCOLUMNS * 2 + j + 3))
    '/ NCOLUMNS
Next j

TEMP_MATRIX(0, NCOLUMNS * 4 + 5) = "DYNAMIC REBALANCE"

TEMP_MATRIX(0, NCOLUMNS * 5 + 1 + 5) = "CHANGE FACTORS: HISTORIC"
TEMP_MATRIX(0, NCOLUMNS * 5 + 2 + 5) = "CHANGE FACTORS: DYNAMIC"
TEMP_MATRIX(0, NCOLUMNS * 5 + 3 + 5) = "CHANGE %: HISTORIC"
TEMP_MATRIX(0, NCOLUMNS * 5 + 4 + 5) = "CHANGE %: DYNAMIC"
TEMP_MATRIX(0, NCOLUMNS * 5 + 5 + 5) = "COMPOSITE: HISTORIC"
TEMP_MATRIX(0, NCOLUMNS * 5 + 6 + 5) = "COMPOSITE: DYNAMIC"
For j = 1 To 6: TEMP_MATRIX(1, NCOLUMNS * 5 + j + 5) = "": Next j

'---------------------------------------------------------------------------------------------------------------
For i = 1 To NROWS
    TEMP_MATRIX(i + 1, 1) = DATES_VECTOR(i, 1)
    TEMP_MATRIX(i + 1, NCOLUMNS + 2) = 0
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(i + 1, j + 1) = DATA_MATRIX(i, j)
        TEMP_MATRIX(i + 1, NCOLUMNS + 2) = TEMP_MATRIX(i + 1, NCOLUMNS + 2) + TEMP_MATRIX(i + 1, j + 1) * WEIGHTS_VECTOR(1, j)
    Next j
    TEMP_MATRIX(i + 1, NCOLUMNS + 2) = TEMP_MATRIX(i + 1, NCOLUMNS + 2)
    TEMP_MATRIX(i + 1, NCOLUMNS * 2 + 3) = TEMP_MATRIX(i, NCOLUMNS * 2 + 3) * (1 + (TEMP_MATRIX(i + 1, NCOLUMNS + 2) / FACTOR_VAL))
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(i + 1, NCOLUMNS + j + 2) = (TEMP_MATRIX(i, NCOLUMNS * 2 + 3) * TEMP_MATRIX(1, j + 1)) * (1 + (TEMP_MATRIX(i + 1, j + 1) / FACTOR_VAL))
        TEMP_MATRIX(i + 1, NCOLUMNS * 2 + j + 3) = TEMP_MATRIX(i, NCOLUMNS * 4 + j + 5) * (1 + (TEMP_MATRIX(i + 1, j + 1) / FACTOR_VAL))
        TEMP_MATRIX(i + 1, NCOLUMNS * 3 + 4) = TEMP_MATRIX(i + 1, NCOLUMNS * 3 + 4) + TEMP_MATRIX(i + 1, NCOLUMNS * 2 + j + 3)
    Next j
    TEMP_MATRIX(i + 1, NCOLUMNS * 4 + 5) = False
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(i + 1, NCOLUMNS * 3 + j + 4) = TEMP_MATRIX(i + 1, NCOLUMNS * 2 + j + 3) / TEMP_MATRIX(i + 1, NCOLUMNS * 3 + 4)
        If TEMP_MATRIX(i + 1, NCOLUMNS * 3 + j + 4) < LOWER_VECTOR(1, j) Or TEMP_MATRIX(i + 1, NCOLUMNS * 3 + j + 4) > UPPER_VECTOR(1, j) Then: TEMP_MATRIX(i + 1, NCOLUMNS * 4 + 5) = True
    Next j
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(i + 1, NCOLUMNS * 4 + j + 5) = IIf(TEMP_MATRIX(i + 1, NCOLUMNS * 4 + 5) = True, TEMP_MATRIX(i + 1, NCOLUMNS * 3 + 4) * WEIGHTS_VECTOR(1, j), TEMP_MATRIX(i + 1, NCOLUMNS * 2 + j + 3))
        '/ NCOLUMNS
    Next j
    
    TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 1 + 5) = TEMP_MATRIX(i + 1, NCOLUMNS * 2 + 3) / TEMP_MATRIX(i, NCOLUMNS * 2 + 3)
    TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 2 + 5) = TEMP_MATRIX(i + 1, NCOLUMNS * 3 + 4) / TEMP_MATRIX(i, NCOLUMNS * 3 + 4)
    
    TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 3 + 5) = (TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 1 + 5) - 1) * 100
    TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 4 + 5) = (TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 2 + 5) - 1) * 100
    
    If i > 1 Then
        TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 5 + 5) = TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 1 + 5) * TEMP_MATRIX(i, NCOLUMNS * 5 + 5 + 5)
        TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 6 + 5) = TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 2 + 5) * TEMP_MATRIX(i, NCOLUMNS * 5 + 6 + 5)
    Else
        TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 5 + 5) = TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 1 + 5)
        TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 6 + 5) = TEMP_MATRIX(i + 1, NCOLUMNS * 5 + 2 + 5)
    End If
Next i
'---------------------------------------------------------------------------------------

PORT_RETURNS_REBALANCE2_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_RETURNS_REBALANCE2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_OPTIMAL_REBALANCE_FUNC

'DESCRIPTION   : Optimal level of re-balancing: A manager has developed
'a market neutral program with the inputed expected performance
'characteristics, in an optimal environment.

'LIBRARY       : PORTFOLIO
'GROUP         : REBALANCE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_OPTIMAL_REBALANCE_FUNC(ByVal EXPECTED_ALPHA As Double, _
ByVal EXPECTED_SIGMA As Double, _
Optional ByVal FIXED_EXPENSE As Double = 0.00002, _
Optional ByVal PROPORTIONAL_EXPENSE As Double = 0.001, _
Optional ByVal PORTFOLIO_DECAY As Double = 0.05, _
Optional ByVal NO_PERIODS As Long = 20, _
Optional ByVal DAYS_PER_PERIOD As Long = 250, _
Optional ByVal OUTPUT As Integer = 0)

'EXPECTED_ALPHA: Expected Alpha Return Component

'EXPECTED_SIGMA: Expected Program Volatility

'FIXED & PROPORTIONAL: Trading expenses consist of a fixed cost
'and a cost proportional to the trade size

'PORTFOLIO_DECAY: Decay Factor. Assume that the decay occurs at the
'end of the day. Assume returns are earn't on a straight-line fashion.

Dim i As Long
Dim MIN_VAL As Double
Dim RETURN_VAL As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

RETURN_VAL = (1 + EXPECTED_ALPHA) ^ (1 / DAYS_PER_PERIOD) - 1 'Daily Return

ReDim TEMP_MATRIX(0 To NO_PERIODS, 1 To 11)

TEMP_MATRIX(0, 1) = "PERIOD [DAY]"
TEMP_MATRIX(0, 2) = "EXPOSURE (PROGRAM) - START"
TEMP_MATRIX(0, 3) = "EXPOSURE (NOISE) - START"
TEMP_MATRIX(0, 4) = "E (EXPECTED ACTIVE RETURN)"
TEMP_MATRIX(0, 5) = "COST OF MISMATCH"
TEMP_MATRIX(0, 6) = "CUMMULATIVE COST OF MISMATCH"
TEMP_MATRIX(0, 7) = "EXPOSURE (NOISE) - END"
TEMP_MATRIX(0, 8) = "COST OF CORRECTING MISMATCH"
TEMP_MATRIX(0, 9) = "ANNUALISED COST OF MISMATCH"
TEMP_MATRIX(0, 10) = "TRADING COSTS"
TEMP_MATRIX(0, 11) = "TOTAL ANNUALISED COSTS"

i = 1
TEMP_MATRIX(i, 1) = i
TEMP_MATRIX(i, 2) = 1
TEMP_MATRIX(i, 3) = 1 - TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) * RETURN_VAL
TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3) * RETURN_VAL
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5)
TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 2) * PORTFOLIO_DECAY
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) * PROPORTIONAL_EXPENSE + FIXED_EXPENSE
TEMP_MATRIX(i, 9) = (1 + TEMP_MATRIX(i, 6)) ^ (DAYS_PER_PERIOD / TEMP_MATRIX(i, 1)) - 1
TEMP_MATRIX(i, 10) = (1 + TEMP_MATRIX(i, 8)) ^ (DAYS_PER_PERIOD / TEMP_MATRIX(i, 1)) - 1
TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 9) + TEMP_MATRIX(i, 10)

MIN_VAL = TEMP_MATRIX(1, 11)
For i = 2 To NO_PERIODS
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i - 1, 2) * (1 - PORTFOLIO_DECAY)
    TEMP_MATRIX(i, 3) = 1 - TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) * RETURN_VAL
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3) * RETURN_VAL
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 6) + TEMP_MATRIX(i, 5)
    TEMP_MATRIX(i, 7) = 1 - (TEMP_MATRIX(i, 2) * (1 - PORTFOLIO_DECAY))
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) * PROPORTIONAL_EXPENSE + FIXED_EXPENSE
    TEMP_MATRIX(i, 9) = (1 + TEMP_MATRIX(i, 6)) ^ (DAYS_PER_PERIOD / TEMP_MATRIX(i, 1)) - 1
    TEMP_MATRIX(i, 10) = (1 + TEMP_MATRIX(i, 8)) ^ (DAYS_PER_PERIOD / TEMP_MATRIX(i, 1)) - 1
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 9) + TEMP_MATRIX(i, 10)
    If TEMP_MATRIX(i, 11) < MIN_VAL Then: MIN_VAL = TEMP_MATRIX(i, 11)
Next i

'---------------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------------
Case 0 'Optimal level of re-balancing
'---------------------------------------------------------------------------------------
    PORT_OPTIMAL_REBALANCE_FUNC = TEMP_MATRIX
'---------------------------------------------------------------------------------------
Case 1 'Annual active expected return now
'---------------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 5, 1 To 2)
    TEMP_VECTOR(1, 1) = "EXPECTED ALPHA RETURN COMPONENT"
    TEMP_VECTOR(1, 2) = EXPECTED_ALPHA
        
    TEMP_VECTOR(2, 1) = "TRADING / DECAY COSTS"
    TEMP_VECTOR(2, 2) = MIN_VAL
        
    TEMP_VECTOR(3, 1) = "ACTIVE RETURN"
    TEMP_VECTOR(3, 2) = TEMP_VECTOR(1, 2) - TEMP_VECTOR(2, 2)
        
    TEMP_VECTOR(4, 1) = "RISK"
    TEMP_VECTOR(4, 2) = EXPECTED_SIGMA
        
    TEMP_VECTOR(5, 1) = "RETURN / RISK"
    TEMP_VECTOR(5, 2) = TEMP_VECTOR(3, 2) / TEMP_VECTOR(4, 2)
    
    PORT_OPTIMAL_REBALANCE_FUNC = TEMP_VECTOR
'---------------------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------------------
    PORT_OPTIMAL_REBALANCE_FUNC = Array(TEMP_MATRIX, TEMP_VECTOR)
'---------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_OPTIMAL_REBALANCE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_REBALANCE_TOLERANCE_FUNC

'DESCRIPTION   : Rebalancing a portfolio to an initial allocation if weight
'drift has exceeded a certain threshold.

'LIBRARY       : PORTFOLIO
'GROUP         : WEIGHTS_REBALANCE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function RNG_PORT_REBALANCE_TOLERANCE_FUNC(ByVal NASSETS As Long, _
ByVal nPeriods As Long, _
ByRef DST_RNG As Excel.Range)

'DATA_RNG: First Column of DATA_RNG = Periods (e.g., dates)
'DATA_RNG: First Row Ticker Symbols
'INIT_WEIGHTS_RNG: Row Vector of Weights per Symbol

Dim i As Long
Dim j As Long

Dim REBAL_RNG As Excel.Range
Dim RETURNS_RNG As Excel.Range
Dim PERIODS_RNG As Excel.Range
Dim SYMBOLS_RNG As Excel.Range

Dim TEMP1_RNG As Excel.Range
Dim TEMP2_RNG As Excel.Range
Dim TEMP3_RNG As Excel.Range
Dim TEMP4_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_PORT_REBALANCE_TOLERANCE_FUNC = False

'----------------------------------------------------------------------------------------------
Set DST_RNG = DST_RNG.Offset(2, 0)
'----------------------------------------------------------------------------------------------

DST_RNG.Cells(0, 1).value = "Time (bop)"
For j = 1 To NASSETS
    With DST_RNG.Cells(0, j + 1)
        .value = "Asset " & CStr(j)
        .Font.ColorIndex = 3
    End With
    For i = 1 To nPeriods
        If j = 1 Then: DST_RNG.Cells(i, 1).value = i
        With DST_RNG.Cells(i, j + 1)
            .value = 0
            .Font.ColorIndex = 5
        End With
    Next i
Next j

Set SYMBOLS_RNG = _
    Range(DST_RNG.Cells(1, 2), DST_RNG.Cells(1, NASSETS + 1))
Set PERIODS_RNG = _
    Range(DST_RNG.Cells(2, 1), DST_RNG.Cells(nPeriods + 1, 1))
Set RETURNS_RNG = _
    Range(DST_RNG.Cells(2, 2), DST_RNG.Cells(nPeriods + 1, NASSETS + 1))

Set REBAL_RNG = DST_RNG.Columns(NASSETS * 2 + 3).Rows(1)
With REBAL_RNG
    .value = 0.05
    .Font.Bold = True
    .Font.ColorIndex = 3
    .AddComment ("Rebalancing Tolerance")
End With
'----------------------------------------------------------------------------------------------

Set DST_RNG = Range(DST_RNG, DST_RNG.Cells(nPeriods, NASSETS * 3 + 6))

With DST_RNG
        
    Set TEMP1_RNG = Range(.Columns(2), .Columns(2 + NASSETS - 1))
    Set TEMP2_RNG = Range(.Columns(2 + NASSETS), .Columns(NASSETS * 2 + 2 - 1))
    Set TEMP3_RNG = TEMP2_RNG.Rows(1)
    Set TEMP4_RNG = TEMP2_RNG.Offset(0, 5 + NASSETS - 1)
  
    TEMP2_RNG.Rows(0).FormulaArray = "=" & TEMP1_RNG.Rows(0).Address
    TEMP4_RNG.Rows(0).FormulaArray = "=" & TEMP1_RNG.Rows(0).Address
    
    
    TEMP1_RNG.Rows(-1).Cells(1).value = "Constituent Period Returns"
    TEMP1_RNG.Rows(-1).Cells(1).Font.Bold = True
    
    TEMP1_RNG.Rows(0).Cells(1, 0).value = "Period"
    TEMP1_RNG.Rows(0).Cells(1, 0).Font.Bold = True

    TEMP2_RNG.Rows(-1).Cells(1).value = "Constituent Weights Beginning of Period"
    TEMP2_RNG.Rows(-1).Cells(1).Font.Bold = True
    
    TEMP4_RNG.Rows(-1).Cells(1).value = "Constituent Weights Beginning of Period"
    TEMP4_RNG.Rows(-1).Cells(1).Font.Bold = True


    For i = 1 To nPeriods
        If i = 1 Then 'Initial Weights Range
            TEMP2_RNG.Rows(1).value = 0
            TEMP2_RNG.Rows(1).Font.ColorIndex = 5
            TEMP4_RNG.Rows(1).value = 0
            TEMP4_RNG.Rows(1).Font.ColorIndex = 5
        Else
            TEMP2_RNG.Rows(i).FormulaArray = _
                "=IF(OR(ABS(" & TEMP2_RNG.Rows(i - 1).Address & _
                "*(1+" & TEMP1_RNG.Rows(i - 1).Address & ")/(1+MMULT(" & _
                TEMP2_RNG.Rows(i - 1).Address & ",TRANSPOSE(" & _
                TEMP1_RNG.Rows(i - 1).Address & _
                ")))-" & TEMP3_RNG.Address & ")>" & _
                REBAL_RNG.Address & ")," & _
                TEMP3_RNG.Address & "," & _
                TEMP2_RNG.Rows(i - 1).Address & _
                "*(1+" & TEMP1_RNG.Rows(i - 1).Address & _
                ")/(1+MMULT(" & TEMP2_RNG.Rows(i - 1).Address & _
                ",TRANSPOSE(" & TEMP1_RNG.Rows(i - 1).Address & "))))"
                
            TEMP4_RNG.Rows(i).FormulaArray = _
                "=" & TEMP4_RNG.Rows(i - 1).Address & _
                "*(1+" & TEMP1_RNG.Rows(i - 1).Address & _
                ")/(1+MMULT(" & TEMP4_RNG.Rows(i - 1).Address & _
                ",TRANSPOSE(" & TEMP1_RNG.Rows(i - 1).Address & ")))"

            .Columns(NASSETS * 2 + 3).Rows(i).FormulaArray = _
                "=OR(ABS(" & TEMP2_RNG.Rows(i - 1).Address & _
                "*(1+" & TEMP1_RNG.Rows(i - 1).Address & _
                ")/(1+MMULT(" & TEMP2_RNG.Rows(i - 1).Address & _
                ",TRANSPOSE(" & TEMP1_RNG.Rows(i - 1).Address & _
                ")))-" & TEMP3_RNG.Address & ")>" & _
                REBAL_RNG.Address & ")"


        End If


        .Columns(NASSETS * 2 + 2).Rows(i).formula = _
            "=SUM(" & TEMP2_RNG.Rows(i).Address & ")"
    
        .Columns(NASSETS * 2 + 4).Rows(i).FormulaArray = _
            "=MMULT(" & TEMP2_RNG.Rows(i).Address & _
            ",TRANSPOSE(" & TEMP1_RNG.Rows(i).Address & "))"
            
        .Columns(NASSETS * 2 + 5).Rows(i).FormulaArray = _
            "=MMULT(" & TEMP3_RNG.Address & ",TRANSPOSE(" & _
            TEMP1_RNG.Rows(i).Address & "))"
    
        .Columns(NASSETS * 3 + 6).Rows(i).FormulaArray = _
            "=MMULT(" & TEMP4_RNG.Rows(i).Address & ",TRANSPOSE(" & _
            TEMP1_RNG.Rows(i).Address & "))"
    
    Next i

    .Columns(NASSETS * 2 + 2).Rows(0).value = "Total"
    .Columns(NASSETS * 2 + 2).Rows(0).Font.Bold = True

    .Columns(NASSETS * 2 + 3).Rows(0).value = "Rebalanced"
    .Columns(NASSETS * 2 + 3).Rows(0).Font.Bold = True

    .Columns(NASSETS * 2 + 4).Rows(0).value = "Threshold Rebalance Portfolio"
    .Columns(NASSETS * 2 + 4).Rows(0).Font.Bold = True

    .Columns(NASSETS * 2 + 5).Rows(0).value = "Always Rebalance Portfolio"
    .Columns(NASSETS * 2 + 5).Rows(0).Font.Bold = True

    .Columns(NASSETS * 3 + 6).Rows(0).value = "Buy & Hold Portfolio"
    .Columns(NASSETS * 3 + 6).Rows(0).Font.Bold = True

End With

RNG_PORT_REBALANCE_TOLERANCE_FUNC = True

Exit Function
ERROR_LABEL:
RNG_PORT_REBALANCE_TOLERANCE_FUNC = False
End Function

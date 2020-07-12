Attribute VB_Name = "FINAN_ASSET_SIMUL_WEALTH_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSETS_REAL_WEALTH_SIMULATION_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal INFLATION_RATE As Double = 0.03, _
Optional ByVal WITHDRAWAL_RATE As Double = 0.05, _
Optional ByVal NO_PERIODS As Long = 252 * 10, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal COUNT_BASIS As Double = 252)

'NO_PERIODS --> Forward
'nLOOPS --> number of NO_PERIODS simulations

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long ' counts number of portfolios that don't survive nLOOPS periods
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim PORT_VAL As Double
Dim RETURN_VAL As Double
Dim RISK_VAL As Double
Dim INFLATION_VAL As Double
Dim WITHDRAWAL_VAL As Double

Dim TICKER_STR As String

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NCOLUMNS = UBound(TICKERS_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 2)
TEMP_MATRIX(0, 1) = "SYMBOL"
TEMP_MATRIX(0, 2) = "RISK"

For k = 1 To NCOLUMNS
    TICKER_STR = TICKERS_VECTOR(k, 1)
    DATA_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, _
                  END_DATE, "DAILY", "A", False, False, True)
    DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 0)
    NROWS = UBound(DATA_VECTOR, 1)
    l = 0                                     ' counts number of portfolios that don't survive 40 years
    INFLATION_VAL = 1 + (INFLATION_RATE / COUNT_BASIS) '0.03 / 12  ' Assume 3% annual inflation
    For i = 1 To nLOOPS                            ' start k simulations
    
        WITHDRAWAL_VAL = WITHDRAWAL_RATE / COUNT_BASIS '0.06 / 12  ' Assume a 6% annual withdrawal rate
        PORT_VAL = 1  ' start with $1.00 portfolio
        Randomize
        For j = 1 To NO_PERIODS ' no periods forward
            WITHDRAWAL_VAL = WITHDRAWAL_VAL * INFLATION_VAL          ' increase withdrawal
            h = (NROWS - 1) * Rnd + 1
            RETURN_VAL = DATA_VECTOR(h, 1) ' pick a monthly return
            PORT_VAL = PORT_VAL * (1 + RETURN_VAL) - WITHDRAWAL_VAL  ' update portfolio
            If PORT_VAL < 0 Then
                l = l + 1                   ' count number of dead portfolios
                PORT_VAL = 0                        ' portfolio dies !!
                j = NO_PERIODS + 1 ' stop this simulation
            End If
        Next j
    Next i
    RISK_VAL = l / nLOOPS                    ' what number fail to survive?
    TEMP_MATRIX(k, 1) = TICKER_STR
    TEMP_MATRIX(k, 2) = RISK_VAL
Next k

ASSETS_REAL_WEALTH_SIMULATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_REAL_WEALTH_SIMULATION_FUNC = Err.number
End Function

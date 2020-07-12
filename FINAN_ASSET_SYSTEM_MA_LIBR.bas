Attribute VB_Name = "FINAN_ASSET_SYSTEM_MA_LIBR"

'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------
Private PUB_INITIAL_CASH As Double
Private PUB_DATA_MATRIX As Variant
'compare a Buy&Hold scheme with a x-day Moving Average scheme.
'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------

Function ASSET_MA_SYSTEM_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Double = 200, _
Optional ByVal INITIAL_CASH As Double = 10000, _
Optional ByVal BUY_ABOVE As Double = 0.01, _
Optional ByVal SELL_BELOW As Double = 0.02, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NO_DAYS As Long

Dim NROWS As Long

Dim XCAGR_VAL As Double
Dim SCAGR_VAL As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim TEMP_VAL As Double

Dim UNITS_VAL As Long 'double

Dim TEMP_SUM As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DA", False, False, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"

'----------------------------------------------------------------------------
TEMP_MATRIX(0, 6) = "CASH BALANCE"
TEMP_MATRIX(0, 7) = "SHARES TRADED"
'----------------------------------------------------------------------------

i = 1
TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
TEMP_SUM = TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 3) = TEMP_SUM / i
TEMP_MATRIX(i, 4) = ""
TEMP_MATRIX(i, 6) = ""
TEMP_MATRIX(i, 7) = ""
TEMP_MATRIX(i, 8) = ""
TEMP_MATRIX(i, 9) = ""

For i = 2 To MA_PERIOD - 2
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = TEMP_SUM / i
    For j = 4 To 9: TEMP_MATRIX(i, j) = "": Next j
Next i

TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 3) = TEMP_SUM / i

TEMP_VAL = TEMP_MATRIX(i, 2) / TEMP_MATRIX(i, 3) - 1
MIN_VAL = TEMP_VAL
MAX_VAL = TEMP_VAL

TEMP_MATRIX(i, 4) = ""
TEMP_MATRIX(i, 5) = ""
TEMP_MATRIX(i, 6) = INITIAL_CASH
TEMP_MATRIX(i, 7) = 0
TEMP_MATRIX(i, 8) = INITIAL_CASH
TEMP_MATRIX(i, 9) = INITIAL_CASH

MEAN_VAL = 0
'-----------------------------------------------------------------------------------------------------------
For i = MA_PERIOD To NROWS
'-----------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 2)
                    
    TEMP_MATRIX(i, 3) = TEMP_SUM / MA_PERIOD
    TEMP_VAL = TEMP_MATRIX(i, 2) / TEMP_MATRIX(i, 3) - 1
    If TEMP_VAL < MIN_VAL Then: MIN_VAL = TEMP_VAL
    If TEMP_VAL > MAX_VAL Then: MAX_VAL = TEMP_VAL
    
    k = i - MA_PERIOD + 1
    TEMP_SUM = TEMP_SUM - TEMP_MATRIX(k, 2)
    
    If TEMP_MATRIX(i, 2) > (1 + BUY_ABOVE) * TEMP_MATRIX(i, 3) Then
        TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) 'buy
        TEMP_MATRIX(i, 5) = ""
        TEMP_MATRIX(i, 6) = 0
        If TEMP_MATRIX(i - 1, 6) > 0 Then
            UNITS_VAL = TEMP_MATRIX(i - 1, 6) / TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 7) = UNITS_VAL
        Else
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i - 1, 7)
        End If
    Else
        If TEMP_MATRIX(i, 2) < (1 - SELL_BELOW) * TEMP_MATRIX(i, 3) Then
            TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 2) 'sell
            TEMP_MATRIX(i, 4) = ""
            If TEMP_MATRIX(i - 1, 7) > 0 Then
                TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 7) * TEMP_MATRIX(i, 2)
            Else
                TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 6)
            End If
            TEMP_MATRIX(i, 7) = 0
        Else
            TEMP_MATRIX(i, 4) = ""
            TEMP_MATRIX(i, 5) = ""
            TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 6)
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i - 1, 7)
        End If
    End If
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 6) + TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 2)
    MEAN_VAL = MEAN_VAL + TEMP_MATRIX(i, 8) / TEMP_MATRIX(i - 1, 8) - 1
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i - 1, 9) * (1 + TEMP_MATRIX(i, 2) / TEMP_MATRIX(i - 1, 2) - 1)
'-----------------------------------------------------------------------------------------------------------
Next i
'-----------------------------------------------------------------------------------------------------------

i = MA_PERIOD - 1
NO_DAYS = TEMP_MATRIX(NROWS, 1) - TEMP_MATRIX(i, 1)
If NO_DAYS = 0 Then: NO_DAYS = 1
SCAGR_VAL = (TEMP_MATRIX(NROWS, 8) / TEMP_MATRIX(i, 8)) ^ (365 / NO_DAYS) - 1
XCAGR_VAL = (TEMP_MATRIX(NROWS, 9) / TEMP_MATRIX(i, 9)) ^ (365 / NO_DAYS) - 1


TEMP_MATRIX(0, 3) = MA_PERIOD & "-MA PRICE = " & Format(TEMP_MATRIX(NROWS, 3), "0.00")

TEMP_MATRIX(0, 4) = "BUY SIGNAL: MAX DEV = " & Format(MAX_VAL, "0.00%")
TEMP_MATRIX(0, 5) = "SELL SIGNAL: MIN DEV = " & Format(MIN_VAL, "0.00%")

TEMP_MATRIX(0, 8) = "SYSTEM BALANCE: CAGR = " & Format(SCAGR_VAL, "0.00%")
TEMP_MATRIX(0, 9) = "BUY HOLD BALANCE: CAGR = " & Format(XCAGR_VAL, "0.00%")

If OUTPUT = 0 Then
    ASSET_MA_SYSTEM_FUNC = TEMP_MATRIX
Else

    MEAN_VAL = MEAN_VAL / (NROWS - MA_PERIOD + 1)
    SIGMA_VAL = 0
    For j = MA_PERIOD To NROWS
        SIGMA_VAL = SIGMA_VAL + ((TEMP_MATRIX(j, 8) / TEMP_MATRIX(j - 1, 8) - 1) - MEAN_VAL) ^ 2
    Next j
    SIGMA_VAL = (SIGMA_VAL / (NROWS - MA_PERIOD + 1)) ^ 0.5
    If OUTPUT = 1 Then
        ASSET_MA_SYSTEM_FUNC = MEAN_VAL / SIGMA_VAL
    ElseIf OUTPUT = 2 Then
        ASSET_MA_SYSTEM_FUNC = Array(MEAN_VAL, SIGMA_VAL, MEAN_VAL / SIGMA_VAL, SCAGR_VAL, XCAGR_VAL)
    Else 'If OUTPUT = 3 Then
        ASSET_MA_SYSTEM_FUNC = Array(SCAGR_VAL, XCAGR_VAL)
    End If
End If

Exit Function
ERROR_LABEL:
ASSET_MA_SYSTEM_FUNC = "--"
End Function

Function ASSETS_MA_SYSTEM_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Double = 200, _
Optional ByVal INITIAL_CASH As Double = 10000, _
Optional ByVal BUY_ABOVE As Double = 0.01, _
Optional ByVal SELL_BELOW As Double = 0.02)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim NO_DAYS As Long
Dim UNITS_VAL As Long

Dim SCAGR_VAL As Double
Dim XCAGR_VAL As Double

Dim DATE0_VAL As Date
Dim DATE1_VAL As Date
Dim DATE2_VAL As Date

Dim PRICE0_VAL As Double
Dim PRICE1_VAL As Double
Dim PRICE2_VAL As Double

Dim MA_VAL As Double
Dim SIGNAL_STR As String
Dim TICKER_STR As String

Dim CASH1_VAL As Double
Dim CASH2_VAL As Double

Dim SHARES1_VAL As Long
Dim SHARES2_VAL As Long

Dim BUY_HOLD0_VAL As Double
Dim BUY_HOLD1_VAL As Double
Dim BUY_HOLD2_VAL As Double

Dim SYSTEM0_VAL As Double
Dim SYSTEM1_VAL As Double
Dim SYSTEM2_VAL As Double

Dim TEMP_SUM As Double
Dim TEMP_RETURN As Double
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim DATA_ARR() As Double

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

ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 11)
TEMP_MATRIX(0, 1) = "TICKER"
TEMP_MATRIX(0, 2) = "STARTING PERIOD"
TEMP_MATRIX(0, 3) = "ENDING PERIOD"
TEMP_MATRIX(0, 4) = "CLOSING PRICE"
TEMP_MATRIX(0, 5) = MA_PERIOD & "-MA PRICE"
TEMP_MATRIX(0, 6) = "SYSTEM CAGR"
TEMP_MATRIX(0, 7) = "BUY HOLD CAGR"
TEMP_MATRIX(0, 8) = "CURRENT SIGNAL"
TEMP_MATRIX(0, 9) = "SYSTEM DAILY AVG RETURN"
TEMP_MATRIX(0, 10) = "SYSTEM DAILY VOLATILITY"
TEMP_MATRIX(0, 11) = "SYSTEM SHARPE"

'-------------------------------------------------------------------------------------------------------
For j = 1 To NCOLUMNS
'-------------------------------------------------------------------------------------------------------

    TICKER_STR = TICKERS_VECTOR(j, 1)
    TEMP_MATRIX(j, 1) = TICKER_STR
    
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DA", False, False, True)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    NROWS = UBound(DATA_MATRIX, 1)
    
    TEMP_SUM = 0
    For i = 1 To MA_PERIOD - 2
        PRICE0_VAL = DATA_MATRIX(i, 2)
        TEMP_SUM = TEMP_SUM + PRICE0_VAL
        MA_VAL = TEMP_SUM / i
    Next i
    
    DATE0_VAL = DATA_MATRIX(i, 1)
    DATE1_VAL = DATE0_VAL
    
    PRICE0_VAL = DATA_MATRIX(i, 2)
    PRICE1_VAL = PRICE0_VAL
    
    TEMP_SUM = TEMP_SUM + PRICE1_VAL
    MA_VAL = TEMP_SUM / i
    
    CASH1_VAL = INITIAL_CASH
    SHARES1_VAL = 0
    BUY_HOLD0_VAL = INITIAL_CASH
    BUY_HOLD1_VAL = BUY_HOLD0_VAL
    
    SYSTEM0_VAL = INITIAL_CASH
    SYSTEM1_VAL = SYSTEM0_VAL
    
    MEAN_VAL = 0
    ReDim DATA_ARR(MA_PERIOD To NROWS)
    '-----------------------------------------------------------------------------------------------------------
    For i = MA_PERIOD To NROWS
    '-----------------------------------------------------------------------------------------------------------
        DATE2_VAL = DATA_MATRIX(i, 1)
        PRICE2_VAL = DATA_MATRIX(i, 2)
        TEMP_SUM = TEMP_SUM + PRICE2_VAL
        MA_VAL = TEMP_SUM / MA_PERIOD
        
        k = i - MA_PERIOD + 1
        TEMP_SUM = TEMP_SUM - DATA_MATRIX(k, 2)
        
        If PRICE2_VAL > (1 + BUY_ABOVE) * MA_VAL Then
            SIGNAL_STR = "BUY"
            CASH2_VAL = 0
            If CASH1_VAL > 0 Then
                UNITS_VAL = CASH1_VAL / PRICE2_VAL
                SHARES2_VAL = UNITS_VAL
            Else
                SHARES2_VAL = SHARES1_VAL
            End If
        Else
            If PRICE2_VAL < (1 - SELL_BELOW) * MA_VAL Then
                SIGNAL_STR = "SELL"
                If SHARES1_VAL > 0 Then
                    CASH2_VAL = SHARES1_VAL * PRICE2_VAL
                Else
                    CASH2_VAL = CASH1_VAL
                End If
                SHARES2_VAL = 0
            Else
                SIGNAL_STR = ""
                CASH2_VAL = CASH1_VAL
                SHARES2_VAL = SHARES1_VAL
            End If
        End If
        SYSTEM2_VAL = CASH2_VAL + SHARES2_VAL * PRICE2_VAL
        DATA_ARR(i) = SYSTEM2_VAL / SYSTEM1_VAL - 1
        MEAN_VAL = MEAN_VAL + DATA_ARR(i)
        
        TEMP_RETURN = PRICE2_VAL / PRICE1_VAL - 1
        BUY_HOLD2_VAL = BUY_HOLD1_VAL * (1 + TEMP_RETURN)
        
        DATE1_VAL = DATE2_VAL
        PRICE1_VAL = PRICE2_VAL
        
        CASH1_VAL = CASH2_VAL
        SHARES1_VAL = SHARES2_VAL
        
        SYSTEM1_VAL = SYSTEM2_VAL
        BUY_HOLD1_VAL = BUY_HOLD2_VAL
    Next i
    
    MEAN_VAL = MEAN_VAL / (NROWS - MA_PERIOD + 1)
    SIGMA_VAL = 0
    For i = MA_PERIOD To NROWS
        SIGMA_VAL = SIGMA_VAL + (DATA_ARR(i) - MEAN_VAL) ^ 2
    Next i
    SIGMA_VAL = (SIGMA_VAL / (NROWS - MA_PERIOD + 1)) ^ 0.5

    NO_DAYS = DATE1_VAL - DATE0_VAL
    If NO_DAYS = 0 Then: NO_DAYS = 1
    SCAGR_VAL = (SYSTEM1_VAL / SYSTEM0_VAL) ^ (365 / NO_DAYS) - 1
    XCAGR_VAL = (BUY_HOLD1_VAL / BUY_HOLD0_VAL) ^ (365 / NO_DAYS) - 1

    TEMP_MATRIX(j, 2) = DATE0_VAL
    TEMP_MATRIX(j, 3) = DATE1_VAL
    TEMP_MATRIX(j, 4) = PRICE1_VAL
    TEMP_MATRIX(j, 5) = MA_VAL
    TEMP_MATRIX(j, 6) = SCAGR_VAL
    TEMP_MATRIX(j, 7) = XCAGR_VAL
    TEMP_MATRIX(j, 8) = SIGNAL_STR
    TEMP_MATRIX(j, 9) = MEAN_VAL
    TEMP_MATRIX(j, 10) = SIGMA_VAL
    TEMP_MATRIX(j, 11) = MEAN_VAL / SIGMA_VAL

1983:
'-------------------------------------------------------------------------------------------------------
Next j
'-------------------------------------------------------------------------------------------------------

ASSETS_MA_SYSTEM_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_MA_SYSTEM_FUNC = Err.number
End Function


Function ASSET_MA_SYSTEM_OPTIMIZER_FUNC(ByRef PARAM_RNG As Variant, _
ByRef CONST_RNG As Variant, _
ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal INITIAL_CASH As Double = 1000)

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    PUB_DATA_MATRIX = TICKER_STR
Else
    PUB_DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    "d", "DA", False, False, True)
End If
PUB_INITIAL_CASH = INITIAL_CASH

'ASSET_MA_SYSTEM_OPTIMIZER_FUNC = _
    NELDER_MEAD_OPTIMIZATION_FRAME_FUNC("ASSET_MA_SYSTEM_OBJ_FUNC", _
    PARAM_RNG, CONST_RNG, False, 0, 10000, 0.000000000001)

ASSET_MA_SYSTEM_OPTIMIZER_FUNC = _
    PIKAIA_OPTIMIZATION_FUNC("ASSET_MA_SYSTEM_OBJ_FUNC", _
    CONST_RNG, False, , , , , , , , , , , , , , 0)

Exit Function
ERROR_LABEL:
ASSET_MA_SYSTEM_OPTIMIZER_FUNC = Err.number
End Function

Function ASSET_MA_SYSTEM_OBJ_FUNC(ByRef PARAM_VECTOR As Variant)

Dim THETA_VAL As Variant

On Error GoTo ERROR_LABEL

THETA_VAL = _
    ASSET_MA_SYSTEM_FUNC(PUB_DATA_MATRIX, , , _
    PARAM_VECTOR(1, 1), _
    PUB_INITIAL_CASH, _
    PARAM_VECTOR(2, 1), _
    PARAM_VECTOR(3, 1), 1)
    
If IsNumeric(THETA_VAL) = False Then: GoTo ERROR_LABEL

ASSET_MA_SYSTEM_OBJ_FUNC = THETA_VAL

Exit Function
ERROR_LABEL:
ASSET_MA_SYSTEM_OBJ_FUNC = 1 / 1E+100
End Function

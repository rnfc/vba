Attribute VB_Name = "FINAN_ASSET_SYSTEM_GCYCLE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_GCYCLE_VOLATILITY_SYSTEM_FUNC

'DESCRIPTION   : Identify trends in the stock market and to identify when
'they start and end: Upper and Lower Bollinger bands and
'the current price and price

'LIBRARY       : FINAN_ASSET
'GROUP         : SIGNAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_GCYCLE_VOLATILITY_SYSTEM_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal BUY_THRESHOLD As Double = 0.01, _
Optional ByVal SELL_THRESHOLD As Double = 0.01, _
Optional ByRef WINDOW_PERIODS As Long = 10, _
Optional ByVal MA_PERIOD As Double = 50, _
Optional ByRef SIGMA_OPT As Integer = 0, _
Optional ByVal SIGMA_FACT As Double = 1, _
Optional ByVal DIVISOR As Integer = 3)

'IF SIGMA_OPT = 0 Then: g-Cycle
'IF SIGMA_OPT = 1 Then: Standard Deviation

'BUY_THRESHOLD = in Percentage
'SELL_THRESHOLD = in Percentage

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim TEMP_VAL As Double
Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If SIGMA_OPT > 1 Then: SIGMA_OPT = 1
If SIGMA_OPT < 0 Then: SIGMA_OPT = 0

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 13)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "MEAN[PRICE]"
TEMP_MATRIX(0, 4) = "SIGMA[PRICE]"
TEMP_MATRIX(0, 5) = "UB"
TEMP_MATRIX(0, 6) = "LB"
TEMP_MATRIX(0, 7) = "GCYCLE"
TEMP_MATRIX(0, 8) = "VIGOR"

TEMP_MATRIX(0, 9) = "DATE2"
TEMP_MATRIX(0, 10) = "WIN2"
TEMP_MATRIX(0, 11) = "WIN3"

TEMP_MATRIX(0, 12) = "SELL SIGNAL"
TEMP_MATRIX(0, 13) = "BUY SIGNAL"

TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 2)
    If (i - 1) <= MA_PERIOD Then
        TEMP_VAL = TEMP_MATRIX(i, 2)
        TEMP_MATRIX(i, 3) = TEMP_SUM / i
        TEMP_MATRIX(i, 4) = 0
        For j = 1 To i
            TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 4) + (TEMP_MATRIX(j, 2) - TEMP_MATRIX(i, 3)) ^ 2
        Next j
        TEMP_MATRIX(i, 4) = (TEMP_MATRIX(i, 4) / i) ^ 0.5
    Else
        TEMP_VAL = TEMP_MATRIX(i - 1 - MA_PERIOD, 2)
        SROW = i - MA_PERIOD - 1
        TEMP_MATRIX(i, 3) = TEMP_SUM / (MA_PERIOD + 2)
        TEMP_MATRIX(i, 4) = 0
        For j = SROW To i
            TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 4) + (TEMP_MATRIX(j, 2) - TEMP_MATRIX(i, 3)) ^ 2
        Next j
        TEMP_MATRIX(i, 4) = (TEMP_MATRIX(i, 4) / (MA_PERIOD + 2)) ^ 0.5
        TEMP_SUM = TEMP_SUM - TEMP_MATRIX(SROW, 2)
    End If

    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3) + SIGMA_FACT * TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 3) - SIGMA_FACT * TEMP_MATRIX(i, 4)

    TEMP_MATRIX(i, 7) = Sqr((1 - 0.5 * SIGMA_OPT) * (TEMP_MATRIX(i, 5) ^ 2 + TEMP_MATRIX(i, 6) ^ 2) - (TEMP_MATRIX(i, 3) + SIGMA_OPT * (TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 3)) - SIGMA_OPT * TEMP_VAL) ^ 2 / DIVISOR)
    If (i - 1) <= MA_PERIOD Then
        TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(1, 7) - 1
        TEMP_MATRIX(i, 9) = DATA_MATRIX(i - 1 + WINDOW_PERIODS, 1)
        TEMP_MATRIX(i, 10) = DATA_MATRIX(i - 1 + WINDOW_PERIODS, 2)
    Else
        TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1 - MA_PERIOD, 7) - 1
        TEMP_MATRIX(i, 9) = ""
        TEMP_MATRIX(i, 10) = ""
    End If
Next i

TEMP_VAL = TEMP_MATRIX(MA_PERIOD + 1, 9)

For i = 1 To NROWS
    If (TEMP_MATRIX(i, 1) >= TEMP_MATRIX(1, 9)) And (TEMP_MATRIX(i, 1) <= TEMP_VAL) Then
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 2)
    Else
        TEMP_MATRIX(i, 11) = 0
    End If
    
    If TEMP_MATRIX(i, 8) >= SELL_THRESHOLD Then
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 2)
    Else
        TEMP_MATRIX(i, 12) = 0
    End If
    
    If TEMP_MATRIX(i, 8) <= BUY_THRESHOLD Then
        TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 2)
    Else
        TEMP_MATRIX(i, 13) = 0
    End If
Next i
ASSET_GCYCLE_VOLATILITY_SYSTEM_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_GCYCLE_VOLATILITY_SYSTEM_FUNC = Err.number
End Function

